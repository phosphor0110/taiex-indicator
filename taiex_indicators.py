#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
先行指標每日自動更新腳本
=========================
複製自「先行指標.xlsm」中 Power Query + 公式的邏輯，改用 Python 實作，
可放在雲端排程（例如 GitHub Actions）每天自動執行，不需要開電腦或用手機操作 Excel。

執行後會：
  1. 向證交所(TWSE)、期交所(TAIFEX)抓取當日公開資料
  2. 依照原始 Excel 檔的公式，計算出「先行指標」8 個欄位
  3. 把新的一列資料附加寫入 history.csv（若當天已存在則跳過，避免重複）

已實際用 web_fetch 驗證過的部分（欄位順序正確）：
  - futContractsDateExcel（三大法人期貨未平倉）
  - largeTraderFutQryTbl（大額交易人前五/十大留倉）
  - pcRatioExcel（選擇權 PCR）
  - callsAndPutsDateExcel（三大法人選擇權），但發現一個重要限制：見下方。

尚未實際驗證、風險較高的部分（收盤後請務必檢查結果是否合理）：
  - TWSE 成交量 / 三大法人買賣超（改用 rwd JSON API，格式已確認存在，
    但完整解析邏輯未逐欄核對）
  - futDailyMarketExcel（大台/小台 每日行情，用來加總未沖銷口數）
  - pd.read_html 在這些多層表頭(rowspan/colspan)的表格上是否能正確斷欄，
    沒有在真正的 Python 環境跑過，需要你第一次執行時人工核對數字。

重要限制（已證實）：
  - 期交所「三大法人-選擇權買賣權分計」的未平倉欄位是隔一個交易日才公布，
    收盤當天查詢會全部顯示「-」。所以「外(選)」這個欄位沒辦法在收盤當晚
    抓到，建議排程改成「隔天開盤前」執行，或把這欄獨立成晚一天的排程。

若抓取失敗，腳本會印出錯誤並中止該欄位，不會用假資料填補。
"""

import csv
import datetime as dt
import io
import os
import re
import sys

import pandas as pd
import requests

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}
TIMEOUT = 20
HISTORY_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), "history.csv")

COLUMNS = [
    "日期", "成交量", "外資", "投信", "自營", "外資大小",
    "前五大交易人留倉", "前十大交易人留倉", "選PCR", "外(選)",
    "韭菜指數", "未平倉口數",
]


# --------------------------------------------------------------------------
# 小工具
# --------------------------------------------------------------------------

def _to_number(text):
    """把 '76,720\n(76,720)' 這種格式清成數字，只取換行前、括號前的主數字。"""
    if text is None:
        return None
    text = str(text).split("\n")[0].split("(")[0]
    text = text.replace(",", "").replace("%", "").strip()
    if text in ("", "-", "－"):
        return 0.0
    try:
        return float(text)
    except ValueError:
        return None


def _get(url, params=None):
    resp = requests.get(url, headers=HEADERS, params=params, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp


# --------------------------------------------------------------------------
# 1. TWSE 大盤成交量
# --------------------------------------------------------------------------

def fetch_twse_volume():
    """證交所每日成交量統計，取『總計(1~15)』成交金額(元)。"""
    url = "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX"
    # TWSE 的 rwd JSON API 比原本的 HTML 頁面更穩定，直接查當日
    today = dt.date.today().strftime("%Y%m%d")
    resp = _get(url, params={"date": today, "type": "ALL", "response": "json"})
    data = resp.json()
    for table in data.get("tables", []):
        if "總計" in "".join(table.get("title", "")) or table.get("fields") and "成交金額" in "".join(table["fields"]):
            for row in table.get("data", []):
                if row and "總計" in row[0]:
                    amount = _to_number(row[1])
                    return amount  # 元
    raise RuntimeError("找不到 TWSE 總計成交金額，頁面格式可能已變更")


# --------------------------------------------------------------------------
# 2. TWSE 三大法人買賣超
# --------------------------------------------------------------------------

def fetch_twse_institutional():
    """回傳 dict：外資、投信、自營（自行+避險合計），單位：元。"""
    url = "https://www.twse.com.tw/rwd/zh/fund/BFI82U"
    today = dt.date.today().strftime("%Y%m%d")
    resp = _get(url, params={"dayDate": today, "type": "day", "response": "json"})
    data = resp.json()
    rows = data.get("data", [])
    result = {"自營自行": 0.0, "自營避險": 0.0, "投信": 0.0, "外資": 0.0}
    for row in rows:
        name, buy, sell, diff = row[0], row[1], row[2], row[3]
        diff_val = _to_number(diff)
        if "自行買賣" in name:
            result["自營自行"] = diff_val
        elif "避險" in name:
            result["自營避險"] = diff_val
        elif name.startswith("投信"):
            result["投信"] = diff_val
        elif "外資及陸資" in name and "自營商" not in name:
            result["外資"] = diff_val
    return {
        "外資": result["外資"],
        "投信": result["投信"],
        "自營": result["自營自行"] + result["自營避險"],
    }


# --------------------------------------------------------------------------
# 3. TAIFEX 三大法人期貨交易口數與契約金額（含未平倉）-> 期貨 分頁
# --------------------------------------------------------------------------

def fetch_taifex_futures_positions():
    """
    回傳 DataFrame，欄位：商品名稱, 身份別, 未平倉多方口數, 未平倉空方口數, 未平倉多空淨額口數
    對應原檔「期貨」分頁 J/L/N 欄。
    """
    url = "https://www.taifex.com.tw/cht/3/futContractsDateExcel"
    resp = _get(url)
    tables = pd.read_html(io.StringIO(resp.text))
    df = max(tables, key=len)  # 取列數最多的表格，通常就是主表
    df.columns = [str(c).strip() for c in range(df.shape[1])]
    records = []
    current_name = None
    for _, row in df.iterrows():
        vals = list(row)
        # 商品名稱欄若非空就更新目前商品
        name_cell = str(vals[0]).strip()
        if name_cell and name_cell not in ("nan", ""):
            current_name = name_cell
        role_cell = str(vals[1]).strip() if len(vals) > 1 else ""
        if role_cell in ("自營商", "投信", "外資"):
            try:
                # 實測欄位順序（已用 web_fetch 驗證過一次真實頁面）：
                # 0序號 1商品名稱 2身份別 3口數多方 4契約金額多方 5口數空方 6契約金額空方
                # 7口數多空淨額 8契約金額多空淨額 9口數未平倉多方 10契約金額未平倉多方
                # 11口數未平倉空方 12契約金額未平倉空方 13口數未平倉淨額 14契約金額未平倉淨額
                long_oi = _to_number(vals[9])    # 未平倉餘額-多方-口數
                short_oi = _to_number(vals[11])  # 未平倉餘額-空方-口數
                net_oi = _to_number(vals[13])    # 未平倉餘額-多空淨額-口數
            except IndexError:
                continue
            records.append({
                "商品名稱": current_name,
                "身份別": role_cell,
                "未平倉多方口數": long_oi,
                "未平倉空方口數": short_oi,
                "未平倉淨額口數": net_oi,
            })
    return pd.DataFrame(records)


def calc_foreign_size(df_futures):
    """外資大小 = 外資-臺股期貨 未平倉淨額 + 外資-小型臺指期貨 未平倉淨額 / 4"""
    def get_net(name):
        sub = df_futures[(df_futures["商品名稱"].str.contains(name, na=False)) & (df_futures["身份別"] == "外資")]
        if sub.empty:
            raise RuntimeError(f"找不到「{name}」外資未平倉資料")
        return sub.iloc[0]["未平倉淨額口數"]

    tx_net = get_net("臺股期貨")
    mtx_net = get_net("小型臺指期貨")
    return tx_net + mtx_net / 4


def calc_leek_index(df_futures, mtx_total_oi):
    """
    韭菜指數 = (小型臺指期貨 三大法人未平倉賣方合計 - 買方合計) / 小型臺指期貨全市場未平倉口數
    """
    sub = df_futures[df_futures["商品名稱"].str.contains("小型臺指期貨", na=False)]
    long_sum = sub["未平倉多方口數"].sum()
    short_sum = sub["未平倉空方口數"].sum()
    retail_long = mtx_total_oi - long_sum
    retail_short = mtx_total_oi - short_sum
    return (retail_long - retail_short) / mtx_total_oi


# --------------------------------------------------------------------------
# 4. TAIFEX 大額交易人未沖銷部位 -> 前5與10 分頁（僅需臺股期貨全市場）
# --------------------------------------------------------------------------

def fetch_taifex_large_trader():
    """回傳 (前五大留倉淨額, 前十大留倉淨額)，取『臺股期貨』『所有契約』列，全市場欄位。"""
    url = "https://www.taifex.com.tw/cht/3/largeTraderFutQryTbl"
    resp = _get(url)
    tables = pd.read_html(io.StringIO(resp.text))
    df = max(tables, key=len)
    df.columns = [str(c).strip() for c in range(df.shape[1])]

    target_idx = None
    for i, row in df.iterrows():
        name_cell = str(row[0])
        month_cell = str(row[1])
        if "臺股期貨" in name_cell:
            target_idx = i
        if target_idx is not None and "所有契約" in month_cell:
            target_row = df.iloc[i]
            top5_buy = _to_number(target_row[2])   # 前五大-買方-部位數(全市場)
            top5_sell = _to_number(target_row[6])  # 前五大-賣方-部位數(全市場)
            top10_buy = _to_number(target_row[4])  # 前十大-買方-部位數(全市場)
            top10_sell = _to_number(target_row[8])  # 前十大-賣方-部位數(全市場)
            return top5_buy - top5_sell, top10_buy - top10_sell
    raise RuntimeError("找不到臺股期貨『所有契約』列，頁面格式可能已變更")


# --------------------------------------------------------------------------
# 5. TAIFEX 選擇權買賣權比 (PCR)
# --------------------------------------------------------------------------

def fetch_taifex_pcr():
    """回傳當日『買賣權未平倉量比率%』。"""
    url = "https://www.taifex.com.tw/cht/3/pcRatioExcel"
    resp = _get(url)
    tables = pd.read_html(io.StringIO(resp.text))
    df = max(tables, key=len)
    df.columns = [str(c).strip() for c in range(df.shape[1])]
    first_row = df.iloc[0]
    return _to_number(first_row[6])  # 買賣權未平倉量比率%


# --------------------------------------------------------------------------
# 6. TAIFEX 三大法人商品別買賣口數金額 - 選擇權 -> 法人選擇權 分頁
# --------------------------------------------------------------------------

def fetch_taifex_options_institutional():
    """
    外(選) = 外資-買權未平倉買賣差額(契約金額) - 外資-賣權未平倉買賣差額(契約金額)
    只取『臺指選擇權』列。

    !! 重要（已用 web_fetch 實測確認）：
    期交所這份『未平倉餘額』欄位是「隔一個交易日」才公布的——收盤當天查詢會看到
    「未平倉口數與契約金額尚未揭露」，全部是「-」。也就是說：
      - 如果排程設在「收盤後當晚」執行，這一項會抓不到值（回傳 None 或報錯）。
      - 建議把排程時間改成「隔天開盤前」執行，才能抓到前一交易日的未平倉資料，
        或是把這個欄位獨立成另一個排程，晚一天寫入。
    """
    url = "https://www.taifex.com.tw/cht/3/callsAndPutsDateExcel"
    resp = _get(url)
    tables = pd.read_html(io.StringIO(resp.text))
    df = max(tables, key=len)
    df.columns = [str(c).strip() for c in range(df.shape[1])]

    call_diff = None
    put_diff = None
    current_product = None
    current_type = None
    for _, row in df.iterrows():
        vals = list(row)
        p = str(vals[1]).strip()
        if p and p not in ("nan", ""):
            current_product = p
        t = str(vals[2]).strip()
        if t in ("買權", "賣權"):
            current_type = t
        role = str(vals[3]).strip()
        if current_product == "臺指選擇權" and role == "外資":
            raw_val = str(vals[-1]).strip()
            if raw_val in ("-", "", "nan"):
                raise RuntimeError(
                    "臺指選擇權未平倉資料尚未公布（期交所通常隔一個交易日才揭露），"
                    "請改在隔天開盤前重新執行這個項目。"
                )
            # 未平倉餘額-買賣差額-契約金額 欄位（依原表結構為倒數第一欄）
            diff_amount = _to_number(raw_val)
            if current_type == "買權":
                call_diff = diff_amount
            elif current_type == "賣權":
                put_diff = diff_amount
    if call_diff is None or put_diff is None:
        raise RuntimeError("找不到臺指選擇權-外資 未平倉買賣差額，頁面格式可能已變更")
    return call_diff - put_diff


# --------------------------------------------------------------------------
# 7. TAIFEX 期貨每日行情 -> 期貨未平倉 / 小台(散戶指標) 分頁
# --------------------------------------------------------------------------

def fetch_taifex_daily_market(commodity_id=None):
    """
    回傳 DataFrame：契約, 到期月份, 未沖銷契約量
    commodity_id=None 時抓「大台指(TX)」為主的全部契約頁；
    commodity_id='MTX' 時抓小型臺指期貨頁。
    """
    url = "https://www.taifex.com.tw/cht/3/futDailyMarketExcel"
    params = {"commodity_id": commodity_id} if commodity_id else {}
    resp = _get(url, params=params)
    tables = pd.read_html(io.StringIO(resp.text))
    df = max(tables, key=len)
    df.columns = [str(c).strip() for c in range(df.shape[1])]
    df = df.rename(columns={"0": "契約", "1": "到期月份", "12": "未沖銷契約量"})
    return df


def calc_open_interest_total(df_market, contract_code):
    """加總指定契約代碼（如 TX）在所有到期月份的未沖銷契約量。"""
    sub = df_market[df_market["契約"] == contract_code]
    total = 0.0
    for v in sub["未沖銷契約量"]:
        n = _to_number(v)
        if n is not None:
            total += n
    return total


# --------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------

def build_today_row():
    """
    逐項抓取，單一項目失敗不影響其他項目：失敗的欄位填 None，並在 errors
    清單記錄原因，最後印出來讓你知道哪些要人工補或延後重跑。
    """
    today = dt.date.today()
    row = {col: None for col in COLUMNS}
    row["日期"] = today.strftime("%Y-%m-%d")
    errors = []

    def safe(step_name, fn):
        print(f"抓取 {step_name} ...")
        try:
            return fn()
        except Exception as e:
            errors.append(f"[{step_name}] {e}")
            print(f"  -> 失敗: {e}")
            return None

    volume_amount = safe("TWSE 成交量", fetch_twse_volume)
    if volume_amount is not None:
        row["成交量"] = f"{volume_amount / 1e8:,.1f}億"

    inst = safe("TWSE 三大法人買賣超", fetch_twse_institutional)
    if inst is not None:
        row["外資"] = round(inst["外資"] / 1e8, 1)
        row["投信"] = round(inst["投信"] / 1e8, 1)
        row["自營"] = round(inst["自營"] / 1e8, 1)

    df_futures = safe("TAIFEX 三大法人期貨未平倉", fetch_taifex_futures_positions)
    if df_futures is not None:
        foreign_size = safe("計算外資大小", lambda: calc_foreign_size(df_futures))
        row["外資大小"] = foreign_size

    large_trader = safe("TAIFEX 大額交易人未沖銷部位", fetch_taifex_large_trader)
    if large_trader is not None:
        row["前五大交易人留倉"], row["前十大交易人留倉"] = large_trader

    row["選PCR"] = safe("TAIFEX 選擇權 PCR", fetch_taifex_pcr)

    # 外(選)：收盤當晚通常會失敗（期交所隔一個交易日才公布未平倉），屬預期內
    row["外(選)"] = safe("TAIFEX 法人選擇權", fetch_taifex_options_institutional)

    df_tx_market = safe("TAIFEX 大台指每日行情", fetch_taifex_daily_market)
    if df_tx_market is not None:
        row["未平倉口數"] = safe(
            "計算大台未沖銷口數合計", lambda: calc_open_interest_total(df_tx_market, "TX")
        )

    if df_futures is not None:
        df_mtx_market = safe(
            "TAIFEX 小台指每日行情", lambda: fetch_taifex_daily_market("MTX")
        )
        if df_mtx_market is not None:
            mtx_oi_total = calc_open_interest_total(df_mtx_market, "MTX")
            leek_index = safe(
                "計算韭菜指數", lambda: calc_leek_index(df_futures, mtx_oi_total)
            )
            if leek_index is not None:
                row["韭菜指數"] = round(leek_index, 6)

    if errors:
        print("\n以下項目本次沒有抓到，已略過（其餘欄位仍會正常寫入）：")
        for e in errors:
            print(" -", e)

    return row


def append_to_history(row, path=HISTORY_CSV):
    file_exists = os.path.exists(path)
    if file_exists:
        with open(path, newline="", encoding="utf-8-sig") as f:
            existing_dates = {r["日期"] for r in csv.DictReader(f)}
        if row["日期"] in existing_dates:
            print(f"{row['日期']} 已存在，略過寫入。")
            return
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        if not file_exists:
            writer.writeheader()
        writer.writerow(row)
    print(f"已寫入 {row['日期']} 的資料到 {path}")


def main():
    row = build_today_row()
    print(row)
    append_to_history(row)


if __name__ == "__main__":
    main()
