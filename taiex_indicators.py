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

版本紀錄（依實際執行結果修正過）：
  v1 第一次上線後，實測發現兩類問題：
    1. TWSE 三大法人買賣超「外資」欄一律抓到 0：因為原本用「排除含
       『自營商』字樣」來判斷，但正確項目全名是「外資及陸資(不含外資
       自營商)」，字串裡本來就含有「自營商」，排除法誤判——已修正為
       直接比對開頭文字。
    2. 外資大小 / 前五大十大留倉 / 選PCR / 韭菜指數 抓到空值：因為
       pandas.read_html 在期交所這種多層合併儲存格(rowspan/colspan)
       的表格上常常抓錯欄位對齊——已改用 BeautifulSoup 手動展開合併
       儲存格（parse_table_grid），不再依賴 pandas 的自動判斷。

執行需要的套件：requests、pandas、lxml、beautifulsoup4
  （pip install requests pandas lxml beautifulsoup4）

已知限制（已用 web_fetch 實測證實）：
  - 期交所「三大法人-選擇權買賣權分計」的未平倉欄位是隔一個交易日才公布，
    收盤當天查詢會全部顯示「-」。所以「外(選)」這個欄位沒辦法在收盤當晚
    抓到，建議排程改成「隔天開盤前」執行，或把這欄獨立成晚一天的排程。

若抓取失敗，腳本會印出錯誤並中止該欄位，不會用假資料填補；其餘欄位仍會
正常寫入，不會整批失敗。
"""

import csv
import datetime as dt
import io
import os
import re
import sys

import pandas as pd
import requests
from bs4 import BeautifulSoup

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


def _clean(text):
    """去除多餘空白、換行，方便比對文字。"""
    return re.sub(r"\s+", "", text or "")


def parse_table_grid(html):
    """
    手動解析 HTML 表格並展開 rowspan/colspan，回傳「乾淨的二維陣列」。

    背景：期交所這些「匯出成 Excel 用」的網頁表格大量使用合併儲存格
    （rowspan/colspan），而且沒有用標準的 <th> 標記表頭。實測發現
    pandas.read_html 在這種表格上常常抓錯欄位對齊（表頭沒被跳過、
    合併儲存格沒有正確複製到每一列），是先前「外資大小/前五大十大
    留倉/選PCR/韭菜指數」抓空值的主因。這裡改用 BeautifulSoup 手動
    展開合併儲存格，讓每一列、每一欄都對應到畫面上看到的實際位置，
    表頭列也會被還原成一般文字列（後續用文字比對來跳過，而不是用
    位置跳過），穩定度比較高。
    """
    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")
    if not tables:
        raise RuntimeError("頁面中找不到任何表格，網站可能已經改版")

    def expand(table):
        grid = []
        pending = {}  # col_index -> [剩餘列數, 值]
        for tr in table.find_all("tr"):
            cells = tr.find_all(["td", "th"])
            grid_row = []
            col = 0
            cell_iter = iter(cells)
            while True:
                if col in pending and pending[col][0] > 0:
                    grid_row.append(pending[col][1])
                    pending[col][0] -= 1
                    col += 1
                    continue
                try:
                    cell = next(cell_iter)
                except StopIteration:
                    break
                text = cell.get_text(strip=True)
                try:
                    colspan = int(cell.get("colspan", 1) or 1)
                except ValueError:
                    colspan = 1
                try:
                    rowspan = int(cell.get("rowspan", 1) or 1)
                except ValueError:
                    rowspan = 1
                for _ in range(colspan):
                    grid_row.append(text)
                    if rowspan > 1:
                        pending[col] = [rowspan - 1, text]
                    col += 1
            if grid_row:
                grid.append(grid_row)
        width = max((len(r) for r in grid), default=0)
        return [r + [""] * (width - len(r)) for r in grid]

    candidates = [expand(t) for t in tables]
    return max(candidates, key=len)  # 資料列最多的表格通常就是主表


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
        elif name.startswith("外資及陸資"):
            # 注意：不能用「排除含『自營商』字樣」來判斷，因為這個項目
            # 的完整名稱是「外資及陸資(不含外資自營商)」，字串裡本來就
            # 含有「自營商」三個字，用排除法會誤判成別的項目而抓到 0。
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

    欄位順序（rowspan 展開後，對應畫面上實際看到的位置）：
    0序號 1商品名稱 2身份別 3口數多方 4契約金額多方 5口數空方 6契約金額空方
    7口數多空淨額 8契約金額多空淨額 9口數未平倉多方 10契約金額未平倉多方
    11口數未平倉空方 12契約金額未平倉空方 13口數未平倉淨額 14契約金額未平倉淨額
    """
    url = "https://www.taifex.com.tw/cht/3/futContractsDateExcel"
    resp = _get(url)
    grid = parse_table_grid(resp.text)
    records = []
    for row in grid:
        if len(row) < 15:
            continue
        role_cell = row[2].strip()
        if role_cell not in ("自營商", "投信", "外資"):
            continue  # 跳過表頭列、合計列等非資料列
        records.append({
            "商品名稱": row[1].strip(),
            "身份別": role_cell,
            "未平倉多方口數": _to_number(row[9]),
            "未平倉空方口數": _to_number(row[11]),
            "未平倉淨額口數": _to_number(row[13]),
        })
    if not records:
        raise RuntimeError("解析不到任何期貨三大法人資料列，頁面格式可能已變更")
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
    """
    回傳 (前五大留倉淨額, 前十大留倉淨額)，取『臺股期貨』『所有契約』列，全市場欄位。

    欄位順序：0契約名稱 1到期月份 2前五大買方部位數 3前五大買方百分比
    4前十大買方部位數 5前十大買方百分比 6前五大賣方部位數 7前五大賣方百分比
    8前十大賣方部位數 9前十大賣方百分比 10全市場未沖銷部位數
    """
    url = "https://www.taifex.com.tw/cht/3/largeTraderFutQryTbl"
    resp = _get(url)
    grid = parse_table_grid(resp.text)
    for row in grid:
        if len(row) < 9:
            continue
        name_cell = _clean(row[0])
        month_cell = _clean(row[1])
        if "臺股期貨" in name_cell and "所有契約" in month_cell:
            top5_buy = _to_number(row[2])
            top10_buy = _to_number(row[4])
            top5_sell = _to_number(row[6])
            top10_sell = _to_number(row[8])
            if None in (top5_buy, top10_buy, top5_sell, top10_sell):
                continue
            return top5_buy - top5_sell, top10_buy - top10_sell
    raise RuntimeError("找不到臺股期貨『所有契約』列，頁面格式可能已變更")


# --------------------------------------------------------------------------
# 5. TAIFEX 選擇權買賣權比 (PCR)
# --------------------------------------------------------------------------

def fetch_taifex_pcr():
    """
    回傳當日『買賣權未平倉量比率%』。
    欄位順序：0日期 1賣權成交量 2買權成交量 3買賣權成交量比率% 4賣權未平倉量
    5買權未平倉量 6買賣權未平倉量比率%
    用「日期欄是不是像 2026/8/26 這種格式」來找出第一筆真正的資料列，
    不依賴固定的列號（表頭有沒有被 pandas 正確跳過並不影響這裡）。
    """
    url = "https://www.taifex.com.tw/cht/3/pcRatioExcel"
    resp = _get(url)
    grid = parse_table_grid(resp.text)
    date_pattern = re.compile(r"^\d{4}/\d{1,2}/\d{1,2}$")
    for row in grid:
        if len(row) >= 7 and date_pattern.match(row[0].strip()):
            return _to_number(row[6])
    raise RuntimeError("找不到選擇權 PCR 資料列，頁面格式可能已變更")


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
    grid = parse_table_grid(resp.text)
    # 欄位順序：0序號 1商品名稱 2權別 3身份別 ... 最後一欄(-1)=未平倉買賣差額契約金額
    call_diff = None
    put_diff = None
    for row in grid:
        if len(row) < 5:
            continue
        product = row[1].strip()
        option_type = row[2].strip()
        role = row[3].strip()
        if product == "臺指選擇權" and role == "外資":
            raw_val = row[-1].strip()
            if raw_val in ("-", "", "nan"):
                raise RuntimeError(
                    "臺指選擇權未平倉資料尚未公布（期交所通常隔一個交易日才揭露），"
                    "請改在隔天開盤前重新執行這個項目。"
                )
            diff_amount = _to_number(raw_val)
            if option_type == "買權":
                call_diff = diff_amount
            elif option_type == "賣權":
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
