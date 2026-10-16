#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
先行指標每日自動更新腳本
=========================
複製自「先行指標.xlsm」中 Power Query + 公式的邏輯，改用 Python 實作，
可放在雲端排程（例如 GitHub Actions）每天自動執行，不需要開電腦或用手機操作 Excel。

執行後會：
  1. 向證交所(TWSE)、期交所(TAIFEX)抓取當日公開資料
  2. 依照原始 Excel 檔的公式，計算出「先行指標」11 個欄位
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
  v2 發現休市日（2026-09-25、09-28）寫進了假資料：證交所欄位變成 0，
     期交所欄位是前一個交易日的數字。原因是期交所網址沒帶日期、預設回傳
     「最新已公布的一天」，證交所查無資料時程式又預設填 0。已修正為：
       - 期交所每個頁面都先核對頁面上的資料日期，不是今天就不採用
       - 證交所檢查 stat / date 欄位，查無資料或日期不符就不採用
       - 三大法人找不到對應項目時改為失敗（留空），不再預設填 0
       - 「今天」一律用台灣時間計算，手動執行時也不會跨日算錯
       - 同一天重跑時，只補上之前沒抓到的空白欄位（不用再手動刪列）
  v3 2026-10-16 期交所頁面連線逾時造成欄位空白：連線逾時/錯誤時自動重試 3 次，
     逾時秒數由 20 秒放寬為 30 秒。
     重跑時「投信」改以最新數字為準（證交所公布後會修正）。

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
import time
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from bs4 import BeautifulSoup

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}
TIMEOUT = 30
RETRIES = 3
HISTORY_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), "history.csv")

COLUMNS = [
    "日期", "成交量", "外資", "投信", "自營", "外資大小",
    "前五大交易人留倉", "前十大交易人留倉", "選PCR", "外(選)",
    "韭菜指數", "未平倉口數",
]

# 同一天重跑時「以最新抓到的數字為準」的欄位（其他欄位只補空白、不覆蓋）。
# 證交所的投信買賣超在公布後會再修正（例如 2026-10-16：117.5 -> 110.5 億）。
OVERWRITE_COLUMNS = {"投信"}


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
    """
    期交所網站在 GitHub Actions 上常常「Read timed out」（實測幾乎每天都有一兩個
    頁面逾時，造成該欄位空白），所以連線逾時、連線錯誤、伺服器 5xx 錯誤時
    最多重試 RETRIES 次，每次間隔逐漸拉長。
    """
    for attempt in range(1, RETRIES + 1):
        try:
            resp = requests.get(url, headers=HEADERS, params=params, timeout=TIMEOUT)
            if resp.status_code >= 500:
                resp.raise_for_status()
            break
        except (requests.Timeout, requests.ConnectionError, requests.HTTPError) as e:
            if attempt == RETRIES:
                raise
            wait = 5 * attempt
            print(f"  -> 第 {attempt} 次連線失敗（{e}），{wait} 秒後重試...")
            time.sleep(wait)
    resp.raise_for_status()
    return resp


def _clean(text):
    """去除多餘空白、換行，方便比對文字。"""
    return re.sub(r"\s+", "", text or "")


def _today():
    """台灣時間的今天。GitHub Actions 主機是 UTC，手動在台灣早上 8 點前執行會差一天。"""
    return dt.datetime.now(ZoneInfo("Asia/Taipei")).date()


def _check_taifex_date(html, page_name):
    """
    確認期交所頁面上的資料日期就是今天，不是的話就報錯（該欄位留空）。

    期交所這些網址沒帶日期參數時，會回傳「最新已公布的那一天」：休市日或
    當天資料還沒公布時，拿到的其實是前一個交易日的數字。實測每個頁面裡
    第一個出現的 YYYY/M/D 日期就是資料日期（PCR 頁則是最新一筆資料列）。
    """
    m = re.search(r"(20\d{2})/(\d{1,2})/(\d{1,2})", html)
    if not m:
        raise RuntimeError(f"{page_name} 頁面找不到資料日期，網站可能已經改版")
    page_date = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    if page_date != _today():
        raise RuntimeError(
            f"{page_name} 目前的資料日期是 {page_date}，不是今天（今天的資料可能還沒公布），不採用"
        )


def _check_twse_json(data, page_name):
    """證交所查無資料時 stat 不是 OK；有資料時 date 欄位是 YYYYMMDD，要等於今天。"""
    if str(data.get("stat", "")).upper() != "OK":
        raise RuntimeError(f"{page_name} 查無今天資料（{data.get('stat')}）")
    if data.get("date") and data["date"] != _today().strftime("%Y%m%d"):
        raise RuntimeError(f"{page_name} 回傳的資料日期是 {data['date']}，不是今天，不採用")


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
    return max(parse_all_table_grids(html), key=len)  # 資料列最多的表格通常就是主表


def parse_all_table_grids(html):
    """
    展開 rowspan/colspan 的實際邏輯，回傳頁面上「每一張」表格（而不是只挑
    列數最多的那張）。

    背景：期交所「期貨每日行情」這類頁面常常同時放好幾張表格（一般交易
    時段、盤後交易時段、價差行情表……），用「列數最多」去猜哪張是主表
    並不可靠——這是造成「未沖銷契約量」欄位有時候抓不到的原因。改成
    把每張表都展開，之後用表頭文字去比對、找出真正需要的那張。
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

    return [expand(t) for t in tables]


# --------------------------------------------------------------------------
# 1. TWSE 大盤成交量
# --------------------------------------------------------------------------

def fetch_twse_volume():
    """證交所每日成交量統計，取『總計(1~15)』成交金額(元)。"""
    url = "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX"
    # TWSE 的 rwd JSON API 比原本的 HTML 頁面更穩定，直接查當日
    today = _today().strftime("%Y%m%d")
    resp = _get(url, params={"date": today, "type": "ALL", "response": "json"})
    data = resp.json()
    _check_twse_json(data, "TWSE 成交量")
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
    today = _today().strftime("%Y%m%d")
    resp = _get(url, params={"dayDate": today, "type": "day", "response": "json"})
    data = resp.json()
    _check_twse_json(data, "TWSE 三大法人")
    rows = data.get("data", [])
    # 預設 None 而不是 0：找不到項目時要能看出「沒抓到」，不能跟「買賣超剛好 0」混在一起
    result = {"自營自行": None, "自營避險": None, "投信": None, "外資": None}
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
    missing = [k for k, v in result.items() if v is None]
    if missing:
        raise RuntimeError(f"三大法人資料缺少 {'、'.join(missing)}，頁面格式可能已變更")
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
    _check_taifex_date(resp.text, "期交所三大法人期貨")
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
    _check_taifex_date(resp.text, "期交所大額交易人")
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
    _check_taifex_date(resp.text, "期交所選擇權 PCR")
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
    _check_taifex_date(resp.text, "期交所法人選擇權")
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
    回傳 DataFrame，只含『到期月份』『未沖銷契約量』兩欄。

    這個頁面同時放了好幾張表格（一般交易時段／盤後交易時段／價差行情表），
    不能用「列數最多」去猜哪張是主表——這裡改成掃過每一張表格，用表頭
    文字（同時出現「到期月份」和「未沖銷契約量」）找出真正要的那張，
    欄位位置用表頭比對，不寫死固定欄號。
    """
    url = "https://www.taifex.com.tw/cht/3/futDailyMarketExcel"
    params = {"commodity_id": commodity_id} if commodity_id else {}
    resp = _get(url, params=params)
    _check_taifex_date(resp.text, f"期交所每日行情({commodity_id or 'TX'})")
    grids = parse_all_table_grids(resp.text)

    for grid in grids:
        for row_idx, row in enumerate(grid[:5]):  # 表頭通常在前幾列
            cleaned = [_clean(c) for c in row]
            has_month = any("到期月份" in c for c in cleaned)
            has_oi = any("未沖銷契約量" in c for c in cleaned)
            if not (has_month and has_oi):
                continue
            col_month = next(i for i, c in enumerate(cleaned) if "到期月份" in c)
            col_oi = next(i for i, c in enumerate(cleaned) if "未沖銷契約量" in c)
            records = []
            for data_row in grid[row_idx + 1:]:
                if len(data_row) <= max(col_month, col_oi):
                    continue
                month = data_row[col_month].strip()
                oi = _to_number(data_row[col_oi])
                if month and oi is not None:
                    records.append({"到期月份": month, "未沖銷契約量": oi})
            if records:
                return pd.DataFrame(records)
    raise RuntimeError("找不到期貨每日行情的未沖銷契約量欄位，頁面格式可能已變更")


def calc_open_interest_total(df_market):
    """加總表格裡所有到期月份的未沖銷契約量。"""
    return float(df_market["未沖銷契約量"].sum())


# --------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------

# --------------------------------------------------------------------------
# 休市日判斷（國定假日 / 補假 / 颱風之外的例行休市）
# --------------------------------------------------------------------------

def _normalize_date(raw):
    """把 '2026-09-28'、'2026/09/28'、'115/09/28'、'115年09月28日' 統一成 'YYYY-MM-DD'。"""
    m = re.search(r"(\d{2,4})\D+(\d{1,2})\D+(\d{1,2})", str(raw))
    if not m:
        return None
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 1911:  # 民國年
        y += 1911
    return f"{y:04d}-{mo:02d}-{d:02d}"


def is_market_closed(day):
    """
    回傳 (是否休市, 原因)。
    1. 週六、週日 -> 休市
    2. 證交所「市場開休市日期」官方 API 列出的日期 -> 休市
       （名稱含「開始交易」「最後交易」的是有開盤的資訊列，不算休市）
    查詢官方日曆失敗時「不擋」（回傳 False），寧可多抓一次，也不要漏掉資料；
    此時原本的逐欄位容錯仍會運作。
    """
    if day.weekday() >= 5:
        return True, "週末"
    url = "https://www.twse.com.tw/rwd/zh/holidaySchedule/holidaySchedule"
    try:
        resp = _get(url, params={"response": "json", "date": f"{day.year}0101"})
        data = resp.json()
    except Exception as e:
        print(f"[休市判斷] 查詢證交所休市日曆失敗（{e}），視為有開盤，照常執行。")
        return False, None
    if str(data.get("stat", "")).lower() != "ok":
        print(f"[休市判斷] 休市日曆回傳異常（stat={data.get('stat')}），視為有開盤，照常執行。")
        return False, None
    iso = day.isoformat()
    for r in data.get("data", []):
        if not r or _normalize_date(r[0]) != iso:
            continue
        name = str(r[1]).strip() if len(r) > 1 else ""
        desc = str(r[2]).strip() if len(r) > 2 else ""
        if "開始交易" in name + desc or "最後交易" in name + desc:
            return False, None
        return True, (name + " " + desc).strip()
    return False, None


def build_today_row():
    """
    逐項抓取，單一項目失敗不影響其他項目：失敗的欄位填 None，並在 errors
    清單記錄原因，最後印出來讓你知道哪些要人工補或延後重跑。
    """
    today = _today()
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
            "計算大台未沖銷口數合計", lambda: calc_open_interest_total(df_tx_market)
        )

    if df_futures is not None:
        df_mtx_market = safe(
            "TAIFEX 小台指每日行情", lambda: fetch_taifex_daily_market("MTX")
        )
        if df_mtx_market is not None:
            mtx_oi_total = safe(
                "計算小台未沖銷口數合計", lambda: calc_open_interest_total(df_mtx_market)
            )
            if mtx_oi_total is not None:
                leek_index = safe(
                    "計算韭菜指數", lambda: calc_leek_index(df_futures, mtx_oi_total)
                )
                if leek_index is not None:
                    row["韭菜指數"] = f"{leek_index * 100:.2f}%"

    if errors:
        print("\n以下項目本次沒有抓到，已略過（其餘欄位仍會正常寫入）：")
        for e in errors:
            print(" -", e)

    return row


def append_to_history(row, path=HISTORY_CSV):
    """
    把今天這一列寫進 history.csv：
      - 日期還不存在 -> 新增一列
      - 日期已存在   -> 只補上原本空白、這次有抓到的欄位（已有的值不覆蓋）。
        這樣某次執行因為連線逾時少抓幾欄時，重跑一次就能補齊，不用手動刪列。
        例外：OVERWRITE_COLUMNS 裡的欄位只要這次有抓到，就更新成最新數字。
    每次都把整個檔案重新寫出：換行一律是 CRLF、去掉空白列，就算檔案在 GitHub
    網頁上手動編輯過（換行混用、結尾沒換行），也不會讓新的一列黏到上一列後面。
    """
    rows = []
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = [r for r in csv.DictReader(f) if (r.get("日期") or "").strip()]

    existing = next((r for r in rows if r["日期"].strip() == row["日期"]), None)
    if existing is None:
        rows.append(row)
        print(f"新增 {row['日期']} 的資料到 {path}")
    else:
        filled, updated = [], []
        for col in COLUMNS:
            if col == "日期" or row.get(col) is None:
                continue
            old = str(existing.get(col) or "").strip()
            if not old:
                filled.append(col)
            elif col in OVERWRITE_COLUMNS and old != str(row[col]):
                updated.append(f"{col} {old} -> {row[col]}")
            else:
                continue
            existing[col] = row[col]
        if not filled and not updated:
            print(f"{row['日期']} 已存在，沒有需要補上或更新的欄位。")
            return
        if filled:
            print(f"{row['日期']} 已存在，補上空白欄位：{'、'.join(filled)}")
        if updated:
            print(f"{row['日期']} 已存在，更新為最新數字：{'、'.join(updated)}")

    # 不寫 BOM：前端 PapaParse 會把 BOM 當成第一個欄位名稱的一部分，讀不到「日期」
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main():
    today = _today()
    force = os.environ.get("FORCE_RUN", "").strip().lower() in ("1", "true", "yes")
    closed, reason = is_market_closed(today)
    if closed and not force:
        print(f"{today} 台股休市（{reason}），略過本次更新，不寫入任何資料。")
        return
    if closed and force:
        print(f"{today} 台股休市（{reason}），但已指定強制執行，繼續。")
    try:
        row = build_today_row()
    except Exception:
        # 保底防線：就算 build_today_row() 本身邏輯出現非預期例外，
        # 也印出完整錯誤內容方便排查，同時讓這次執行清楚地算失敗，
        # 而不是留下一個內容不完整、卻被誤判成功的 history.csv。
        import traceback
        traceback.print_exc()
        sys.exit(1)
    print(row)
    if all(row[col] is None for col in COLUMNS if col != "日期"):
        # 一個欄位都沒抓到（例如休市或資料還沒公布）就不要寫入只有日期的空白列，
        # 並讓這次執行算失敗，方便在 GitHub Actions 發現。
        print("所有欄位都沒有抓到今天的資料，不寫入 history.csv。")
        sys.exit(1)
    append_to_history(row)


if __name__ == "__main__":
    main()
