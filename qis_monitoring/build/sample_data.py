#!/usr/bin/env python3
"""架空のテストデータ（すべてサンプル。実在の商品・指数・価格ではない）。

ここが唯一の正本で、build_workbook.py と verify_workbook.py の双方がここから読む。
`python3 sample_data.py` を実行すると ../sample_data/ に CSV を書き出す
（Power Query 取り込みテスト用のポジションファイルを含む）。

テストシナリオ（verify_workbook.py が期待値を検証する）:
  T01 NAV 下落           F001 週次 -2.80% / 月次 -5.3%（High アラート）
  T02 NAV 不一致         F002 Bloomberg NAV が公式比 +0.80%（MISMATCH）
                         F003 +0.10%（WARNING）/ F004 BBG 日付ずれ（DATE DIFFERENCE）
                         F005 BBG NAV なし（MISSING DATA）
  T03 新規ポジション     NDX 01/15/27 P19000、NDX 11/04/26 C21000
  T04 ポジション解消     NDX 11/06/26 P19600（Closed）、10/30/26 P19800（Expired）
  T05 Strike 接近        NDX 12/18/26 P20200 が OTM→ITM、11/20/26 P19500 が 2.74%
  T06 Vol Target 縮小    IDX_VT10 Exposure 118% → 83%（-29.7%）
  T07 Trend シグナル反転 IDX_TC Equity ES1 +0.60 → -0.20
  T08 Bloomberg 欠損     SX5E（非重要）・QIS_TC（重要）欠損、JGB10 は STALE
  T09 複数地銀・複数ファンド  5 行 × 5 ファンド
  T10 メール生成         山田 太郎（A銀行・B銀行）
"""
from __future__ import annotations

import csv
import datetime as dt
import random
from pathlib import Path

D = dt.date
EVAL_DATE = D(2026, 10, 30)
PREV_DATE = D(2026, 10, 23)
RUN_DATE = D(2026, 11, 2)

POS_COLUMNS = [
    "Position Date", "Index Name", "Name", "BloombergTicker", "Option Mid Price", "Strike",
    "Maturity Date", "Notional(USD)", "Delta", "Gamma", "Vega", "Theta", "Implied Vol",
    "NDX Spot", "Index Level", "Weight", "open_date",
]

IDX_PW = "Sample QIS NDX Put Write Index"
IDX_PS = "Sample QIS NDX Put Spread Index"


def _opt(pos_date, idx, ticker, mid, strike, mat, notional, delta, gamma, vega, theta, iv, spot, lvl, w, open_d):
    # Name は Bloomberg の銘柄名に近い形（"NDX 11/20/26 P19500"）
    parts = ticker.split(" ")
    name = f"{parts[0]} {parts[2]} {parts[3]}"
    return dict(zip(POS_COLUMNS, [pos_date, idx, name, ticker, mid, strike, mat, notional, delta,
                                  gamma, vega, theta, iv, spot, lvl, w, open_d]))


def positions():
    p, c = PREV_DATE, EVAL_DATE
    s0, s1 = 20850.0, 20050.0
    pw0, pw1 = 1234.56, 1208.31
    ps0, ps1 = 987.65, 979.10
    prev = [
        _opt(p, IDX_PW, "NDX US 10/30/26 P19800 Index", 18.40, 19800, D(2026, 10, 30), 20_000_000, -0.08, 0.00012, 3.1, -9.5, 0.198, s0, pw0, -0.100, D(2026, 9, 25)),
        _opt(p, IDX_PW, "NDX US 11/06/26 P19600 Index", 42.10, 19600, D(2026, 11, 6), 30_000_000, -0.10, 0.00010, 6.2, -6.8, 0.202, s0, pw0, -0.150, D(2026, 10, 2)),
        _opt(p, IDX_PW, "NDX US 11/20/26 P19500 Index", 95.30, 19500, D(2026, 11, 20), 50_000_000, -0.14, 0.00009, 12.5, -5.1, 0.205, s0, pw0, -0.250, D(2026, 10, 16)),
        _opt(p, IDX_PW, "NDX US 12/18/26 P20200 Index", 365.00, 20200, D(2026, 12, 18), 25_000_000, -0.33, 0.00011, 28.0, -6.0, 0.208, s0, pw0, -0.125, D(2026, 10, 9)),
        _opt(p, IDX_PW, "NDX US 12/18/26 P18500 Index", 120.50, 18500, D(2026, 12, 18), 25_000_000, -0.12, 0.00005, 18.0, -3.2, 0.221, s0, pw0, -0.125, D(2026, 9, 18)),
        _opt(p, IDX_PW, "NDX US 12/18/26 P18500 Index", 120.50, 18500, D(2026, 12, 18), 15_000_000, -0.12, 0.00005, 18.0, -3.2, 0.221, s0, pw0, -0.075, D(2026, 10, 9)),
        _opt(p, IDX_PS, "NDX US 12/18/26 P19000 Index", 180.20, 19000, D(2026, 12, 18), 40_000_000, -0.17, 0.00006, 22.0, -3.9, 0.215, s0, ps0, 0.300, D(2026, 9, 18)),
        _opt(p, IDX_PS, "NDX US 12/18/26 P17500 Index", 70.40, 17500, D(2026, 12, 18), 40_000_000, -0.07, 0.00003, 12.0, -2.0, 0.238, s0, ps0, -0.300, D(2026, 9, 18)),
    ]
    cur = [
        _opt(c, IDX_PW, "NDX US 11/20/26 P19500 Index", 210.60, 19500, D(2026, 11, 20), 50_000_000, -0.30, 0.00019, 18.2, -12.4, 0.248, s1, pw1, -0.250, D(2026, 10, 16)),
        _opt(c, IDX_PW, "NDX US 11/20/26 P19500 Index", 210.60, 19500, D(2026, 11, 20), 10_000_000, -0.30, 0.00019, 18.2, -12.4, 0.248, s1, pw1, -0.050, D(2026, 10, 29)),
        _opt(c, IDX_PW, "NDX US 12/18/26 P20200 Index", 640.00, 20200, D(2026, 12, 18), 25_000_000, -0.52, 0.00016, 31.0, -8.2, 0.252, s1, pw1, -0.125, D(2026, 10, 9)),
        _opt(c, IDX_PW, "NDX US 12/18/26 P18500 Index", 260.30, 18500, D(2026, 12, 18), 25_000_000, -0.22, 0.00009, 25.0, -5.0, 0.262, s1, pw1, -0.125, D(2026, 9, 18)),
        _opt(c, IDX_PW, "NDX US 12/18/26 P18500 Index", 260.30, 18500, D(2026, 12, 18), 15_000_000, -0.22, 0.00009, 25.0, -5.0, 0.262, s1, pw1, -0.075, D(2026, 10, 9)),
        _opt(c, IDX_PW, "NDX US 01/15/27 P19000 Index", 520.50, 19000, D(2027, 1, 15), 30_000_000, -0.27, 0.00008, 38.5, -4.4, 0.255, s1, pw1, -0.150, D(2026, 10, 29)),
        _opt(c, IDX_PW, "NDX US 11/04/26 C21000 Index", 8.20, 21000, D(2026, 11, 4), 10_000_000, 0.06, 0.00007, 2.0, -9.0, 0.232, s1, pw1, -0.050, D(2026, 10, 28)),
        _opt(c, IDX_PS, "NDX US 12/18/26 P19000 Index", 395.00, 19000, D(2026, 12, 18), 40_000_000, -0.33, 0.00010, 30.0, -5.8, 0.258, s1, ps1, 0.300, D(2026, 9, 18)),
        _opt(c, IDX_PS, "NDX US 12/18/26 P17500 Index", 160.20, 17500, D(2026, 12, 18), 40_000_000, -0.14, 0.00005, 18.0, -3.1, 0.275, s1, ps1, -0.300, D(2026, 9, 18)),
    ]
    return [("positions_20261023.csv", prev), ("positions_20261030.csv", cur)]


# ------------------------------------------------------------------ masters
FUNDS = [
    # id, name, short, class, strategy, manager, ccy, inception, inc_nav, basis, dist, hedged, lag, bbg, maturity, status, fxkey
    ("F001", "サンプルQISファンドA（NDXプット・ライト）", "ファンドA", "JPY Hedged", "OPTION", "サンプル運用会社X", "JPY", D(2025, 6, 27), 10000.0, "TR", "N", "Y", 1, "SMPFNDA JP Equity", D(2028, 6, 30), "ACTIVE", "FX_USDJPY"),
    ("F002", "サンプルQISファンドB（NDXオプション複合）", "ファンドB", "USD", "OPTION", "サンプル運用会社X", "USD", D(2025, 9, 26), 100.0, "TR", "N", "N", 1, "SMPFNDB KY Equity", D(2027, 12, 24), "ACTIVE", "FX_USDJPY"),
    ("F003", "サンプルQISファンドC（米国株VolTarget）", "ファンドC", "JPY Unhedged", "VOLTARGET", "サンプル運用会社Y", "JPY", D(2025, 6, 27), 10000.0, "TR", "N", "N", 2, "SMPFNDC JP Equity", None, "ACTIVE", "FX_USDJPY"),
    ("F004", "サンプルQISファンドD（Trend/Carry・分配型）", "ファンドD", "JPY Dist", "TRENDCARRY", "サンプル運用会社Z", "JPY", D(2025, 6, 27), 10000.0, "PRICE", "Y", "Y", 2, "SMPFNDD JP Equity", None, "ACTIVE", "FX_USDJPY"),
    ("F005", "サンプルQISファンドE（新規設定）", "ファンドE", "USD", "OPTION", "サンプル運用会社X", "USD", D(2026, 9, 25), 100.0, "TR", "N", "N", 1, "", None, "ACTIVE", "FX_USDJPY"),
]

INDICES = [
    # id, file name, strategy, underlying, und key, qis key, vol key, ccy, greeks ok, spot col valid, notes
    ("IDX_NDXPW", IDX_PW, "OPTION", "NASDAQ 100", "EQ_NDX", "QIS_NDXPW", "VOL_VXN", "USD", "N", "Y", "サンプル"),
    ("IDX_NDXPS", IDX_PS, "OPTION", "NASDAQ 100", "EQ_NDX", "QIS_NDXPS", "VOL_VXN", "USD", "N", "Y", "サンプル"),
    ("IDX_VT10", "Sample QIS US Equity Vol Target 10% Index", "VOLTARGET", "S&P 500", "EQ_SPX", "QIS_VT10", "VOL_VIX", "USD", "N", "N", "サンプル"),
    ("IDX_TC", "Sample QIS Multi-Asset Trend & Carry Index", "TRENDCARRY", "マルチアセット", "EQ_SPX", "QIS_TC", "VOL_VIX", "USD", "N", "N", "サンプル"),
]

MAPPINGS = [
    ("M01", "F001", "IDX_NDXPW", 1.0, "Y", ""),
    ("M02", "F002", "IDX_NDXPW", 0.5, "Y", "複数指数に配分"),
    ("M03", "F002", "IDX_NDXPS", 0.5, "N", "複数指数に配分"),
    ("M04", "F003", "IDX_VT10", 1.0, "Y", ""),
    ("M05", "F004", "IDX_TC", 1.0, "Y", ""),
    ("M06", "F005", "IDX_NDXPS", 1.0, "Y", "保有先なし（DQテスト）"),
]

REPS = [
    ("R01", "山田 太郎", "taro.yamada@example.com", "金融法人営業部", "Y"),
    ("R02", "佐藤 花子", "hanako.sato@example.com", "金融法人営業部", "Y"),
    ("R03", "鈴木 一郎", "ichiro.suzuki@example.com", "地域金融機関営業部", "Y"),
]

CLIENTS = [
    ("C001", "A銀行", "地方銀行", "関東", "R01", D(2026, 10, 20), D(2026, 11, 5), ""),
    ("C002", "B銀行", "地方銀行", "近畿", "R01", D(2026, 9, 15), None, "前回接触から30日超"),
    ("C003", "C銀行", "第二地方銀行", "九州", "R02", D(2026, 10, 27), D(2026, 11, 20), ""),
    ("C004", "D銀行", "地方銀行", "東北", "R03", D(2026, 8, 30), D(2026, 11, 4), ""),
    ("C005", "E銀行", "地方銀行", "中部", "R02", D(2026, 10, 15), D(2026, 12, 1), ""),
]

# holding id, client, fund, book JPY, entry date, redemption plan
HOLDINGS_SPEC = [
    ("H001", "C001", "F001", 3_000_000_000, D(2025, 7, 25), D(2027, 6, 30)),
    ("H002", "C001", "F003", 1_000_000_000, D(2025, 10, 31), None),
    ("H003", "C002", "F002", 1_050_000_000, D(2025, 10, 31), D(2026, 12, 25)),
    ("H004", "C003", "F003", 800_000_000, D(2025, 12, 26), None),
    ("H005", "C003", "F004", 1_200_000_000, D(2025, 8, 29), None),
    ("H006", "C004", "F001", 500_000_000, D(2026, 3, 27), None),
    ("H007", "C005", "F004", 300_000_000, D(2026, 1, 30), None),
]
USDJPY_ENTRY = 150.0


# ------------------------------------------------------------------ NAV
def fridays(start, end):
    d = start
    while d <= end:
        yield d
        d += dt.timedelta(days=7)


def nav_history():
    """returns (official rows, bloomberg rows, distributions)"""
    rng = random.Random(20261030)
    official, bbg, dists = [], [], []
    tail = {  # 直近の週次（価格ベース）リターンを固定する
        "F001": [-0.008, -0.006, -0.005, -0.007, -0.028],
        "F002": [-0.004, -0.006, 0.002, -0.009, -0.016],
        "F003": [0.003, 0.002, -0.004, 0.001, -0.012],
        "F004": [-0.010, -0.012, -0.008, -0.006, -0.002],
        "F005": [0.002, -0.003, 0.001, -0.004, -0.010],
    }
    for f in FUNDS:
        fid, ccy, inc, nav0, dist_flag = f[0], f[6], f[7], f[8], f[10]
        days = list(fridays(inc, EVAL_DATE))
        n = len(days)
        navs = [nav0]
        for i in range(1, n):
            k = i - (n - 5)
            if k >= 0:
                r = tail[fid][k]
            elif fid == "F004":
                # 前半上昇 → 後半下落（現在ドローダウン > 10%）
                r = 0.006 + rng.gauss(0, 0.004) if i < n * 0.55 else -0.0045 + rng.gauss(0, 0.004)
            else:
                r = 0.0015 + rng.gauss(0, 0.009)
            navs.append(navs[-1] * (1 + r))
        for d, v in zip(days, navs):
            official.append((fid, d, round(v, 2), ccy, "OFFICIAL", "WEEKLY", f[9], f[3]))
        if dist_flag == "Y":
            for ex in (D(2026, 1, 28), D(2026, 4, 28), D(2026, 7, 28), D(2026, 10, 28)):
                dists.append((fid, ex, 50.0, ccy, "運用会社 分配金通知（サンプル）"))
        # Bloomberg NAV（照合用）: 直近 8 週
        if fid == "F005":
            continue
        for d, v in list(zip(days, navs))[-8:]:
            val = round(v, 2)
            if d == EVAL_DATE:
                if fid == "F002":
                    val = round(v * 1.008, 2)       # MISMATCH
                elif fid == "F003":
                    val = round(v * 1.001, 2)       # WARNING
                elif fid == "F004":
                    bbg.append((fid, d - dt.timedelta(days=1), round(v * 1.0005, 2), ccy, "BLOOMBERG", "DAILY", f[9], f[3]))
                    continue                         # DATE DIFFERENCE
            bbg.append((fid, d, val, ccy, "BLOOMBERG", "WEEKLY", f[9], f[3]))
    # 分配落ち: 分配金ぶん価格 NAV を下げる（TR では足し戻される）
    off2 = []
    for row in official:
        fid, d, v = row[0], row[1], row[2]
        cut = sum(x[2] for x in dists if x[0] == fid and x[1] <= d)
        off2.append((fid, d, round(v - cut, 2)) + row[3:])
    bbg2 = []
    for row in bbg:
        fid, d, v = row[0], row[1], row[2]
        cut = sum(x[2] for x in dists if x[0] == fid and x[1] <= d)
        bbg2.append((fid, d, round(v - cut, 2)) + row[3:])
    return off2, bbg2, dists


def holdings():
    off, _, _ = nav_history()
    out = []
    for hid, cid, fid, book, entry, red in HOLDINGS_SPEC:
        cands = [r for r in off if r[0] == fid and r[1] <= entry]
        entry_nav = cands[-1][2]
        ccy = next(f[6] for f in FUNDS if f[0] == fid)
        units = book / entry_nav if ccy == "JPY" else book / USDJPY_ENTRY / entry_nav
        out.append((hid, cid, fid, round(units, 2), book, entry, entry_nav, red, ""))
    return out


ATTRIBUTION = [
    # fund, date, component, value, level, source
    ("F001", EVAL_DATE, "Underlying Market", -0.0305, "CONFIRMED", "運用会社 週次寄与度レポート（サンプル）"),
    ("F001", EVAL_DATE, "Option Premium", 0.0038, "CONFIRMED", "運用会社 週次寄与度レポート（サンプル）"),
    ("F001", EVAL_DATE, "Hedging Costs", -0.0006, "CONFIRMED", "運用会社 週次寄与度レポート（サンプル）"),
    ("F001", EVAL_DATE, "Fees", -0.0002, "CONFIRMED", "運用会社 週次寄与度レポート（サンプル）"),
]

# ------------------------------------------------------------------ market
# key, desc, class, ticker, field, critical, unit, cur, prev, mend, asof, note
MARKET = [
    ("EQ_NDX", "NASDAQ 100", "Equity", "NDX Index", "PX_LAST", "Y", "PRICE", 20050.00, 20850.00, 20600.00, EVAL_DATE, ""),
    ("EQ_SPX", "S&P 500", "Equity", "SPX Index", "PX_LAST", "Y", "PRICE", 6420.50, 6580.25, 6550.10, EVAL_DATE, ""),
    ("EQ_SX5E", "EURO STOXX 50", "Equity", "SX5E Index", "PX_LAST", "N", "PRICE", None, 5410.30, 5450.20, None, "テスト：欠損"),
    ("EQ_NKY", "日経平均", "Equity", "NKY Index", "PX_LAST", "N", "PRICE", 44800.00, 45950.00, 45200.00, EVAL_DATE, ""),
    ("FUT_NQ", "NASDAQ100先物（期近）", "Futures", "NQ1 Index", "PX_LAST", "N", "PRICE", 20110.00, 20905.00, 20650.00, EVAL_DATE, ""),
    ("VOL_VIX", "VIX", "Volatility", "VIX Index", "PX_LAST", "Y", "VOL", 22.40, 18.30, 16.90, EVAL_DATE, ""),
    ("VOL_VXN", "VXN（NDX Vol指数）", "Volatility", "VXN Index", "PX_LAST", "Y", "VOL", 26.80, 21.90, 20.50, EVAL_DATE, ""),
    ("RV_NDX", "NDX 30日実現Vol", "Volatility", "NDX Index", "VOLATILITY_30D", "N", "VOL", 21.50, 17.20, 15.80, EVAL_DATE, "フィールド要確認"),
    ("IV_NDX", "NDX 1M ATM IV", "Volatility", "NDX Index", "30DAY_IMPVOL_100.0%MNY_DF", "N", "VOL", 25.90, 21.00, 19.80, EVAL_DATE, "フィールド要確認"),
    ("RT_UST2", "米国債2年", "Rates", "USGG2YR Index", "PX_LAST", "N", "YIELD", 3.62, 3.70, 3.75, EVAL_DATE, ""),
    ("RT_UST5", "米国債5年", "Rates", "USGG5YR Index", "PX_LAST", "N", "YIELD", 3.78, 3.84, 3.86, EVAL_DATE, ""),
    ("RT_UST10", "米国債10年", "Rates", "USGG10YR Index", "PX_LAST", "N", "YIELD", 4.21, 4.25, 4.18, EVAL_DATE, ""),
    ("RT_JGB2", "国債2年", "Rates", "GJGB2 Index", "PX_LAST", "N", "YIELD", 0.92, 0.91, 0.90, EVAL_DATE, ""),
    ("RT_JGB5", "国債5年", "Rates", "GJGB5 Index", "PX_LAST", "N", "YIELD", 1.21, 1.20, 1.18, EVAL_DATE, ""),
    ("RT_JGB10", "国債10年", "Rates", "GJGB10 Index", "PX_LAST", "N", "YIELD", 1.68, 1.66, 1.63, D(2026, 10, 28), "テスト：STALE"),
    ("SW_USD10", "USD SOFR OIS 10年", "Rates", "USOSFR10 Curncy", "PX_LAST", "N", "YIELD", 3.95, 3.99, 3.93, EVAL_DATE, "Ticker要確認"),
    ("SW_JPY10", "JPY TONA OIS 10年", "Rates", "JYSO10 Curncy", "PX_LAST", "N", "YIELD", 1.35, 1.33, 1.30, EVAL_DATE, "Ticker要確認"),
    ("FX_USDJPY", "ドル円", "FX", "USDJPY Curncy", "PX_LAST", "Y", "FX", 152.30, 151.10, 149.80, EVAL_DATE, ""),
    ("FX_EURJPY", "ユーロ円", "FX", "EURJPY Curncy", "PX_LAST", "N", "FX", 165.20, 164.80, 163.10, EVAL_DATE, ""),
    ("FX_EURUSD", "ユーロドル", "FX", "EURUSD Curncy", "PX_LAST", "N", "FX", 1.0847, 1.0906, 1.0888, EVAL_DATE, ""),
    ("CR_CDXIG", "CDX IG 5Y", "Credit", "CDX IG CDSI GEN 5Y Corp", "PX_LAST", "N", "SPREAD", 58.5, 54.2, 52.0, EVAL_DATE, "Ticker要確認"),
    ("CMD_CL", "WTI原油（期近）", "Commodity", "CL1 Comdty", "PX_LAST", "N", "PRICE", 68.40, 70.10, 69.30, EVAL_DATE, ""),
    ("CMD_GC", "金（期近）", "Commodity", "GC1 Comdty", "PX_LAST", "N", "PRICE", 2650.0, 2620.0, 2580.0, EVAL_DATE, ""),
    ("QIS_NDXPW", "Sample QIS NDX Put Write Index", "QIS Index", "SQNDXPW Index", "PX_LAST", "Y", "PRICE", 1208.31, 1234.56, 1262.00, EVAL_DATE, "架空Ticker"),
    ("QIS_NDXPS", "Sample QIS NDX Put Spread Index", "QIS Index", "SQNDXPS Index", "PX_LAST", "Y", "PRICE", 979.10, 987.65, 990.20, EVAL_DATE, "架空Ticker"),
    ("QIS_VT10", "Sample QIS US Eq VolTarget 10%", "QIS Index", "SQVT10 Index", "PX_LAST", "Y", "PRICE", 1105.40, 1118.80, 1121.50, EVAL_DATE, "架空Ticker"),
    ("QIS_TC", "Sample QIS Multi-Asset Trend & Carry", "QIS Index", "SQTC Index", "PX_LAST", "Y", "PRICE", None, 1050.20, 1060.00, None, "架空Ticker／テスト：欠損"),
]


def series():
    """週次時系列（チャート用）。最後の 2 点は MARKET の prev/cur と一致させる。"""
    rng = random.Random(7)
    keys = ["QIS_NDXPW", "QIS_NDXPS", "QIS_VT10", "QIS_TC", "EQ_NDX", "EQ_SPX"]
    days = list(fridays(D(2025, 6, 27), EVAL_DATE))
    out = []
    for k in keys:
        m = next(x for x in MARKET if x[0] == k)
        cur, prev = m[7], m[8]
        vals = [cur if cur is not None else None, prev]
        v = prev
        for _ in range(len(days) - 2):
            v = v / (1 + 0.002 + rng.gauss(0, 0.012))
            vals.append(v)
        vals = list(reversed(vals))
        for d, v in zip(days, vals):
            if v is None:
                continue  # 欠損はそのまま欠損として残す
            out.append((k, d, round(v, 2), "SAMPLE"))
    return out


VT_INPUT_RAW = [
    # date, target, rv, exposure, min, max, lev, und, strat, rebalance
    (D(2026, 9, 4), 0.10, 0.085, 1.18, 0.0, 1.5, 1.18, 6480.2, 1101.2, "N"),
    (D(2026, 9, 11), 0.10, 0.083, 1.20, 0.0, 1.5, 1.20, 6512.8, 1106.9, "N"),
    (D(2026, 9, 18), 0.10, 0.088, 1.14, 0.0, 1.5, 1.14, 6495.1, 1104.1, "Y"),
    (D(2026, 9, 25), 0.10, 0.090, 1.11, 0.0, 1.5, 1.11, 6533.7, 1109.0, "N"),
    (D(2026, 10, 2), 0.10, 0.092, 1.09, 0.0, 1.5, 1.09, 6561.0, 1112.3, "N"),
    (D(2026, 10, 9), 0.10, 0.091, 1.10, 0.0, 1.5, 1.10, 6570.4, 1113.5, "N"),
    (D(2026, 10, 16), 0.10, 0.094, 1.06, 0.0, 1.5, 1.06, 6549.9, 1111.0, "N"),
    (D(2026, 10, 23), 0.10, 0.096, 1.18, 0.0, 1.5, 1.18, 6580.25, 1118.80, "N"),
    (D(2026, 10, 30), 0.10, 0.138, 0.83, 0.0, 1.5, 0.83, 6420.50, 1105.40, "Y"),
]
VT_INPUT = [(d, "IDX_VT10", *rest, "運用会社 週次レポート（サンプル）") for d, *rest in VT_INPUT_RAW]

TC_INSTR = [
    # asset, instrument, ticker, (sig_prev, sig_cur), (w_prev, w_cur), (carry_prev, carry_cur), roll, ratediff
    ("Equity", "S&P500 E-mini", "ES1 Index", (0.60, -0.20), (0.25, -0.08), (0.012, 0.002), 0.0, None),
    ("Equity", "日経225先物", "NK1 Index", (0.30, 0.10), (0.10, 0.04), (0.008, 0.008), 0.0, None),
    ("Rates", "米10年債先物", "TY1 Comdty", (-0.40, -0.50), (-0.20, -0.25), (0.005, 0.005), 0.0, None),
    ("Rates", "長期国債先物", "JB1 Comdty", (-0.50, -0.55), (-0.15, -0.16), (0.002, 0.002), 0.0, None),
    ("FX", "USDJPY", "USDJPY Curncy", (0.45, 0.40), (0.30, 0.28), (0.045, 0.010), 0.0, 0.034),
    ("Commodity", "WTI原油", "CL1 Comdty", (0.20, 0.25), (0.08, 0.10), (-0.010, -0.010), -0.020, None),
    ("Commodity", "金", "GC1 Comdty", (0.70, 0.65), (0.15, 0.14), (0.0, 0.0), 0.0, None),
]


def tc_input():
    out = []
    for i, d in enumerate((PREV_DATE, EVAL_DATE)):
        for a, inst, tk, sig, w, carry, roll, rd in TC_INSTR:
            out.append((d, "IDX_TC", a, inst, tk, sig[i], w[i], carry[i], roll, rd, None, "", "運用会社 週次ポジション（サンプル）"))
    return out


EVENTS = [
    (D(2026, 11, 3), "（サンプル）日本 祝日：国内市場休場", "JP", "Low"),
    (D(2026, 11, 4), "（サンプル）NDXオプション満期（ポジション該当）", "US", "Medium"),
    (D(2026, 11, 6), "（サンプル）米国 雇用統計", "US", "High"),
    (D(2026, 11, 12), "（サンプル）米国 CPI", "US", "High"),
    (D(2026, 11, 20), "（サンプル）NDX月次オプション満期", "US", "Medium"),
]

HOLIDAYS = [  # 米国市場休場日（サンプル：要更新）
    D(2026, 11, 26), D(2026, 12, 25), D(2027, 1, 1), D(2027, 1, 18), D(2027, 2, 15),
    D(2027, 3, 26), D(2027, 5, 31), D(2027, 6, 18), D(2027, 7, 5), D(2027, 9, 6),
]


def write_csvs(outdir: Path):
    outdir.mkdir(parents=True, exist_ok=True)
    pos_dir = outdir / "positions"
    pos_dir.mkdir(exist_ok=True)
    for fname, rows in positions():
        with open(pos_dir / fname, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=POS_COLUMNS)
            w.writeheader()
            for r in rows:
                rr = {k: (v.strftime("%Y-%m-%d") if isinstance(v, dt.date) else v) for k, v in r.items()}
                w.writerow(rr)
    off, bbg, dists = nav_history()
    with open(outdir / "nav_official.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["Fund ID", "NAV Date", "NAV", "Currency", "Source", "NAV Type", "Basis", "Share Class"])
        for r in off:
            w.writerow([r[0], r[1].isoformat(), *r[2:]])
    with open(outdir / "nav_bloomberg.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["Fund ID", "NAV Date", "NAV", "Currency", "Source", "NAV Type", "Basis", "Share Class"])
        for r in bbg:
            w.writerow([r[0], r[1].isoformat(), *r[2:]])
    with open(outdir / "market_data_manual.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["Key", "Ticker", "Field", "Current", "Prev Week", "Prev Month End", "As-of"])
        for m in MARKET:
            w.writerow([m[0], m[3], m[4], m[7] if m[7] is not None else "", m[8], m[9], m[10].isoformat() if m[10] else ""])


if __name__ == "__main__":
    here = Path(__file__).resolve().parent.parent / "sample_data"
    write_csvs(here)
    print("wrote", here)
