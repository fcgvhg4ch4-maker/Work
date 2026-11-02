#!/usr/bin/env python3
"""生成したブックを LibreOffice で再計算し、Python で独立に計算した期待値と照合する。

  python3 verify_workbook.py            # dist/QIS_Fund_Monitoring_Workbook.xlsx を検証
結果は dist/test_results.json と ../docs/07_test_results.md に書き出す。

基本シナリオ（サンプルデータそのまま）に加え、ブックを一部書き換えた派生シナリオを
再計算して、欠損・重複・承認・担当者切替の挙動を確認する。
"""
from __future__ import annotations

import datetime as dt
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import numpy as np
from openpyxl import load_workbook

import sample_data as SD

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "dist" / "QIS_Fund_Monitoring_Workbook.xlsx"
TESTS = []
ERR = ("#VALUE!", "#NAME?", "#REF!", "#DIV/0!", "#NUM!", "#NULL!", "Err:", "#N/A")


def recalc(src: Path) -> Path:
    out = Path(tempfile.mkdtemp(prefix="qisv_"))
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(out), str(src)],
                   check=True, capture_output=True, timeout=600)
    return out / src.name


def check(tid, scenario, what, expected, actual, ok=None, tol=1e-6):
    if ok is None:
        if isinstance(expected, float) and isinstance(actual, (int, float)):
            ok = abs(expected - actual) <= tol
        else:
            ok = expected == actual
    TESTS.append(dict(id=tid, scenario=scenario, check=what, expected=_s(expected), actual=_s(actual), result="PASS" if ok else "FAIL"))
    return ok


def _s(v):
    if isinstance(v, float):
        return f"{v:.6g}"
    if isinstance(v, (dt.date, dt.datetime)):
        return v.strftime("%Y-%m-%d")
    return str(v)


def d(v):
    return v.date() if isinstance(v, dt.datetime) else v


class WB:
    def __init__(self, path):
        self.wb = load_workbook(path, data_only=True)

    def find_row(self, sheet, col, value, start=6, end=3000):
        ws = self.wb[sheet]
        for r in range(start, end):
            if ws[f"{col}{r}"].value == value:
                return r
        return None

    def v(self, sheet, ref):
        return self.wb[sheet][ref].value

    def col_values(self, sheet, col, start, end):
        ws = self.wb[sheet]
        return [ws[f"{col}{r}"].value for r in range(start, end + 1)]

    def header_map(self, sheet, row):
        ws = self.wb[sheet]
        return {ws.cell(row, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(row, c).value is not None}

    def table(self, sheet, hdr_row, start_col=1, n=None):
        """header -> list of row dicts (until key col empty)"""
        ws = self.wb[sheet]
        heads = {}
        c = start_col
        while ws.cell(hdr_row, c).value is not None:
            heads[ws.cell(hdr_row, c).value] = c
            c += 1
        rows = []
        r = hdr_row + 1
        while True:
            vals = {h: ws.cell(r, ci).value for h, ci in heads.items()}
            if n is not None and len(rows) >= n:
                break
            if all(v in (None, "") for k, v in vals.items() if k not in ("#",)):
                break
            rows.append(vals)
            r += 1
        return rows


# ------------------------------------------------------------------ expected values (independent)
def expected_nav():
    off, bbg, dists = SD.nav_history()
    exp = {}
    for f in SD.FUNDS:
        fid = f[0]
        rows = sorted([r for r in off if r[0] == fid], key=lambda r: r[1])
        tri = {}
        prev = None
        for dte, nav in [(r[1], r[2]) for r in rows]:
            if prev is None:
                tri[dte] = 100.0
            else:
                pd_, pn = prev
                dist = sum(x[2] for x in dists if x[0] == fid and pd_ < x[1] <= dte)
                tri[dte] = tri[pd_] * ((nav + dist) / pn)
            prev = (dte, nav)
        dates = sorted(tri)
        def base(limit):
            c = [x for x in dates if x <= limit]
            return c[-1] if c else None
        cur = SD.EVAL_DATE
        wk = tri[cur] / tri[base(SD.PREV_DATE)] - 1
        mb = base(dt.date(2026, 9, 30))
        mo = tri[cur] / tri[mb] - 1
        peak = max(tri[x] for x in dates)
        b = {r[1]: r[2] for r in bbg if r[0] == fid}
        offnav = dict((r[1], r[2]) for r in rows)[cur]
        if not b:
            st = "MISSING DATA"
        elif cur not in b:
            st = "DATE DIFFERENCE"
        else:
            diff = abs(b[cur] / offnav - 1)
            st = "MATCH" if diff <= 0.0005 else ("WARNING" if diff <= 0.003 else "MISMATCH")
        # realized vol (26 weeks)
        win = [x for x in dates if x > cur - dt.timedelta(days=7 * 26) and x <= cur]
        rets = []
        for x in win:
            i = dates.index(x)
            if i > 0:
                rets.append(tri[x] / tri[dates[i - 1]] - 1)
        vol = float(np.std(rets, ddof=1) * np.sqrt(52)) if len(rets) >= 4 else None
        exp[fid] = dict(weekly=wk, monthly=mo, dd=tri[cur] / peak - 1, status=st, vol=vol, nav=offnav)
    return exp


def expected_options():
    (_, prev), (_, cur) = SD.positions()
    hol = [np.datetime64(h) for h in SD.HOLIDAYS]
    out = {}
    for idx in (SD.IDX_PW, SD.IDX_PS):
        c = [r for r in cur if r["Index Name"] == idx]
        p = [r for r in prev if r["Index Name"] == idx]
        def dist(r):
            return (r["NDX Spot"] - r["Strike"]) / r["NDX Spot"]
        def iv(rows):
            g = sum(abs(r["Notional(USD)"]) for r in rows)
            return sum(r["Implied Vol"] * abs(r["Notional(USD)"]) for r in rows) / g
        gam = lambda rows: sum(r["Gamma"] * r["Notional(USD)"] for r in rows)
        nearest = min(abs(dist(r)) for r in c)
        mat = min(r["Maturity Date"] for r in c)
        bdays = int(np.busday_count(np.datetime64(SD.EVAL_DATE) + np.timedelta64(1, 'D'), np.datetime64(mat) + np.timedelta64(1, 'D'), holidays=hol))
        keys_c = {r["BloombergTicker"]: 0 for r in c}
        keys_p = {r["BloombergTicker"]: 0 for r in p}
        for r in c:
            keys_c[r["BloombergTicker"]] += r["Notional(USD)"]
        for r in p:
            keys_p[r["BloombergTicker"]] += r["Notional(USD)"]
        cls = {}
        for k in set(keys_c) | set(keys_p):
            mat_k = next(r["Maturity Date"] for r in c + p if r["BloombergTicker"] == k)
            if k not in keys_p:
                cls[k] = "New"
            elif k not in keys_c:
                cls[k] = "Expired" if mat_k <= SD.EVAL_DATE else "Closed"
            else:
                ch = (abs(keys_c[k]) - abs(keys_p[k])) / abs(keys_p[k])
                cls[k] = "Unchanged" if abs(ch) <= 0.001 else ("Increased" if ch > 0 else "Reduced")
        def mny(r):
            x = dist(r)
            put = " P" in r["BloombergTicker"]
            if abs(x) <= 0.005:
                return "ATM"
            return ("OTM" if x > 0 else "ITM") if put else ("OTM" if x < 0 else "ITM")
        mp = {r["BloombergTicker"]: mny(r) for r in p}
        mc = {r["BloombergTicker"]: mny(r) for r in c}
        itm = sorted(k for k in mc if k in mp and mc[k] == "ITM" and mp[k] != "ITM")
        out[idx] = dict(nearest=nearest, iv_chg=iv(c) - iv(p), gamma_chg=(gam(c) - gam(p)) / abs(gam(p)), nearest_mat=mat,
                        bdays=bdays, cls=cls, itm=itm, n=len(c), gross=sum(abs(r["Notional(USD)"]) for r in c),
                        types={r["BloombergTicker"]: ("PUT" if " P" in r["BloombergTicker"] else "CALL") for r in c + p})
    return out


# ------------------------------------------------------------------ scenarios
def base_tests(path):
    S = "基本"
    w = WB(path)
    nav = expected_nav()
    pf = {r["Fund ID"]: r for r in w.table("09_Performance", 5)}
    rc = {r["Fund ID"]: r for r in w.table("08_NAV_History", 5, start_col=30)}

    # T01 NAV decline
    check("T01-1", "NAV下落", "F001 週次リターン（TR）", nav["F001"]["weekly"], pf["F001"]["Weekly Return"])
    check("T01-2", "NAV下落", "F001 月次リターン（TR）", nav["F001"]["monthly"], pf["F001"]["Monthly Return"])
    for f in ("F002", "F003", "F004", "F005"):
        check(f"T01-3{f}", "NAV下落", f"{f} 週次リターン（分配金調整後 TR）", nav[f]["weekly"], pf[f]["Weekly Return"])
    check("T01-4", "NAV下落", "F004 現在ドローダウン（TR）", nav["F004"]["dd"], pf["F004"]["現在DD"])
    check("T01-5", "NAV下落", "F001 実現Vol（26週・年率）", nav["F001"]["vol"], pf["F001"]["Realized Vol(年率)"], tol=1e-9)

    # T02 NAV mismatch
    for f in ("F001", "F002", "F003", "F004", "F005"):
        check(f"T02-{f}", "NAV不一致", f"{f} NAV照合結果", nav[f]["status"], rc[f]["照合結果"])

    al = w.table("13_Risk_Alerts", 5, n=200)
    al = [a for a in al if a["Severity"]]
    alerts = {(a["Fund"], a["Alert Type"], a["Index"] or "") for a in al}
    sev = {(a["Fund"], a["Alert Type"], a["Index"] or ""): a["Severity"] for a in al}
    check("T01-6", "NAV下落", "F001 週次NAV下落アラート＝High", "High", sev.get(("F001", "週次NAV下落", "")))
    check("T01-7", "NAV下落", "F001 月次NAV下落アラート＝High", "High", sev.get(("F001", "月次NAV下落", "")))
    check("T01-8", "NAV下落", "F004 ドローダウン拡大アラート＝Medium", "Medium", sev.get(("F004", "ドローダウン拡大", "")))
    check("T02-6", "NAV不一致", "F002 NAV不一致アラート＝High", "High", sev.get(("F002", "NAV不一致", "")))
    check("T02-7", "NAV不一致", "F004 日付ずれは差額アラートを出さず『照合不能』", ("Medium", None),
          (sev.get(("F004", "NAV照合不能", "")), sev.get(("F004", "NAV不一致", ""))))

    # T03-T05 options
    ox = expected_options()
    oa = {r["Index ID"]: r for r in w.table("10_Option_Analytics", 5) if r["Index ID"]}
    ch = [r for r in w.table("10_Option_Analytics", 40) if r["Position Key"]]
    pw = ox[SD.IDX_PW]
    cls = {r["Ticker"]: r["基本分類"] for r in ch if r["Index ID"] == "IDX_NDXPW"}
    check("T03-1", "新規ポジション", "NDX 01/15/27 P19000 → New", "New", cls.get("NDX US 01/15/27 P19000 Index"))
    check("T03-2", "新規ポジション", "NDX 11/04/26 C21000 → New（CALL 判定）", ("New", "CALL"),
          (cls.get("NDX US 11/04/26 C21000 Index"), next(r["Type"] for r in ch if r["Ticker"] == "NDX US 11/04/26 C21000 Index")))
    check("T03-3", "新規ポジション", "全銘柄の分類が独立計算と一致", pw["cls"], cls)
    check("T04-1", "ポジション解消", "NDX 11/06/26 P19600 → Closed", "Closed", cls.get("NDX US 11/06/26 P19600 Index"))
    check("T04-2", "ポジション解消", "NDX 10/30/26 P19800 → Expired（満期到来）", "Expired", cls.get("NDX US 10/30/26 P19800 Index"))
    roll = {r["Ticker"]: r["ロール判定（推定）"] for r in ch}
    check("T04-3", "ポジション解消", "解消＋新満期の新規 → Possible Roll（推定）", ("Possible Roll (旧側)", "Possible Roll (新規側)"),
          (roll.get("NDX US 11/06/26 P19600 Index"), roll.get("NDX US 01/15/27 P19000 Index")))
    check("T04-4", "ポジション解消", "Call の新規はロール候補にしない（Put/Call 別グループ）", None, roll.get("NDX US 11/04/26 C21000 Index"))
    check("T04-5", "ポジション解消", "同一銘柄2取引（open_date 別）の集約：P18500 前回/今回 Notional", (40_000_000, 40_000_000),
          tuple(next((r["Notional(前回)"], r["Notional(今回)"]) for r in ch if r["Ticker"] == "NDX US 12/18/26 P18500 Index")))
    check("T05-1", "Strike接近", "IDX_NDXPW Nearest Strike Distance", pw["nearest"], oa["IDX_NDXPW"]["Nearest Strike Dist"])
    check("T05-2", "Strike接近", "OTM→ITM 変化銘柄", pw["itm"], [r["Ticker"] for r in ch if r["ITM変化"] == "→ITM"])
    check("T05-3", "Strike接近", "Strike接近（3%以内）＝High、5%ルールは抑制", ("High", None),
          (sev.get(("F001", "Strike接近（3%以内）", "IDX_NDXPW")), sev.get(("F001", "Strike接近（5%以内）", "IDX_NDXPW"))))
    check("T05-4", "Strike接近", "ITMへの変化アラート＝High", "High", sev.get(("F001", "ITMへの変化", "IDX_NDXPW")))
    check("T05-5", "Strike接近", "IDX_NDXPS（最接近 5.24%）は Strike アラートなし", False,
          any(k[2] == "IDX_NDXPS" and k[1].startswith("Strike接近") for k in alerts))
    check("T05-6", "Strike接近", "最短満期までの営業日数（休日考慮）", pw["bdays"], oa["IDX_NDXPW"]["満期まで営業日"])
    check("T05-7", "Strike接近", "満期接近（3営業日以内）＝Medium、30日ルールは抑制", ("Medium", None),
          (sev.get(("F001", "満期接近（営業日）", "IDX_NDXPW")), sev.get(("F001", "満期接近（30日以内）", "IDX_NDXPW"))))
    check("T05-8", "Strike接近", "加重平均IV変化（|Notional|加重）", pw["iv_chg"], oa["IDX_NDXPW"]["IV変化"])
    check("T05-9", "Strike接近", "Gamma×Notional 変化率（参考値）", pw["gamma_chg"], oa["IDX_NDXPW"]["Γ変化率"])
    check("T05-10", "Strike接近", "有効ポジション数・Gross Notional", (pw["n"], pw["gross"]),
          (oa["IDX_NDXPW"]["ポジション数"], oa["IDX_NDXPW"]["Gross Notional"]))
    check("T05-11", "Strike接近", "Greeks は単位確認前『参考値』表示", "参考値（単位未確認）", oa["IDX_NDXPW"]["Greeks区分"])
    cp = w.table("10_Option_Analytics", 345)
    cp = [r for r in cp if r["Ticker"]]
    check("T05-12", "Strike接近", "前回ファイルの満期到来ポジション（10/30 満期 P19800）は今回の有効明細に含まれない",
          False, any(r["Ticker"] == "NDX US 10/30/26 P19800 Index" for r in cp))

    # T06 Vol target
    vs = w.table("11_VolTarget_Analytics", 5)[0]
    check("T06-1", "Vol Target Exposure縮小", "Exposure 変化率", 0.83 / 1.18 - 1, vs["Exposure変化率"])
    check("T06-2", "Vol Target Exposure縮小", "Exposure急縮小＝High／実現Vol急上昇＝Medium", ("High", "Medium"),
          (sev.get(("F003", "Exposure急縮小", "IDX_VT10")), sev.get(("F003", "実現Vol急上昇", "IDX_VT10"))))
    check("T06-3", "Vol Target Exposure縮小", "コメントに『観測』としての説明", True, "【観測】実現Vol上昇に伴いExposureが縮小" in (vs["コメント"] or ""))

    # T07 trend
    ts = w.table("12_TrendCarry_Analytics", 5)[0]
    check("T07-1", "Trendシグナル反転", "シグナル反転件数／L/S反転件数", (1, 1), (ts["Signal反転"], ts["L/S反転"]))
    check("T07-2", "Trendシグナル反転", "反転銘柄の特定", "Equity:S&P500 E-mini(+0.60→-0.20)", ts["シグナル反転 銘柄"])
    gp = sum(abs(x[4][0]) for x in SD.TC_INSTR)
    gc = sum(abs(x[4][1]) for x in SD.TC_INSTR)
    cp_ = sum(x[5][0] * abs(x[4][0]) for x in SD.TC_INSTR) / gp
    cc_ = sum(x[5][1] * abs(x[4][1]) for x in SD.TC_INSTR) / gc
    check("T07-3", "Trendシグナル反転", "加重Carry変化", cc_ - cp_, ts["Carry変化"])
    check("T07-4", "Trendシグナル反転", "Trendシグナル反転＝Medium・Carry低下＝Low", ("Medium", "Low"),
          (sev.get(("F004", "Trendシグナル反転", "IDX_TC")), sev.get(("F004", "Carry低下", "IDX_TC"))))

    # T08 bloomberg missing
    md = {r["Key"]: r for r in w.table("07_Bloomberg_Data", 5)}
    check("T08-1", "Bloombergデータ欠損", "SX5E 欠損 → MISSING", "MISSING", md["EQ_SX5E"]["Status"])
    check("T08-2", "Bloombergデータ欠損", "JGB10 As-of 10/28 → STALE（値は採用しない）", ("STALE", None),
          (md["RT_JGB10"]["Status"], md["RT_JGB10"]["検証済 値"]))
    check("T08-3", "Bloombergデータ欠損", "QIS_TC 欠損 → F004 参照指数週次は空欄（前回値で代用しない）", None, pf["F004"]["Linked Index Weekly"])
    check("T08-4", "Bloombergデータ欠損", "市場データ欠損アラート（IDX_TC）", "Medium", sev.get(("F004", "市場データ欠損", "IDX_TC")))
    check("T08-5", "Bloombergデータ欠損", "Add-in 未接続時の BDH は NO_CONN", "NO_CONN", md["EQ_NDX"]["BBG 評価基準日"])
    cm = {r["Fund ID"]: r for r in w.table("15_Weekly_Commentary", 5)}
    check("T08-6", "Bloombergデータ欠損", "コメントに欠損の明示", True, "データ欠損のため未算出" in cm["F004"]["Auto ①Performance Summary"])
    dq = {r["ID"]: r for r in w.table("18_Data_Quality", 7)}
    check("T08-7", "Bloombergデータ欠損", "DQ16（重要市場データ欠損）＝OPEN", ("NG", "OPEN"), (dq["DQ16"]["結果"], dq["DQ16"]["判定"]))

    # T09 multi bank / fund
    sp = [r for r in w.table("14_Sales_Priority", 110) if r["Bank"]]
    check("T09-1", "複数地銀・複数ファンド", "優先度一覧の行数（保有 7 件）", 7, len(sp))
    check("T09-2", "複数地銀・複数ファンド", "先頭は B銀行（NAV不一致・償還接近・接触超過）", ("B銀行", "High"), (sp[0]["Bank"], sp[0]["Priority"]))
    check("T09-3", "複数地銀・複数ファンド", "Priority の内訳（High/Medium）", {"High": 4, "Medium": 3}, {k: sum(1 for r in sp if r["Priority"] == k) for k in ("High", "Medium")})
    a1 = next(a for a in al if a["Fund"] == "F001")
    check("T09-4", "複数地銀・複数ファンド", "アラートの Investor に保有地銀を列挙", "A銀行、D銀行", a1["Investor（保有地銀）"])
    check("T09-5", "複数地銀・複数ファンド", "Fund-Index 1対多（F002 は 2 指数のアラートを持つ）", True,
          any(k[0] == "F002" and k[2] == "IDX_NDXPS" for k in alerts) and any(k[0] == "F002" and k[2] == "IDX_NDXPW" for k in alerts))
    check("T09-6", "複数地銀・複数ファンド", "保有先なしファンド（F005）を DQ20 で検出", ("NG", "F005"), (dq["DQ20"]["結果"], dq["DQ20"]["詳細"]))

    # T10 email
    body = w.v("16_Email_Generator", "B25") or ""
    subj = w.v("16_Email_Generator", "C6")
    check("T10-1", "メール生成", "件名フォーマット", "【QIS Weekly Update】担当先ファンド状況｜2026/11/02", subj)
    check("T10-2", "メール生成", "宛名・担当先（A銀行・B銀行）のみ記載", (True, True, True, False, False),
          (body.startswith("山田 太郎さん"), "銀行名：A銀行" in body, "銀行名：B銀行" in body, "C銀行" in body, "D銀行" in body))
    check("T10-3", "メール生成", "重大DQ未解決 → 本文冒頭に【要確認】・ステータス要確認", (True, True),
          ("【要確認】" in body, (w.v("16_Email_Generator", "C7") or "").startswith("要確認")))
    i_b, i_a = body.find("銀行名：B銀行"), body.find("銀行名：A銀行")
    check("T10-4", "メール生成", "アラートのある先が上位（B銀行→A銀行の順）", True, 0 <= i_b < i_a)
    check("T10-5", "メール生成", "必須セクションの存在", True, all(s in body for s in ["【今週のサマリー】", "【担当先別ファンド状況】", "【想定される顧客質問】", "【来週の注目事項】", "以上、よろしくお願いします。"]))
    check("T10-6", "メール生成", "社内向け質問（（社内）〜）は顧客質問欄に出さない", False, "Q（ファンドB）：（社内）" in body)
    check("T10-7", "メール生成", "顧客転送不可の注記", True, "顧客への転送・転載は行わないでください" in body)
    check("T10-8", "メール生成", "本文の長さ（1分で読める目安：3,000字以内）", True, len(body) <= 3000, ok=len(body) <= 3000)

    # formula health
    errs = []
    for ws in w.wb.worksheets:
        if ws.title in ("01_Dashboard",):
            rng = ws.iter_rows(min_row=1, max_row=40)
        else:
            rng = ws.iter_rows()
        for row in rng:
            for c in row:
                if isinstance(c.value, str) and c.value.startswith(ERR):
                    errs.append(f"{ws.title}!{c.coordinate}={c.value}")
    check("H-1", "数式健全性", "全シートにエラー値（#VALUE! 等）がない（チャート補助の #N/A を除く）", 0, len(errs), ok=len(errs) == 0)
    if errs:
        print("errors:", errs[:30])
    return w


def modify(src: Path, fn) -> Path:
    tmp = Path(tempfile.mkdtemp(prefix="qism_")) / src.name
    wb = load_workbook(src)
    fn(wb)
    wb.save(tmp)
    return recalc(tmp)


def variant_tests(src: Path):
    # V1: re-import of the same position file (without the VBA guard) must be flagged
    def dup(wb):
        ws = wb["06_Position_History"]
        last = max(r for r in range(6, 1506) if ws[f"D{r}"].value is not None)
        rows = [r for r in range(6, last + 1) if d(ws[f"D{r}"].value) == SD.EVAL_DATE]
        for i, r in enumerate(rows):
            for c in range(1, 21):
                ws.cell(last + 1 + i, c).value = ws.cell(r, c).value
    w = WB(modify(src, dup))
    dq = {r["ID"]: r for r in w.table("18_Data_Quality", 7)}
    check("V1-1", "同一ファイル再取込", "重複行を DQ11 で検出（Critical）", ("NG", 18), (dq["DQ11"]["結果"], dq["DQ11"]["件数"]))
    check("V1-2", "同一ファイル再取込", "メールは『要確認』", "要確認", w.v("18_Data_Quality", "C4"))

    # V2: Bloomberg mode without a connection: nothing is reused as latest
    def bbg(wb):
        ws = wb["19_Settings"]
        ws["B6"] = "BLOOMBERG"
    w = WB(modify(src, bbg))
    md = w.table("07_Bloomberg_Data", 5)
    check("V2-1", "Bloomberg 未接続で BLOOMBERG モード", "全市場データが MISSING（手入力値・古い値を使わない）", True, all(r["Status"] == "MISSING" for r in md))
    pf = {r["Fund ID"]: r for r in w.table("09_Performance", 5)}
    check("V2-2", "Bloomberg 未接続で BLOOMBERG モード", "参照指数週次は全ファンド空欄", True, all(pf[f]["Linked Index Weekly"] in (None, "") for f in pf))
    check("V2-3", "Bloomberg 未接続で BLOOMBERG モード", "NAV（公式）由来のリターンは影響なし", True, isinstance(pf["F001"]["Weekly Return"], float))

    # V3: official NAV missing for F001 on eval date → no stale substitution
    def nonav(wb):
        ws = wb["08_NAV_History"]
        for r in range(6, 2006):
            if ws[f"A{r}"].value == "F001" and ws[f"E{r}"].value == "OFFICIAL" and d(ws[f"B{r}"].value) == SD.EVAL_DATE:
                ws[f"B{r}"].value = None
                for c in "ACDEFGHI":
                    ws[f"{c}{r}"].value = None
    w = WB(modify(src, nonav))
    pf = {r["Fund ID"]: r for r in w.table("09_Performance", 5)}
    dq = {r["ID"]: r for r in w.table("18_Data_Quality", 7)}
    cm = {r["Fund ID"]: r for r in w.table("15_Weekly_Commentary", 5)}
    check("V3-1", "公式NAV欠損", "週次リターンは空欄（前週 NAV で代用しない）", None, pf["F001"]["Weekly Return"])
    check("V3-2", "公式NAV欠損", "DQ01 で検出", ("NG", "F001"), (dq["DQ01"]["結果"], dq["DQ01"]["詳細"]))
    check("V3-3", "公式NAV欠損", "コメントは【データ欠損】を明示", True, cm["F001"]["Auto ①Performance Summary"].startswith("【データ欠損】"))
    def nonav_late(wb):
        nonav(wb)
        wb["00_Control"]["C12"] = dt.date(2026, 11, 4)
    w = WB(modify(src, nonav_late))
    al = [a for a in w.table("13_Risk_Alerts", 5, n=200) if a["Severity"]]
    check("V3-4", "公式NAV欠損", "公表予定日（WORKDAY+1）超過で NAV公表遅延アラート", True, any(a["Fund"] == "F001" and a["Alert Type"] == "NAV公表遅延" for a in al))

    # V4: acknowledge critical DQ → email becomes reviewable
    def ack(wb):
        ws = wb["18_Data_Quality"]
        for r in range(8, 32):
            if ws[f"D{r}"].value == "Critical":
                ws[f"H{r}"] = "Y"
                ws[f"I{r}"] = "テスト確認者"
                ws[f"J{r}"] = "（テスト）確認済み"
    w = WB(modify(src, ack))
    check("V4-1", "DQ 承認", "Critical を Ack（確認者記入）→ メール『レビュー可』", ("OK", "レビュー可"),
          (w.v("18_Data_Quality", "C4"), w.v("16_Email_Generator", "C7")))
    check("V4-2", "DQ 承認", "本文の【要確認】行が消える", False, "【要確認】" in (w.v("16_Email_Generator", "B25") or ""))

    # V5: switch rep
    def rep(wb):
        wb["16_Email_Generator"]["C4"] = "佐藤 花子"
    w = WB(modify(src, rep))
    body = w.v("16_Email_Generator", "B25") or ""
    check("V5-1", "担当者切替", "佐藤 花子 → C銀行・E銀行のみ", (True, True, True, False, False),
          (body.startswith("佐藤 花子さん"), "銀行名：C銀行" in body, "銀行名：E銀行" in body, "銀行名：A銀行" in body, "銀行名：B銀行" in body))
    check("V5-2", "担当者切替", "宛先メールアドレス", "hanako.sato@example.com", w.v("16_Email_Generator", "C5"))

    # V6: manual override of priority and commentary
    def ovr(wb):
        ws = wb["14_Sales_Priority"]
        col = next(c for c in range(1, ws.max_column + 1) if ws.cell(5, c).value == "Override(手動)")
        ws.cell(10, col).value = "Low"   # 5th holding row (H005 C銀行 F004)
        ws2 = wb["15_Weekly_Commentary"]
        col = next(c for c in range(1, ws2.max_column + 1) if ws2.cell(5, c).value == "Override ④Risk Watch")
        ws2.cell(6, col).value = "（テスト）手動修正したリスクコメント"
    w = WB(modify(src, ovr))
    sp = {r["Holding ID"]: r for r in w.table("14_Sales_Priority", 5) if r["Holding ID"]}
    check("V6-1", "手動変更", "Override 列で優先度を変更・Reason に（手動変更）", ("Low", True), (sp["H005"]["Priority"], "（手動変更）" in sp["H005"]["Reason（スコア内訳）"]))
    cm = {r["Fund ID"]: r for r in w.table("15_Weekly_Commentary", 5)}
    check("V6-2", "手動変更", "Override したコメントが Final に反映", "（テスト）手動修正したリスクコメント", cm["F001"]["Final ④Risk Watch"])

    # V7: threshold change in Settings propagates
    def thr(wb):
        ws = wb["19_Settings"]
        for r in range(6, 60):
            if ws[f"A{r}"].value == "WeeklyNavDecline":
                ws[f"B{r}"] = -0.03
    w = WB(modify(src, thr))
    al = [a for a in w.table("13_Risk_Alerts", 5, n=200) if a["Severity"]]
    check("V7-1", "閾値変更", "週次下落閾値を -3% に変更 → F001（-2.80%）のアラートが消える", False, any(a["Fund"] == "F001" and a["Alert Type"] == "週次NAV下落" for a in al))


def write_reports(results):
    (ROOT / "dist").mkdir(exist_ok=True)
    (ROOT / "dist" / "test_results.json").write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")
    lines = ["# テスト結果（架空データ）", "",
             f"- 実行日時: {results['run_at']}",
             f"- 合計 {results['total']} 件 / 合格 {results['passed']} 件 / 不合格 {results['failed']} 件",
             "- 方法: `build/verify_workbook.py` がブックを LibreOffice（ヘッドレス）で再計算し、`sample_data.py` から Python で独立に計算した期待値と照合。"
             "派生シナリオ（V*）はブックの入力セルを書き換えて再計算。",
             "- 注意: LibreOffice での再計算結果であり、Excel 365 での再計算は本番前に別途確認すること（docs/08）。VBA・Power Query・Bloomberg 接続はこの検証の対象外。", "",
             "| ID | シナリオ | 確認内容 | 期待値 | 結果値 | 判定 |", "|---|---|---|---|---|---|"]
    for t in results["tests"]:
        esc = lambda x: str(x).replace("|", "\\|").replace("\n", " ")
        lines.append(f"| {t['id']} | {esc(t['scenario'])} | {esc(t['check'])} | {esc(t['expected'])} | {esc(t['actual'])} | {t['result']} |")
    (ROOT / "docs" / "07_test_results.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def main():
    src = XLSX
    calc = recalc(src)
    base_tests(calc)
    variant_tests(src)
    passed = sum(1 for t in TESTS if t["result"] == "PASS")
    results = dict(run_at=dt.datetime.now().strftime("%Y-%m-%d %H:%M"), total=len(TESTS), passed=passed,
                   failed=len(TESTS) - passed, tests=TESTS)
    write_reports(results)
    for t in TESTS:
        if t["result"] != "PASS":
            print("FAIL", t)
    print(f"{passed}/{len(TESTS)} passed")
    shutil.rmtree(calc.parent, ignore_errors=True)
    return 0 if passed == len(TESTS) else 1


if __name__ == "__main__":
    sys.exit(main())
