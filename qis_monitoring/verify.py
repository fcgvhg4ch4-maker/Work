#!/usr/bin/env python3
"""Run the 16 mandatory tests against LibreOffice-recalculated copies of the workbook.

Expected values are computed independently in Python from the synthetic fixtures.
Writes docs/test_results.md. Exit code 1 if any test fails.
"""
import re
import sys
import tempfile
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

import build_workbook as BW
from fixtures import EVAL, fixture_data
from recalc import recalc

ROOT = Path(__file__).parent
TMP = Path(tempfile.mkdtemp(prefix="qis_verify_"))
results = []   # (no, title, ok, detail)


def build_and_calc(tag, **over):
    src = TMP / f"{tag}.xlsx"
    BW.build(src, fixture_data(**over))
    return src, load_workbook(recalc(src, TMP / f"calc_{tag}"), data_only=True)


def table(wb, sheet, name):
    ws = wb[sheet]
    t = ws.tables[name]
    c1, r1, c2, r2 = range_boundaries(t.ref)
    hdr = [ws.cell(r1, c).value for c in range(c1, c2 + 1)]
    return [dict(zip(hdr, [ws.cell(r, c).value for c in range(c1, c2 + 1)])) for r in range(r1 + 1, r2 + 1)]


def nm(wb, name):
    dn = wb.defined_names[name]
    sheet, ref = list(dn.destinations)[0]
    return wb[sheet][ref.replace("$", "")].value


def d(x):
    return x.date() if isinstance(x, datetime) else x


def close(a, b, tol=1e-9):
    return isinstance(a, (int, float)) and isinstance(b, (int, float)) and abs(a - b) <= tol


def record(no, title, checks):
    bad = [c for c in checks if not c[0]]
    results.append((no, title, not bad, f"{len(checks) - len(bad)}/{len(checks)} assertions" +
                    ("" if not bad else " — FAILED: " + "; ".join(c[1] for c in bad[:5]))))


fx = fixture_data()
base_src, wb = build_and_calc("base")

# 1 -------------------------------------------------------------------------
fund_rows = table(wb, "01_Live_Fund_Master", "tbl_Fund")
nav_rows = [r for r in table(wb, "07_NAV_Calc", "tbl_NavCalc") if r["fund_code"]]
c = [([r["fund_code"] for r in fund_rows] == [f["fund_code"] for f in fx["funds"]], "01 fund_code order/values"),
     ([r["fund_code"] for r in nav_rows] == [f["fund_code"] for f in fx["funds"]], "07 mirrors 01"),
     ([r["dup_flag"] for r in fund_rows].count("DUPLICATE") == 2, "F005 duplicate detected (2 rows kept)")]
record(1, "Live Fund の fund_code が元データと一致", c)

# 2 -------------------------------------------------------------------------
exp = {}
bad_keys = set()
for h in fx["holdings"]:
    k = f'{h["end_investor_bdr_id"]}|{h["fund_code"]}'
    if isinstance(h["amount"], (int, float)):
        exp[k] = exp.get(k, 0) + h["amount"]
    else:
        bad_keys.add(k)
        exp.setdefault(k, 0)
agg = {r["holding_key"]: r for r in table(wb, "02_End_Investor_Holdings", "tbl_HoldAgg") if r["holding_key"]}
c = [(set(agg) == set(exp), "same holding_key set")]
for k, v in exp.items():
    if k in bad_keys:
        c.append((agg[k]["amount_total"] is None and agg[k]["agg_status"].startswith("Check Required"), f"{k} invalid amount flagged"))
    else:
        c.append((agg[k]["amount_total"] == v, f"{k} amount {agg[k]['amount_total']} != {v}"))
c.append((agg["B001|F001"]["source_rows"] == 2, "B001|F001 aggregated from 2 rows"))
c.append((len(table(wb, "02_End_Investor_Holdings", "tbl_HoldRaw")) == len(fx["holdings"]), "raw rows all retained"))
record(2, "Investor Holdings の Amount 集約", c)

# 3 -------------------------------------------------------------------------
nc = {r["fund_code"]: r for r in nav_rows}
fm = {r["fund_code"]: r for r in table(wb, "03_Strategy_Map", "tbl_FundMap") if r["fund_code"]}
c = [(nc["F001"]["index_name"] == nc["F002"]["index_name"] == nc["F006"]["index_name"], "3 funds share index"),
     (fm["F001"]["resolved_strategy_data_key"] == "STEP_NKY_A", "F001 -> A (explicit)"),
     (fm["F002"]["resolved_strategy_data_key"] == "STEP_NKY_B", "F002 -> B (explicit)"),
     (fm["F006"]["resolved_strategy_data_key"] in (None, "") and fm["F006"]["mapping_status"].startswith("Review Required"),
      "F006 ambiguous index not auto-picked"),
     (not close(nc["F001"]["ret_1W"], nc["F002"]["ret_1W"]), "NAV returns not shared"),
     (nc["F001"]["nav"] != nc["F002"]["nav"] and nc["F001"]["aum"] != nc["F002"]["aum"], "NAV/AUM fund-specific"),
     (fm["F001"]["multiplier_used"] is None and fm["F002"]["multiplier_used"] == 1.5, "multiplier only when confirmed")]
record(3, "同じ Index の複数ファンドを区別", c)

# 4 -------------------------------------------------------------------------
raw_cols = [x.name for x in [] ]  # placeholder
c = []
for sheet, tname, key in [("06_Position_Raw", "tbl_PosRaw", "positions"), ("04_NAV_Input", "tbl_NAV", "nav"),
                          ("02_End_Investor_Holdings", "tbl_HoldRaw", "holdings")]:
    rows = table(wb, sheet, tname)
    mism = 0
    for got, src in zip(rows, fx[key]):
        for k, v in src.items():
            g = got.get(k)
            if isinstance(v, (date, datetime)):
                ok = d(g) == d(v) if not isinstance(v, datetime) or v.hour == 0 else g == v
            elif v is None:
                ok = g is None
            else:
                ok = g == v or close(g, v)
            mism += not ok
    c.append((mism == 0 and len(rows) == len(fx[key]), f"{tname}: {mism} mismatches"))
# formulas never written into raw columns of the source (generated) workbook
wbf = load_workbook(base_src)
ws = wbf["06_Position_Raw"]
t = ws.tables["tbl_PosRaw"]
c1, r1, c2, r2 = range_boundaries(t.ref)
raw_names = set(fx["positions"][0]) | {"leg_id", "underlying_raw", "put_call_raw", "strike_raw", "maturity_raw"}
formula_in_raw = sum(1 for col in range(c1, c2 + 1) if ws.cell(r1, col).value in raw_names
                     for r in range(r1 + 1, r2 + 1) if isinstance(ws.cell(r, col).value, str) and ws.cell(r, col).value.startswith("="))
c.append((formula_in_raw == 0, f"{formula_in_raw} formulas inside raw columns"))
record(4, "Position Raw の元データが改変されない", c)

# 5 -------------------------------------------------------------------------
pc_rows = [r for r in table(wb, "08_Position_Calc", "tbl_PosCalc") if r["src_row"]]
pat = re.compile(r"^(.*) (\d{2})/(\d{2})/(\d{2}) ([PC])([\d.]+) (\w+)$")
c = []
for r in pc_rows:
    m = pat.match(r["ticker"])
    if m:
        c.append((r["pc_parsed"] == m.group(5) and close(r["strike_parsed"], float(m.group(6))) and
                  d(r["expiry_parsed"]) == date(2000 + int(m.group(4)), int(m.group(2)), int(m.group(3))) and
                  r["undl_parsed"] == f"{m.group(1)} {m.group(7)}", f"parse {r['ticker']}"))
    else:
        c.append((r["parse_ok"] in (False, None) and r["pc_parsed"] is None, f"unparseable {r['ticker']} flagged"))
record(5, "BloombergTicker の Put/Call・Strike 解析", c)

# 6 -------------------------------------------------------------------------
spot = {(m["series_id"], m["data_date"]): m["value"] for m in fx["market"]}
c = []
for r in pc_rows:
    if r["calc_status"] == "OK":
        s = r["spot_used"]
        k = r["strike_used"]
        ser = "SPX Index" if r["undl_used"] == "SPX US Index" else r["undl_used"]
        c.append((close(s, spot[(ser, d(r["position_date"]))]), f"spot {r['ticker']}"))
        c.append((close(r["strike_distance"], k / s - 1), f"distance {r['ticker']}"))
        itm = (s < k) if r["pc_used"] == "P" else (s > k)
        c.append((r["moneyness"] == ("ITM" if itm else "OTM"), f"moneyness {r['ticker']} {r['moneyness']}"))
        c.append((r["buffer_display"].endswith(f"({r['moneyness']})"), "buffer shows ITM/OTM"))
mm = [r for r in pc_rows if r["ticker"] == "NKY 12/11/26 P37000 Index" and r["snap_role"] == "Current"][0]
c.append((mm["calc_status"].startswith("STOP") and mm["strike_distance"] is None, "underlying mismatch stops calc"))
call = [r for r in pc_rows if r["ticker"] == "NKY 12/11/26 C45000 Index" and r["snap_role"] == "Current"][0]
c.append((call["moneyness"] == "OTM", "call OTM when spot<strike"))
record(6, "STEP の Strike Distance と ITM 判定", c)

# 7 -------------------------------------------------------------------------
_, wb7 = build_and_calc("spot_missing", drop_market="SPX Index")
r7 = [r for r in table(wb7, "08_Position_Calc", "tbl_PosCalc") if r["ticker"] == "SPX US 12/18/26 P5000 Index"][0]
cov7 = {r["fund_code"]: r for r in table(wb7, "11_Update_Coverage", "tbl_Cov") if r["fund_code"]}
c = [(r7["calc_status"] == "Spot Missing", f"status {r7['calc_status']}"),
     (r7["strike_distance"] is None and r7["moneyness"] is None, "no distance/moneyness"),
     (cov7["F002"]["data_gate"] == "DATA REQUIRED", "F002 gate DATA REQUIRED"),
     (nm(wb7, "chk_FinalGate") == "DATA REQUIRED", "final gate DATA REQUIRED")]
record(7, "Spot 欠損時に Strike Distance を計算しない", c)

# 8 -------------------------------------------------------------------------
_, wb8 = build_and_calc("nav_missing", sel_fund="F004", sel_investor="B001")
n8 = {r["fund_code"]: r for r in table(wb8, "07_NAV_Calc", "tbl_NavCalc") if r["fund_code"]}["F004"]
c = [(all(n8[f"ret_{p}"] is None for p in ["1D", "1W", "MTD", "3M", "YTD", "1Y", "SI"]), "F004 returns blank"),
     (n8["nav_status"] == "Missing", "status Missing"),
     ("【データ不足】" in (nm(wb8, "cmt_Points") or ""), "comment says データ不足"),
     (nm(wb8, "dsh_Ret_1W") == "—", "dashboard shows —"),
     (nc["F003"]["ret_1Y"] is None and nc["F003"]["ret_3M"] is None, "F003 short history -> 3M/1Y blank"),
     (nm(wb8, "chk_FinalGate") == "DATA REQUIRED", "gate DATA REQUIRED")]
# independent check of weekly return (7 calendar days, nearest prior valid NAV) for F001 & F002
for f in ("F001", "F002"):
    s = {x["nav_date"]: x["nav"] for x in fx["nav"] if x["fund_code"] == f}
    end = max(k for k in s if k <= EVAL)
    base = max(k for k in s if k <= end - timedelta(days=7))
    c.append((close(nc[f]["ret_1W"], s[end] / s[base] - 1, 1e-12), f"{f} 1W return"))
    ytdb = max(k for k in s if k <= date(end.year - 1, 12, 31))
    c.append((close(nc[f]["ret_YTD"], s[end] / s[ytdb] - 1, 1e-12), f"{f} YTD return"))
    peak, mdd = 0, 0
    for k in sorted(s):
        if k > end:
            break
        peak = max(peak, s[k])
        mdd = min(mdd, s[k] / peak - 1)
    c.append((close(nc[f]["max_dd"], mdd, 1e-12), f"{f} max DD"))
    c.append((close(nc[f]["cur_dd"], s[end] / peak - 1, 1e-12), f"{f} current DD"))
record(8, "NAV 欠損時に架空のリターンを表示しない（＋リターン計算の独立検証）", c)

# 9 -------------------------------------------------------------------------
chg = {(r["position_key"], r["snap_role"]): r for r in table(wb, "09_Position_Change", "tbl_PosChg") if r["src"]}
A, B = "STEP_NKY_A", "STEP_NKY_B"
expect = {
    (f"{A}|NKY 11/13/26 P42000 Index|2026-08-14", "Current"): "Existing",
    (f"{A}|NKY 01/15/27 P36000 Index|2026-10-28", "Current"): "New",
    (f"{A}|NKY 10/23/26 P39000 Index|2026-07-24", "Previous"): "Matured / Removed (Matured)",
    (f"{A}|NKY 12/11/26 P35000 Index|2026-09-11", "Previous"): "Matured / Removed (Removed)",
    (f"{A}|NKY 12/11/26 P38000 Index|2026-09-11", "Current"): "Check Required",
    (f"{B}|NKY 12/11/26 P36000 Index|2026-09-11", "Current"): "Existing",
    (f"{B}|SPX US 12/18/26 P5000 Index|2026-10-29", "Current"): "New",
}
c = [(chg.get(k, {}).get("change_type") == v, f"{k[0]}: {chg.get(k, {}).get('change_type')} != {v}") for k, v in expect.items()]
c.append((close(chg[(f"{B}|NKY 12/11/26 P36000 Index|2026-09-11", "Current")]["weight_chg"], 0.05), "weight diff 0.55-0.50"))
c.append((close(chg[(f"{B}|NKY 11/13/26 P35000 Index|2026-08-14", "Current")]["weight_chg"], -0.2), "weight diff 0.30-0.50"))
c.append((chg[(f"{A}|NKY 01/15/27 P36000 Index|2026-10-28", "Current")]["roll_flag"] == "Potential Roll (要確認)",
          "roll only as Potential"))
c.append((not any(r["file_name"] == "comp_A_20261030_v1_TEST.csv" and r["use_row"] == 1 for r in pc_rows), "superseded v1 unused"))
record(9, "ポジション差分の計算", c)

# 10 ------------------------------------------------------------------------
cov = {r["fund_code"]: r for r in table(wb, "11_Update_Coverage", "tbl_Cov") if r["fund_code"]}
c = [(cov["F004"]["coverage_status"] == "Missing" and "NAV" in cov["F004"]["missing_items"], "F004 NAV missing"),
     ("Position" in cov["F003"]["missing_items"], "F003 position missing"),
     ("NAV(遅延)" in cov["F001"]["missing_items"], "F001 stale NAV"),
     (cov["F002"]["coverage_status"] == "Complete", "F002 complete"),
     (all(cov[f]["data_gate"] != "READY" for f in ("F001", "F003", "F004", "F005", "F006")), "incomplete funds never READY")]
alc = {r["fund_code"]: r for r in table(wb, "10_Alert", "tbl_AlertCalc") if r["fund_code"]}
c.append((alc["F004"]["overall_status"] != "GREEN", "missing data never GREEN"))
record(10, "データ欠損が Update Coverage に反映", c)

# 11 ------------------------------------------------------------------------
def sev(wbx, f, code):
    return [r for r in table(wbx, "10_Alert", "tbl_Alert") if r["fund_code"] == f and r["alert_code"] == code][0]["severity"]


_, wb11 = build_and_calc("thr", thresholds={"thr_Buf_Low": 0.05, "thr_DD_High": 0.20, "thr_Mat_Low": 10})
c = [(sev(wb, "F002", "A09") == "Low" and sev(wb11, "F002", "A09") == "OK", "A09 Low->OK when buffer thr 10%->5%"),
     (sev(wb, "F001", "A04") == "High" and sev(wb11, "F001", "A04") == "Medium", "A04 High->Medium when DD High 10%->20%"),
     (sev(wb, "F002", "A11") == "Low" and sev(wb11, "F002", "A11") == "OK", "A11 Low->OK when maturity thr 30->10")]
record(11, "アラート閾値の変更が反映", c)

# 12 ------------------------------------------------------------------------
fmt = lambda v: f"{v * 100:+.2f}%"
pts = nm(wb, "cmt_Points")
fund = nm(wb, "cmt_Fund")
c = [(fmt(nm(wb, "dsh_Ret_1W")) in pts, "1W in points"), (fmt(nm(wb, "dsh_Ret_YTD")) in pts, "YTD in points"),
     (fmt(nm(wb, "dsh_Ret_MTD")) in fund and fmt(nm(wb, "dsh_Ret_1Y")) in fund, "MTD/1Y in fund status"),
     (f"{nm(wb, 'dsh_MinBuf') * 100:.1f}%" in nm(wb, "cmt_Position"), "min buffer"),
     (f"レッグ {nm(wb, 'dsh_NLegs')}本" in nm(wb, "cmt_Position"), "leg count"),
     (nm(wb, "dsh_Ret_1W") == nc["F002"]["ret_1W"], "dashboard == 07")]
wsc = wbf["14_Comment_Output"]
direct = [cl.coordinate for row in wsc.iter_rows(min_row=7, max_row=16) for cl in row
          if isinstance(cl.value, str) and re.search(r"tbl_NavCalc\[(ret|nav|cur_dd|max_dd|aum)", cl.value)]
c.append((not direct, f"comment metrics reference Dashboard names only ({direct})"))
record(12, "Dashboard と Comment Output の数値一致", c)

# 13 ------------------------------------------------------------------------
ws99 = wb["99_Check"]
rows = [[ws99.cell(r, k).value for k in range(1, 13)] for r in range(13, 60) if ws99.cell(r, 1).value]
nonok = [r for r in rows if r[6] in ("DATA REQUIRED", "REVIEW REQUIRED")]
c = [(len(nonok) > 0, "issues detected"),
     (all(r[7] and r[9] for r in nonok), "every issue has reason + action"),
     (all(r[10] and r[11] for r in rows), "owner + last checked present"),
     (any(r[0] == "C03" and "F005" in (r[8] or "") for r in rows), "dup fund listed"),
     (any(r[0] == "C18" and "P37000" in (r[8] or "") for r in rows), "STOP leg listed")]
record(13, "99_Check でエラー原因を確認できる", c)

# 14 ------------------------------------------------------------------------
_, wb14 = build_and_calc("unapproved", review_status="Reviewed")
det = {r["holding_key"]: r for r in table(wb14, "12_Distribution_List", "tbl_DistDetail") if r["holding_key"]}
c = [(det["B001|F002"]["gate"] == "REVIEW REQUIRED", f"detail gate {det['B001|F002']['gate']}"),
     (nm(wb14, "chk_FinalGate") == "REVIEW REQUIRED", "final gate REVIEW REQUIRED"),
     (nm(wb14, "cmt_Final") in (None, ""), "final_text empty"),
     (nm(wb, "chk_FinalGate") == "READY", "approved base case READY"),
     ({r["holding_key"]: r for r in table(wb, "12_Distribution_List", "tbl_DistDetail") if r["holding_key"]}["B002|F002"]["gate"]
      == "DATA REQUIRED", "missing e-mail -> DATA REQUIRED")]
record(14, "未承認の文章が配信可能と判定されない", c)

# 15 ------------------------------------------------------------------------
extra_h = [dict(as_of_date=EVAL, end_investor_bdr_id="B005", end_investor_name="TEST Bank Epsilon", fund_code="F002",
                amount=5_000_000, currency="JPY", sales_rep="TEST Rep", source_file="added_row_TEST", source_row=99)]
extra_n = [dict(fund_code="F003", nav_date=date(2026, 10, 31), nav=2000.0, source="TEST")]  # after EVAL: must be ignored
_, wb15 = build_and_calc("addrow", holdings_extra=extra_h, nav_extra=extra_n)
agg15 = {r["holding_key"]: r for r in table(wb15, "02_End_Investor_Holdings", "tbl_HoldAgg") if r["holding_key"]}
f15 = {r["fund_code"]: r for r in table(wb15, "01_Live_Fund_Master", "tbl_Fund")}
n15 = {r["fund_code"]: r for r in table(wb15, "07_NAV_Calc", "tbl_NavCalc") if r["fund_code"]}
c = [(agg15.get("B005|F002", {}).get("amount_total") == 5_000_000, "new holding aggregated"),
     (f15["F002"]["investor_count"] == 3, "investor_count updated"),
     (d(n15["F003"]["nav_date"]) == EVAL, "NAV after eval date ignored")]
# every formula column of every table carries a calculatedColumnFormula (Excel auto-fill on new rows)
with zipfile.ZipFile(base_src) as z:
    tx = [z.read(n).decode() for n in z.namelist() if n.startswith("xl/tables/")]
nform = sum(x.count("<calculatedColumnFormula") for x in tx)
ncols = sum(1 for sh in wbf for t in sh.tables.values()
            for col in range(range_boundaries(t.ref)[0], range_boundaries(t.ref)[2] + 1)
            if isinstance(sh.cell(range_boundaries(t.ref)[1] + 1, col).value, str)
            and str(sh.cell(range_boundaries(t.ref)[1] + 1, col).value).startswith("="))
arr_cols = 2  # tbl_Dist fund_list / update_target_funds are ArrayFormula objects
c.append((nform == ncols + arr_cols, f"calculatedColumnFormula {nform} vs formula columns {ncols}+{arr_cols}"))
record(15, "テーブルに行を追加しても数式が機能", c)

# 16 ------------------------------------------------------------------------
re_wb = load_workbook(recalc(base_src, TMP / "reopen"))  # LibreOffice save -> reopen with formulas
errs = []
for sh in wb:
    for row in sh.iter_rows():
        for cl in row:
            if isinstance(cl.value, str) and cl.value.startswith("#") and not (sh.title == "13_Dashboard" and cl.value == "#N/A"):
                errs.append(f"{sh.title}!{cl.coordinate}={cl.value}")
names = set(wbf.defined_names.keys())
c = [(not errs, f"error cells: {errs[:5]}"),
     (all(n in re_wb.defined_names for n in ("ctl_EvalDate", "chk_FinalGate", "dsh_Ret_1W", "thr_Buf_Low")), "defined names kept"),
     (all(len(re_wb[s].tables) == len(wbf[s].tables) for s in wbf.sheetnames), "tables kept"),
     (sum(1 for sh in re_wb for row in sh.iter_rows() for cl in row if isinstance(cl.value, str) and cl.value.startswith("="))
      > 0.95 * sum(1 for sh in wbf for row in sh.iter_rows() for cl in row if isinstance(cl.value, str) and cl.value.startswith("=")),
      "formulas still present after save/reopen"),
     (len(names) > 100, f"{len(names)} defined names")]
record(16, "再オープン後も数式参照が維持", c)

# ---------------------------------------------------------------- report
ok_all = all(r[2] for r in results)
lines = ["# テスト結果（自動生成：verify.py）", "",
         f"- 実行日時: {datetime.now():%Y-%m-%d %H:%M}",
         "- 再計算エンジン: LibreOffice（headless）。Excel 365 本体での再計算は未実施（下記注記参照）。",
         "- 入力: fixtures.py の架空テストデータ（TEST）", "",
         "| # | テスト | 結果 | 詳細 |", "|---|---|---|---|"]
for no, title, ok, detail in results:
    lines.append(f"| {no} | {title} | {'PASS' if ok else 'FAIL'} | {detail} |")
(ROOT / "docs" / "test_results.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
for no, title, ok, detail in results:
    print(f"{no:>2} {'PASS' if ok else 'FAIL'} {title} — {detail}")
sys.exit(0 if ok_all else 1)
