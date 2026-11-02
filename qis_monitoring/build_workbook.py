#!/usr/bin/env python3
"""QIS Fund Monitoring & Sales Update Workbook generator (Excel 365, no VBA).

    python3 build_workbook.py                 -> dist/QIS_Fund_Monitoring.xlsx (empty template)
    python3 build_workbook.py --fixtures      -> dist/QIS_Fund_Monitoring_TESTDATA.xlsx

The template contains NO market / fund / investor data. Input tables are empty and
must be filled from the official sources (Live Fund List, end_investor_holdings_view,
NAV files, market data, STEP Composition / Portfolio files).
"""
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula

from xlhelpers import (BORDER, F_BODY, F_BOLD, F_HEAD, F_HELP, F_NOTE, F_SECT, FILL, FMT,
                       Col, TableSpec, add_status_cf, expand, intro, legend, write_table, xlfn)

ROOT = Path(__file__).parent

# ------------------------------------------------------------------ capacities
CAP_FUND = 60        # 07/09/10/11 per-fund views
CAP_HOLD = 400       # aggregated investor x fund rows
CAP_POS = 1500       # position rows mirrored into 08/09
N_LEGS = 12          # legs shown on Dashboard
N_CHART = 60         # NAV observations on Dashboard chart

ALERTS = [  # code, name, category, data source, threshold text formula
    ("A01", "NAV_DATA", "Data", "04_NAV_Input",
     '"Stale M≥"&thr_NavStaleMed&"bd / H≥"&thr_NavStaleHigh&"bd, Missing=High"'),
    ("A02", "DAILY_MOVE", "Performance", "07_NAV_Calc",
     '"|1D| L≥"&TEXT(thr_D1_Low,"0.0%")&" M≥"&TEXT(thr_D1_Med,"0.0%")&" H≥"&TEXT(thr_D1_High,"0.0%")'),
    ("A03", "WEEKLY_MOVE", "Performance", "07_NAV_Calc",
     '"|1W| L≥"&TEXT(thr_W1_Low,"0.0%")&" M≥"&TEXT(thr_W1_Med,"0.0%")&" H≥"&TEXT(thr_W1_High,"0.0%")'),
    ("A04", "DRAWDOWN", "Performance", "07_NAV_Calc",
     '"DD L≥"&TEXT(thr_DD_Low,"0.0%")&" M≥"&TEXT(thr_DD_Med,"0.0%")&" H≥"&TEXT(thr_DD_High,"0.0%")'),
    ("A05", "AUM_CHANGE_EST", "Flow (estimated)", "04_NAV_Input (AUM)",
     '"|AUM 1W| M≥"&TEXT(thr_AUM_Med,"0%")&" H≥"&TEXT(thr_AUM_High,"0%")'),
    ("A06", "MAPPING", "Mapping", "03_Strategy_Map", '"Mapped=OK / Review=Medium / Required=High"'),
    ("A07", "POSITION_DATA", "Data", "06_Position_Raw",
     '"Stale M≥"&thr_PosStaleMed&"bd / H≥"&thr_PosStaleHigh&"bd, Missing=High"'),
    ("A08", "STEP_ITM", "Position", "08_Position_Calc", '"ITM legs > 0 = High"'),
    ("A09", "STEP_STRIKE_BUFFER", "Position", "08_Position_Calc",
     '"OTM buffer L≤"&TEXT(thr_Buf_Low,"0.0%")&" M≤"&TEXT(thr_Buf_Med,"0.0%")&" H≤"&TEXT(thr_Buf_High,"0.0%")'),
    ("A10", "POSITION_CALC_ISSUE", "Data", "08_Position_Calc", '"Spot/Strike/Parse/Underlying issue > 0 = High"'),
    ("A11", "MATURITY", "Position", "08_Position_Calc",
     '"Days to maturity L≤"&thr_Mat_Low&" M≤"&thr_Mat_Med'),
    ("A12", "POSITION_KEY_UNIQUE", "Data", "08_Position_Calc", '"Non-unique legs > 0 = Medium"'),
    ("A13", "BENCHMARK", "Market", "03_Strategy_Map / 05_Market_Input", '"Not set / unconfirmed / missing = Low"'),
    ("A14", "FUND_CODE_DUP", "Data", "01_Live_Fund_Master", '"Duplicate fund_code = High"'),
    ("A15", "HOLDINGS_DATA", "Data", "02_End_Investor_Holdings", '"Amount / master issue rows > 0 = Medium"'),
]
T = len(ALERTS)

PERIODS = [  # code, label, target-date expression on [@nav_date]
    ("1D", "1D", "[@nav_date]-1"),
    ("1W", "1W", "[@nav_date]-7"),
    ("MTD", "MTD", "EOMONTH([@nav_date],-1)"),
    ("3M", "3M", "EDATE([@nav_date],-3)"),
    ("YTD", "YTD", "DATE(YEAR([@nav_date])-1,12,31)"),
    ("1Y", "1Y", "EDATE([@nav_date],-12)"),
]

THRESHOLDS = [  # name, label, default, fmt
    ("thr_NavStaleMed", "NAV遅延 Medium（営業日）", 2, "int"),
    ("thr_NavStaleHigh", "NAV遅延 High（営業日）", 4, "int"),
    ("thr_D1_Low", "日次変動 Low（絶対値）", 0.01, "pct"),
    ("thr_D1_Med", "日次変動 Medium", 0.02, "pct"),
    ("thr_D1_High", "日次変動 High", 0.03, "pct"),
    ("thr_W1_Low", "週次変動 Low（絶対値）", 0.02, "pct"),
    ("thr_W1_Med", "週次変動 Medium", 0.04, "pct"),
    ("thr_W1_High", "週次変動 High", 0.06, "pct"),
    ("thr_DD_Low", "ドローダウン Low（下落率）", 0.03, "pct"),
    ("thr_DD_Med", "ドローダウン Medium", 0.05, "pct"),
    ("thr_DD_High", "ドローダウン High", 0.10, "pct"),
    ("thr_AUM_Med", "AUM週次変化 Medium（推定・絶対値）", 0.10, "pct"),
    ("thr_AUM_High", "AUM週次変化 High", 0.25, "pct"),
    ("thr_PosStaleMed", "Position遅延 Medium（営業日）", 2, "int"),
    ("thr_PosStaleHigh", "Position遅延 High（営業日）", 5, "int"),
    ("thr_Buf_Low", "STEP OTMバッファ Low（≤）", 0.10, "pct"),
    ("thr_Buf_Med", "STEP OTMバッファ Medium（≤）", 0.05, "pct"),
    ("thr_Buf_High", "STEP OTMバッファ High（≤）", 0.02, "pct"),
    ("thr_Mat_Low", "満期接近 Low（日以内）", 30, "int"),
    ("thr_Mat_Med", "満期接近 Medium（日以内）", 10, "int"),
    ("thr_SpotGap", "Spot日付とPosition日付の許容差（暦日）", 3, "int"),
]

DV = {
    "yn": "Yes,No",
    "src": "Composition,Portfolio",
    "navdef": "Cumulative TR index,Distribution-adjusted NAV,NAV + cumulative distributions (simple),Unknown",
    "greeks": "Per unit (weight-scalable),Position total,Unknown",
    "mode": "Individual,Combined",
    "diststat": "Not sent,Sent,On hold",
    "rev": "Unreviewed,Reviewed",
    "crev": "Draft,Reviewed,Approved,Rejected",
    "arev": "Unreviewed,Confirmed,False Positive,Escalated",
    "ares": "Open,Resolved,Accepted",
    "stype": "STEP,Target Return,Target Buying,CTA,Rates,Other",
}


def DK(x):
    """date -> text key (integer serial)."""
    return f'TEXT(INT({x}),"0")'


class Book:
    def __init__(self, fx=None):
        self.fx = fx or {}
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.ws = {}
        for name in ["00_Control", "01_Live_Fund_Master", "02_End_Investor_Holdings",
                     "03_Strategy_Map", "04_NAV_Input", "05_Market_Input", "06_Position_Raw",
                     "07_NAV_Calc", "08_Position_Calc", "09_Position_Change", "10_Alert",
                     "11_Update_Coverage", "12_Distribution_List", "13_Dashboard",
                     "14_Comment_Output", "15_QA_Output", "99_Check"]:
            self.ws[name] = self.wb.create_sheet(name)

    def name(self, nm, sheet, ref):
        self.wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{sheet}'!{ref}")

    def rows(self, key):
        return self.fx.get(key, [])

    def tbl(self, sheet, spec):
        ws = self.ws[sheet]
        res = write_table(ws, spec)
        ws.freeze_panes = ws.cell(spec.top + 1, 1)
        return res


# =============================================================== 00_Control
def sheet_control(b: Book):
    ws = b.ws["00_Control"]
    ctl = b.fx.get("control", {})
    intro(ws, "QIS Fund Monitoring & Sales Update Workbook — 00_Control",
          "評価日・対象投資家・対象ファンド・アラート閾値・休日を一元管理する。",
          "営業担当者の入力（黄色セル）。計算値（青）は他シートからの参照。",
          "評価日を入力 → 対象投資家・fund_codeをプルダウンで選択 → Dashboard/Comment/99_Checkが切替わる。",
          "閾値の初期値は暫定値。元仕様書の値と照合し C12 を Yes にするまで配信判定は READY になりません。")
    legend(ws, 6, 6)
    if ctl.get("banner"):
        ws["D1"] = ctl["banner"]
        ws["D1"].font = F_SECT
    ws["A7"] = "基本設定"
    ws["A7"].font = F_SECT
    items = [
        ("ctl_EvalDate", "評価日", ctl.get("eval_date"), "input", "date"),
        ("ctl_SelInvestor", "対象投資家（end_investor_bdr_id）", ctl.get("sel_investor"), "input", None),
        ("ctl_SelFund", "対象ファンド（fund_code）", ctl.get("sel_fund"), "input", None),
        ("ctl_Owner", "チェック担当者（Owner 既定値）", ctl.get("owner"), "input", None),
        ("ctl_ThrConfirmed", "閾値は元仕様書と照合済み（Yes/No）", ctl.get("thr_confirmed", "No"), "input", None),
        ("ctl_MaxBaseGap", "基準NAV日・指数日の許容ギャップ（暦日）", ctl.get("max_gap", 5), "input", "int"),
    ]
    r = 8
    for nm, lab, val, kind, fmt in items:
        ws.cell(r, 1, lab).font = F_BODY
        c = ws.cell(r, 3, val)
        c.fill, c.border, c.font = FILL[kind], BORDER, F_BOLD
        if fmt:
            c.number_format = FMT[fmt]
        ws.cell(r, 2, nm).font = F_HELP
        b.name(nm, "00_Control", f"$C${r}")
        r += 1
    dv = DataValidation(type="list", formula1="nl_Investors", allow_blank=True)
    dv.add("C9")
    ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1="nl_Funds", allow_blank=True)
    dv.add("C10")
    ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    dv.add("C12")
    ws.add_data_validation(dv)

    r = 15
    ws.cell(r, 1, "選択結果（計算）").font = F_SECT
    calc = [
        ("ctl_SelRow", "対象ファンドの行番号（07/09/10/11）",
         '=IF(ctl_SelFund="",0,IF(COUNTIF(tbl_NavCalc[fund_code],ctl_SelFund)=0,0,MATCH(ctl_SelFund,tbl_NavCalc[fund_code],0)))'),
        ("ctl_SelHoldKey", "holding_key", '=ctl_SelInvestor&"|"&ctl_SelFund'),
        ("ctl_SelHoldRow", "保有行（02集約 / 12明細）",
         '=IF(COUNTIF(tbl_HoldAgg[holding_key],ctl_SelHoldKey)=0,0,MATCH(ctl_SelHoldKey,tbl_HoldAgg[holding_key],0))'),
        ("ctl_SelDistRow", "配信先行（12_Distribution_List）",
         '=IF(ctl_SelInvestor="",0,IF(COUNTIF(tbl_Dist[end_investor_bdr_id],ctl_SelInvestor)=0,0,MATCH(ctl_SelInvestor,tbl_Dist[end_investor_bdr_id],0)))'),
        ("ctl_GlobalGate", "全体ゲート（99_Check Global）", "=chk_GlobalGate"),
        ("ctl_SelGate", "配信判定（選択中の投資家×ファンド）", "=chk_FinalGate"),
    ]
    for nm, lab, f in calc:
        r += 1
        ws.cell(r, 1, lab).font = F_BODY
        ws.cell(r, 2, nm).font = F_HELP
        c = ws.cell(r, 3, xlfn(f))
        c.fill, c.border, c.font = FILL["calc"], BORDER, F_BOLD
        b.name(nm, "00_Control", f"$C${r}")
    add_status_cf(ws, "C16:C21")

    r = 24
    ws.cell(r, 1, "アラート閾値（初期値＝暫定値。元仕様書の値で上書きしてください）").font = F_SECT
    over = ctl.get("thresholds", {})
    for nm, lab, val, fmt in THRESHOLDS:
        r += 1
        ws.cell(r, 1, lab).font = F_BODY
        ws.cell(r, 2, nm).font = F_HELP
        c = ws.cell(r, 3, over.get(nm, val))
        c.fill, c.border, c.number_format = FILL["input"], BORDER, FMT[fmt]
        b.name(nm, "00_Control", f"$C${r}")

    r += 2
    ws.cell(r, 1, "Position 最新版判定ルール").font = F_SECT
    for line in [
        "1) 同一 strategy_data_key × position_date × source_type 内で received_datetime が最大のファイルを最新版とする。",
        "2) received_datetime が同一の場合は file_seq（取込連番）が最大のものを最新版とする。",
        "3) 古い版は削除せず 06_Position_Raw に保持し、version_note=Superseded (kept) と表示する。",
        "4) 同一日に Composition がある場合は Composition を使用し、Portfolio は照合用（08 port_recon）。Composition が無い日のみ Portfolio で補完。",
    ]:
        r += 1
        ws.cell(r, 1, line).font = F_NOTE
    r += 2
    ws.cell(r, 1, "容量（行数）").font = F_SECT
    for lab, v in [("07/09/10/11 ファンド行", CAP_FUND), ("02 集約ビュー行", CAP_HOLD),
                   ("08/09 ポジション行", CAP_POS)]:
        r += 1
        ws.cell(r, 1, lab).font = F_BODY
        ws.cell(r, 3, v).font = F_BODY
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 22
    # Holidays
    b.tbl("00_Control", TableSpec("tbl_Holiday", [
        Col("holiday_date", "input", fmt="date", w=14), Col("holiday_name", "input", w=18),
        Col("market", "input", w=10)], top=7, left=10, rows=b.rows("holidays")))
    ws.cell(6, 10, "休日（NETWORKDAYS.INTL で使用）").font = F_SECT


# =============================================================== 01 Fund master
def sheet_fund(b: Book):
    ws = b.ws["01_Live_Fund_Master"]
    intro(ws, "01_Live_Fund_Master — Live Fund List（fund_code 単位）",
          "Live Fund List を fund_code 単位で保持し、重複・保有状況・同一Index参照ファンド数を確認する。",
          "Live Fund List（グレー列に値貼り付け）。",
          "グレー列に貼り付け（行はテーブル末尾に追加）。青列は自動拡張。重複行は削除せず 99_Check で確認。",
          "列名が元ファイルと異なる場合は、元ファイル側の列をこの順に並べてから貼り付けること（列の意味を推測で変えない）。")
    G = '=IF([@fund_code]="","",{})'
    cols = [
        Col("fund_code", w=12), Col("fund_name", w=28), Col("fund_status", w=10),
        Col("strategy_type", w=13, dv=DV["stype"]), Col("index_name", w=22), Col("index_ticker", w=14),
        Col("currency", w=8), Col("inception_date", fmt="date", w=12), Col("manager", w=14),
        Col("distribution_policy", w=14), Col("source_file", w=20), Col("load_date", fmt="date", w=12),
        Col("dup_count", "calc", G.format("COUNTIFS(§[fund_code],[@fund_code])"), "int", 9),
        Col("dup_flag", "calc", G.format('IF([@dup_count]>1,"DUPLICATE","OK")'), w=11),
        Col("same_index_fund_count", "calc",
            G.format('IF([@index_name]="","",COUNTIFS(§[index_name],[@index_name]))'), "int", 11),
        Col("investor_count", "calc", G.format("COUNTIFS(tbl_HoldAgg[fund_code],[@fund_code])"), "int", 10),
        Col("total_amount", "calc",
            G.format("SUMIFS(tbl_HoldAgg[amount_num],tbl_HoldAgg[fund_code],[@fund_code])"), "amt", 16),
        Col("fund_map_rows", "calc", G.format("COUNTIF(tbl_FundMap[fund_code],[@fund_code])"), "int", 10),
    ]
    b.tbl("01_Live_Fund_Master", TableSpec("tbl_Fund", cols, rows=b.rows("funds")))
    b.wb.defined_names["nl_Funds"] = DefinedName("nl_Funds", attr_text="tbl_Fund[fund_code]")


# =============================================================== 02 holdings
def sheet_holdings(b: Book):
    ws = b.ws["02_End_Investor_Holdings"]
    intro(ws, "02_End_Investor_Holdings — 投資家別保有（明細 tbl_HoldRaw ／ 集約 tbl_HoldAgg）",
          "end_investor_holdings_view の全行を保持（左）し、end_investor_bdr_id × fund_code で Amount を集約（右）。",
          "end_investor_holdings_view.xlsx（グレー列に値貼り付け）。",
          "左表グレー列に全行を貼付け。右の集約ビューは数式で自動生成（容量 %d 行。超過時は 99_Check C06 が DATA REQUIRED）。" % CAP_HOLD,
          "Amount に数値以外がある holding_key は集約額を空欄にし Check Required とする（0扱いしない）。")
    G = '=IF([@holding_key]="","",{})'
    raw = [
        Col("as_of_date", fmt="date", w=11), Col("end_investor_bdr_id", w=14), Col("end_investor_name", w=22),
        Col("fund_code", w=11), Col("fund_name_raw", w=22), Col("amount", fmt="amt", w=15),
        Col("currency", w=8), Col("sales_rep", w=12), Col("source_file", w=18), Col("source_row", fmt="int", w=8),
        Col("seq", "calc", "=ROW()-ROW(§[#Headers])", "int", 6),
        Col("holding_key", "calc",
            '=IF(OR([@end_investor_bdr_id]="",[@fund_code]=""),"",[@end_investor_bdr_id]&"|"&[@fund_code])', w=20),
        Col("amount_check", "calc",
            G.format('IF([@amount]="","Blank",IF(ISNUMBER([@amount]),"OK","Not numeric"))'), w=11),
        Col("key_first", "calc", G.format("IF(MATCH([@holding_key],§[holding_key],0)=[@seq],1,0)"), "int", 8),
        Col("key_rank", "calc", '=IF([@key_first]=1,COUNTIFS(§[key_first],1,§[seq],"<="&[@seq]),"")', "int", 8),
        Col("inv_first", "calc",
            '=IF([@end_investor_bdr_id]="","",IF(MATCH([@end_investor_bdr_id],§[end_investor_bdr_id],0)=[@seq],1,0))',
            "int", 8),
        Col("fund_in_master", "calc",
            G.format('IF(COUNTIF(tbl_Fund[fund_code],[@fund_code])>0,"Yes","No")'), w=9),
        Col("in_dist_list", "calc",
            '=IF([@end_investor_bdr_id]="","",IF(COUNTIF(tbl_Dist[end_investor_bdr_id],[@end_investor_bdr_id])>0,"Yes","No"))',
            w=9),
    ]
    b.tbl("02_End_Investor_Holdings", TableSpec("tbl_HoldRaw", raw, rows=b.rows("holdings")))
    left = len(raw) + 2
    R = lambda x: f"INDEX(tbl_HoldRaw[{x}],MATCH([@holding_key],tbl_HoldRaw[holding_key],0))"
    K = '=IF([@holding_key]="","",{})'
    agg = [
        Col("n", "calc", "=ROW()-ROW(§[#Headers])", "int", 5),
        Col("holding_key", "calc",
            '=IF([@n]>MAX(tbl_HoldRaw[key_rank]),"",INDEX(tbl_HoldRaw[holding_key],MATCH([@n],tbl_HoldRaw[key_rank],0)))',
            w=20),
        Col("end_investor_bdr_id", "calc", K.format(R("end_investor_bdr_id")), w=14),
        Col("end_investor_name", "calc", K.format(R("end_investor_name") + '&""'), w=22),
        Col("fund_code", "calc", K.format(R("fund_code")), w=11),
        Col("fund_name", "calc",
            K.format('IF(COUNTIF(tbl_Fund[fund_code],[@fund_code])=0,"(Not in Live Fund List)",'
                     'INDEX(tbl_Fund[fund_name],MATCH([@fund_code],tbl_Fund[fund_code],0))&"")'), w=24),
        Col("source_rows", "calc", K.format("COUNTIFS(tbl_HoldRaw[holding_key],[@holding_key])"), "int", 8),
        Col("invalid_amount_rows", "calc",
            K.format('COUNTIFS(tbl_HoldRaw[holding_key],[@holding_key],tbl_HoldRaw[amount_check],"<>OK")'), "int", 9),
        Col("amount_total", "calc",
            K.format('IF([@invalid_amount_rows]>0,"",SUMIFS(tbl_HoldRaw[amount],tbl_HoldRaw[holding_key],[@holding_key]))'),
            "amt", 16),
        Col("amount_num", "calc", '=IF(ISNUMBER([@amount_total]),[@amount_total],0)', "amt", 4),
        Col("currency", "calc", K.format(R("currency") + '&""'), w=8),
        Col("as_of_date", "calc",
            K.format("MAXIFS(tbl_HoldRaw[as_of_date],tbl_HoldRaw[holding_key],[@holding_key])"), "date", 11),
        Col("sales_rep", "calc", K.format(R("sales_rep") + '&""'), w=12),
        Col("fund_in_master", "calc", K.format(R("fund_in_master")), w=9),
        Col("agg_status", "calc",
            K.format('IF([@invalid_amount_rows]>0,"Check Required: Amount",'
                     'IF([@fund_in_master]="No","Check Required: Fund not in Live list","OK"))'), w=26),
    ]
    b.tbl("02_End_Investor_Holdings",
          TableSpec("tbl_HoldAgg", agg, left=left, nrows=CAP_HOLD))
    ws.cell(6, left, "集約ビュー（1投資家×1ファンド＝1行。数式で自動生成・手入力禁止）").font = F_SECT
    ws.cell(6, 1, "明細（元データ全行。上書き・削除しない）").font = F_SECT
    b.wb.defined_names["nl_Investors"] = DefinedName("nl_Investors", attr_text="tbl_HoldAgg[end_investor_bdr_id]")


# =============================================================== 03 strategy map
def sheet_map(b: Book):
    ws = b.ws["03_Strategy_Map"]
    intro(ws, "03_Strategy_Map — Fund → Strategy Data Key",
          "fund_code ごとに Strategy Data Key を決定する（①fund_code明示 → ②Index承認済マッピング → ③手動設定）。",
          "営業/商品担当の入力（黄色）。Index 承認済マッピングは右表 tbl_IndexMap。",
          "Live Fund List の全 fund_code を左表に1行ずつ登録。曖昧な場合は自動選択せず Review/Mapping Required を表示。",
          "fund_exposure_multiplier は根拠と確認(Yes)が揃った場合のみ multiplier_used に反映。未確認は空欄（1.0とみなさない）。")
    G = '=IF([@fund_code]="","",{})'
    MS = '[@fund_code]'
    cols = [
        Col("fund_code", "input", w=11), Col("explicit_strategy_data_key", "input", w=18),
        Col("manual_strategy_data_key", "input", w=18), Col("mapping_approved_by", "input", w=12),
        Col("mapping_approved_date", "input", fmt="date", w=12),
        Col("fund_exposure_multiplier", "input", fmt="num4", w=11), Col("multiplier_basis", "input", w=18),
        Col("multiplier_confirmed", "input", w=10, dv=DV["yn"]),
        Col("benchmark_series_id", "input", w=14), Col("benchmark_basis_confirmed", "input", w=11, dv=DV["yn"]),
        Col("nav_dist_definition", "input", w=22, dv=DV["navdef"]),
        Col("nav_dist_def_confirmed", "input", w=10, dv=DV["yn"]), Col("note", "input", w=18),
        Col("in_master", "calc", G.format(f'IF(COUNTIF(tbl_Fund[fund_code],{MS})>0,"Yes","No")'), w=8),
        Col("dup_rows", "calc", G.format(f"COUNTIF(§[fund_code],{MS})"), "int", 7),
        Col("index_name", "calc",
            G.format(f'IF([@in_master]="No","",INDEX(tbl_Fund[index_name],MATCH({MS},tbl_Fund[fund_code],0))&"")'), w=20),
        Col("strategy_type", "calc",
            G.format(f'IF([@in_master]="No","",INDEX(tbl_Fund[strategy_type],MATCH({MS},tbl_Fund[fund_code],0))&"")'), w=12),
        Col("idx_approved_count", "calc",
            G.format('IF([@index_name]="",0,COUNTIFS(tbl_IndexMap[index_name],[@index_name],tbl_IndexMap[approved],"Yes"))'),
            "int", 9),
        Col("idx_candidate_key", "calc",
            G.format('IF([@idx_approved_count]<>1,"",INDEX(tbl_IndexMap[strategy_data_key],'
                     'MATCH([@index_name]&"|Yes",tbl_IndexMap[idx_key],0))&"")'), w=18),
        Col("resolved_strategy_data_key", "calc",
            G.format('IF([@explicit_strategy_data_key]<>"",[@explicit_strategy_data_key],'
                     'IF([@idx_approved_count]=1,[@idx_candidate_key],[@manual_strategy_data_key]&""))'), w=18),
        Col("mapping_method", "calc",
            G.format('IF([@explicit_strategy_data_key]<>"","1: fund_code explicit",IF([@idx_approved_count]=1,'
                     '"2: Index approved",IF([@manual_strategy_data_key]<>"","3: Manual","-")))'), w=16),
        Col("mapping_status", "calc",
            G.format('IF([@in_master]="No","Review Required: not in Live list",'
                     'IF([@dup_rows]>1,"Review Required: duplicate map rows",'
                     'IF([@explicit_strategy_data_key]<>"",IF(AND([@idx_approved_count]=1,[@idx_candidate_key]<>[@explicit_strategy_data_key]),'
                     '"Review Required: explicit key differs from Index map","Mapped"),'
                     'IF([@idx_approved_count]=1,"Mapped",'
                     'IF([@manual_strategy_data_key]<>"",IF([@mapping_approved_by]="","Review Required: manual key not approved","Mapped"),'
                     'IF([@idx_approved_count]>1,"Review Required: multiple keys for Index","Mapping Required"))))))'),
            w=30),
        Col("multiplier_used", "calc",
            G.format('IF(AND(ISNUMBER([@fund_exposure_multiplier]),[@multiplier_basis]<>"",'
                     '[@multiplier_confirmed]="Yes"),[@fund_exposure_multiplier],"")'), "num4", 10),
        Col("key_shared_fund_count", "calc",
            G.format('IF([@resolved_strategy_data_key]="","",COUNTIF(§[resolved_strategy_data_key],[@resolved_strategy_data_key]))'),
            "int", 9),
        Col("pos_rows_for_key", "calc",
            G.format('IF([@resolved_strategy_data_key]="","",COUNTIF(tbl_PosRaw[strategy_data_key],[@resolved_strategy_data_key]))'),
            "int", 9),
    ]
    b.tbl("03_Strategy_Map", TableSpec("tbl_FundMap", cols, rows=b.rows("fundmap")))
    left = len(cols) + 2
    idx = [
        Col("index_name", "input", w=20), Col("strategy_data_key", "input", w=18),
        Col("approved", "input", w=9, dv=DV["yn"]), Col("approved_by", "input", w=12),
        Col("approved_date", "input", fmt="date", w=12), Col("note", "input", w=18),
        Col("idx_key", "calc", '=IF([@index_name]="","",[@index_name]&"|"&[@approved])', w=22),
        Col("dup_count", "calc",
            '=IF([@index_name]="","",COUNTIFS(§[index_name],[@index_name],§[strategy_data_key],[@strategy_data_key]))', "int", 8),
    ]
    b.tbl("03_Strategy_Map", TableSpec("tbl_IndexMap", idx, left=left, rows=b.rows("indexmap")))
    ws.cell(6, left, "Index → Strategy Data Key 承認済マッピング（同一Indexに複数Keyがある場合は自動選択しない）").font = F_SECT
    ws.cell(6, 1, "Fund 別マッピング（黄色＝入力 / 青＝判定）").font = F_SECT


# =============================================================== 04 NAV
def sheet_nav(b: Book):
    ws = b.ws["04_NAV_Input"]
    intro(ws, "04_NAV_Input — ファンド別 NAV 時系列（元データ）",
          "fund_code × nav_date の NAV・分配金・AUM を保持。リターンは 07_NAV_Calc でファンド固有の時系列から計算。",
          "管理会社/運用会社の NAV ファイル（グレー列に値貼り付け）。",
          "新しい日付の行をテーブル末尾に追加（古い行は削除しない）。青列は自動拡張。",
          "同一 fund_code × nav_date が複数あると両方 valid=0（推測で片方を採用しない）→ 99_Check C11。AUMは公式値、フローは推定。")
    G = '=IF([@nav_key]="","",{})'
    cols = [
        Col("fund_code", w=11), Col("nav_date", fmt="date", w=11), Col("nav", fmt="nav", w=12),
        Col("nav_with_distribution", fmt="nav", w=14), Col("distribution_amount", fmt="nav", w=12),
        Col("aum", fmt="amt", w=16), Col("aum_currency", w=8), Col("source", w=12),
        Col("source_file", w=18), Col("received_date", fmt="date", w=11),
        Col("nav_key", "calc",
            f'=IF(OR([@fund_code]="",NOT(ISNUMBER([@nav_date]))),"",[@fund_code]&"|"&{DK("[@nav_date]")})', w=16),
        Col("dup_count", "calc", G.format("COUNTIFS(§[nav_key],[@nav_key])"), "int", 7),
        Col("valid", "calc", G.format("IF(AND(ISNUMBER([@nav]),[@dup_count]=1),IF([@nav]>0,1,0),0)"), "int", 6),
        Col("row_status", "calc",
            G.format('IF([@valid]=1,"OK",IF([@dup_count]>1,"Duplicate date",'
                     'IF(NOT(ISNUMBER([@nav])),"NAV missing / non-numeric","NAV<=0")))'), w=18),
        Col("running_peak", "calc",
            '=IF([@valid]=1,MAXIFS(§[nav],§[fund_code],[@fund_code],§[valid],1,§[nav_date],"<="&[@nav_date]),"")',
            "nav", 12),
        Col("dd_from_peak", "calc", '=IF([@valid]=1,[@nav]/[@running_peak]-1,"")', "pct", 10),
    ]
    b.tbl("04_NAV_Input", TableSpec("tbl_NAV", cols, rows=b.rows("nav")))


# =============================================================== 05 Market
def sheet_market(b: Book):
    ws = b.ws["05_Market_Input"]
    intro(ws, "05_Market_Input — 市場データ（指数・Spot）",
          "参照指数・STEP 原資産 Spot の時系列を保持する。",
          "Bloomberg 等からの値貼り付け（グレー列）。右表 tbl_UndlMap はティッカー原資産 → Spot 系列の確認済対応表（黄色）。",
          "行を末尾に追加。series_id は Bloomberg 形式（例 'NKY Index'）で統一。",
          "通貨・ヘッジ条件が異なる指数はファンドのベンチマークとして扱わない（03 benchmark_basis_confirmed）。")
    G = '=IF([@mkt_key]="","",{})'
    cols = [
        Col("series_id", w=14), Col("series_name", w=22), Col("data_date", fmt="date", w=11),
        Col("value", fmt="num", w=12), Col("currency", w=8), Col("hedge_type", w=10),
        Col("series_type", w=10), Col("source", w=12), Col("source_file", w=18),
        Col("received_date", fmt="date", w=11),
        Col("mkt_key", "calc",
            f'=IF(OR([@series_id]="",NOT(ISNUMBER([@data_date]))),"",[@series_id]&"|"&{DK("[@data_date]")})', w=18),
        Col("dup_count", "calc", G.format("COUNTIFS(§[mkt_key],[@mkt_key])"), "int", 7),
        Col("valid", "calc", G.format("IF(AND(ISNUMBER([@value]),[@dup_count]=1),1,0)"), "int", 6),
    ]
    b.tbl("05_Market_Input", TableSpec("tbl_Mkt", cols, rows=b.rows("market")))
    left = len(cols) + 2
    um = [
        Col("ticker_underlying", "input", w=16), Col("spot_series_id", "input", w=14),
        Col("confirmed", "input", w=9, dv=DV["yn"]), Col("confirmed_by", "input", w=12), Col("note", "input", w=18),
        Col("map_key", "calc", '=IF([@ticker_underlying]="","",[@ticker_underlying]&"|"&[@confirmed])', w=20),
    ]
    b.tbl("05_Market_Input", TableSpec("tbl_UndlMap", um, left=left, rows=b.rows("undlmap")))
    ws.cell(6, left, "原資産マッピング（ティッカー解析の原資産 → Spot 系列。確認済(Yes)のみ使用。未登録は同名系列を使用）").font = F_SECT


# =============================================================== 06 Position raw
def sheet_posraw(b: Book):
    ws = b.ws["06_Position_Raw"]
    intro(ws, "06_Position_Raw — STEP 等ポジション（Composition / Portfolio 原本）",
          "受領ファイルの原本値を改変せず保持し、最新版・使用行を判定する。",
          "Composition / Portfolio ファイル（グレー列に値貼り付け。source_type・file_name・received_datetime・file_seq を必ず記入）。",
          "新ファイルは末尾に追加（古い版は削除しない）。最新版判定ルールは 00_Control 参照。",
          "Ticker 解析値（Strike / Put-Call / Underlying）は 08_Position_Calc の別列に保持し、本シートの原本は上書きしない。")
    G = '=IF([@strategy_data_key]="","",{})'
    KD = "§[strategy_data_key],[@strategy_data_key],§[position_date],[@position_date]"
    cols = [
        Col("source_type", w=11, dv=DV["src"]), Col("strategy_data_key", w=16), Col("position_date", fmt="date", w=11),
        Col("received_datetime", fmt="dt", w=15), Col("file_name", w=22), Col("file_seq", fmt="int", w=7),
        Col("leg_id", w=8), Col("BloombergTicker", w=26), Col("underlying_raw", w=12), Col("put_call_raw", w=8),
        Col("strike_raw", fmt="num", w=10), Col("maturity_raw", fmt="date", w=11), Col("date_of_entry", fmt="date", w=11),
        Col("weight", fmt="num4", w=9), Col("notional", fmt="amt", w=13), Col("quantity", fmt="num", w=10),
        Col("delta_reported", fmt="num4", w=9), Col("gamma_reported", fmt="num4", w=9),
        Col("vega_reported", fmt="num4", w=9), Col("theta_reported", fmt="num4", w=9),
        Col("greeks_basis", w=14, dv=DV["greeks"]), Col("premium_received", fmt="num", w=11),
        Col("premium_paid", fmt="num", w=11), Col("mtm_value", fmt="num", w=11),
        Col("cumulative_premium", fmt="num", w=11), Col("spot_raw", fmt="num", w=10),
        Col("spot_underlying_raw", w=12), Col("data_quality_raw", w=12),
        Col("latest_received", "calc",
            G.format(f"MAXIFS(§[received_datetime],{KD},§[source_type],[@source_type])"), "dt", 15),
        Col("latest_seq", "calc",
            G.format(f"MAXIFS(§[file_seq],{KD},§[source_type],[@source_type],§[received_datetime],[@latest_received])"),
            "int", 7),
        Col("is_latest_file", "calc",
            G.format("IF(AND(N([@received_datetime])=[@latest_received],N([@file_seq])=[@latest_seq]),1,0)"), "int", 7),
        Col("comp_exists", "calc", G.format(f'COUNTIFS({KD},§[source_type],"Composition")'), "int", 7),
        Col("use_row", "calc",
            G.format('IF(AND([@is_latest_file]=1,OR([@source_type]="Composition",[@comp_exists]=0)),1,0)'), "int", 6),
        Col("version_note", "calc", G.format('IF([@is_latest_file]=1,"Latest","Superseded (kept)")'), w=14),
    ]
    b.tbl("06_Position_Raw", TableSpec("tbl_PosRaw", cols, rows=b.rows("positions")))


# =============================================================== 07 NAV calc
def sheet_navcalc(b: Book):
    ws = b.ws["07_NAV_Calc"]
    intro(ws, "07_NAV_Calc — ファンド別 NAV 分析",
          "ファンド固有の NAV 時系列からリターン・ドローダウン・AUM 変化（推定フロー）・参照指数リターンを算出。",
          "04_NAV_Input / 05_Market_Input / 03_Strategy_Map（数式のみ。入力不要）。",
          "自動計算。各期間は『基準日以前で最も近い有効NAV日』を使用（Weekly は7暦日前）。",
          "基準NAV日が許容ギャップ（Control）を超える・履歴不足の期間は空欄。Total Return は nav_with_distribution の定義が確認済の場合のみ。")
    G = '=IF([@fund_code]="","",{})'
    F = "tbl_NAV[fund_code],[@fund_code],tbl_NAV[valid],1"

    def navat(col, d):
        return f'INDEX(tbl_NAV[{col}],MATCH([@fund_code]&"|"&{DK(d)},tbl_NAV[nav_key],0))'

    def nearest(d):
        return (f'IF(COUNTIFS({F},tbl_NAV[nav_date],"<="&{d})=0,"",'
                f'MAXIFS(tbl_NAV[nav_date],{F},tbl_NAV[nav_date],"<="&{d}))')

    FN = lambda x: f'IF(INDEX(tbl_Fund[{x}],[@n])="","",INDEX(tbl_Fund[{x}],[@n]))'
    M = lambda x: f'IF([@map_row]=0,"",INDEX(tbl_FundMap[{x}],[@map_row])&"")'
    cols = [
        Col("n", "calc", "=ROW()-ROW(§[#Headers])", "int", 5),
        Col("fund_code", "calc",
            '=IF([@n]>ROWS(tbl_Fund[fund_code]),"",IF(INDEX(tbl_Fund[fund_code],[@n])="","",INDEX(tbl_Fund[fund_code],[@n])))',
            w=11),
        Col("fund_name", "calc", G.format(FN("fund_name")), w=24),
        Col("strategy_type", "calc", G.format(FN("strategy_type")), w=11),
        Col("index_name", "calc", G.format(FN("index_name")), w=18),
        Col("currency", "calc", G.format(FN("currency")), w=7),
        Col("inception_date", "calc", G.format(FN("inception_date")), "date", 11),
        Col("map_row", "calc",
            G.format("IF(COUNTIF(tbl_FundMap[fund_code],[@fund_code])=0,0,MATCH([@fund_code],tbl_FundMap[fund_code],0))"),
            "int", 6),
        Col("eval_date", "calc", G.format('IF(ISNUMBER(ctl_EvalDate),ctl_EvalDate,"")'), "date", 11),
        Col("nav_valid_count", "calc", G.format(f"COUNTIFS({F})"), "int", 7),
        Col("nav_dup_rows", "calc",
            G.format('COUNTIFS(tbl_NAV[fund_code],[@fund_code],tbl_NAV[dup_count],">1")'), "int", 7),
        Col("nav_date", "calc", G.format('IF([@eval_date]="","",' + nearest("[@eval_date]") + ")"), "date", 11),
        Col("nav", "calc", G.format('IF([@nav_date]="","",' + navat("nav", "[@nav_date]") + ")"), "nav", 11),
        Col("first_nav_date", "calc",
            G.format(f'IF([@nav_valid_count]=0,"",MINIFS(tbl_NAV[nav_date],{F}))'), "date", 11),
        Col("nav_stale_bd", "calc",
            '=IF(OR([@nav_date]="",[@eval_date]=""),"",NETWORKDAYS.INTL([@nav_date],[@eval_date],1,tbl_Holiday[holiday_date])-1)',
            "int", 7),
        Col("nav_status", "calc",
            G.format('IF([@eval_date]="","Eval date missing",IF([@nav_date]="","Missing",'
                     'IF([@nav_stale_bd]>=thr_NavStaleHigh,"Stale (High)",IF([@nav_stale_bd]>=thr_NavStaleMed,"Stale (Medium)","OK"))))'),
            w=14),
    ]
    for code, _, tgt in PERIODS:
        cols += [
            Col(f"tgt_{code}", "calc", f'=IF([@nav_date]="","",{tgt})', "date", 11),
            Col(f"base_{code}", "calc", f'=IF([@tgt_{code}]="","",{nearest(f"[@tgt_{code}]")})', "date", 11),
            Col(f"ret_{code}", "calc",
                f'=IF([@base_{code}]="","",IF([@tgt_{code}]-[@base_{code}]>ctl_MaxBaseGap,"",'
                f'[@nav]/{navat("nav", f"[@base_{code}]")}-1))', "pct", 9),
        ]
    cols += [
        Col("ret_SI", "calc",
            '=IF(OR([@nav_date]="",[@first_nav_date]=""),"",IF([@nav_date]<=[@first_nav_date],"",'
            f'[@nav]/{navat("nav", "[@first_nav_date]")}-1))', "pct", 9),
        Col("si_note", "calc",
            G.format('IF(AND(ISNUMBER([@inception_date]),ISNUMBER([@first_nav_date])),'
                     'IF([@first_nav_date]<>[@inception_date],"初回NAV日≠設定日（SIは初回NAV日起点）",""),"")'), w=20),
        Col("unavailable_periods", "calc",
            G.format('TEXTJOIN(", ",TRUE,' + ",".join(f'IF([@ret_{c}]="","{c}","")' for c, _, _ in PERIODS)
                     + ',IF([@ret_SI]="","SI",""))'), w=16),
        Col("nav_dist_definition", "calc", G.format(M("nav_dist_definition")), w=18),
        Col("tr_ok", "calc",
            G.format('AND(OR([@nav_dist_definition]="Cumulative TR index",[@nav_dist_definition]="Distribution-adjusted NAV"),'
                     + M("nav_dist_def_confirmed") + '="Yes")'), w=7),
    ]
    for code, base in [("YTD", "[@base_YTD]"), ("1Y", "[@base_1Y]"), ("SI", "[@first_nav_date]")]:
        e, s = navat("nav_with_distribution", "[@nav_date]"), navat("nav_with_distribution", base)
        cols.append(Col(f"tr_{code}", "calc",
                        f'=IF(OR([@tr_ok]<>TRUE,[@ret_{code}]=""),"",IF(AND(ISNUMBER({e}),ISNUMBER({s})),{e}/{s}-1,""))',
                        "pct", 9))
    cols += [
        Col("tr_status", "calc",
            G.format('IF([@tr_ok]=TRUE,"TR computed (definition confirmed)","TR not computed: nav_with_distribution definition '
                     'unconfirmed ("&IF([@nav_dist_definition]="","blank",[@nav_dist_definition])&")")'), w=30),
        Col("dist_ytd", "calc",
            '=IF([@nav_date]="","",SUMIFS(tbl_NAV[distribution_amount],tbl_NAV[fund_code],[@fund_code],'
            'tbl_NAV[nav_date],">"&DATE(YEAR([@nav_date])-1,12,31),tbl_NAV[nav_date],"<="&[@nav_date]))', "nav", 9),
        Col("peak_nav", "calc",
            f'=IF([@nav_date]="","",MAXIFS(tbl_NAV[nav],{F},tbl_NAV[nav_date],"<="&[@nav_date]))', "nav", 11),
        Col("cur_dd", "calc", '=IF([@nav_date]="","",[@nav]/[@peak_nav]-1)', "pct", 9),
        Col("max_dd", "calc",
            f'=IF([@nav_date]="","",MINIFS(tbl_NAV[dd_from_peak],{F},tbl_NAV[nav_date],"<="&[@nav_date]))', "pct", 9),
        Col("aum", "calc",
            f'=IF([@nav_date]="","",IF(ISNUMBER({navat("aum", "[@nav_date]")}),{navat("aum", "[@nav_date]")},""))', "amt", 15),
        Col("aum_1w_base", "calc",
            f'=IF([@base_1W]="","",IF(ISNUMBER({navat("aum", "[@base_1W]")}),{navat("aum", "[@base_1W]")},""))', "amt", 15),
        Col("aum_chg_1w", "calc",
            '=IF(AND(ISNUMBER([@aum]),ISNUMBER([@aum_1w_base])),IF([@aum_1w_base]=0,"",[@aum]/[@aum_1w_base]-1),"")', "pct", 9),
        Col("est_flow_1w", "calc",
            '=IF(AND(ISNUMBER([@aum]),ISNUMBER([@aum_1w_base]),ISNUMBER([@ret_1W])),[@aum]-[@aum_1w_base]*(1+[@ret_1W]),"")',
            "amt", 15),
        Col("bm_series", "calc", G.format(M("benchmark_series_id")), w=13),
        Col("bm_label", "calc",
            G.format('IF([@bm_series]="","Not set",IF(' + M("benchmark_basis_confirmed")
                     + '="Yes","Benchmark (confirmed)","Reference only (ccy/hedge unconfirmed)"))'), w=24),
    ]
    BM = "tbl_Mkt[series_id],[@bm_series],tbl_Mkt[valid],1"

    def bm_near(d):
        return (f'IF(COUNTIFS({BM},tbl_Mkt[data_date],"<="&{d})=0,"",'
                f'MAXIFS(tbl_Mkt[data_date],{BM},tbl_Mkt[data_date],"<="&{d}))')

    def bmval(d):
        return f'INDEX(tbl_Mkt[value],MATCH([@bm_series]&"|"&{DK(d)},tbl_Mkt[mkt_key],0))'

    cols.append(Col("bm_end_date", "calc",
                    f'=IF(OR([@bm_series]="",[@nav_date]=""),"",{bm_near("[@nav_date]")})', "date", 11))
    for code, _, _ in PERIODS:
        cols += [
            Col(f"bm_base_{code}", "calc",
                f'=IF(OR([@bm_end_date]="",[@base_{code}]=""),"",{bm_near(f"[@base_{code}]")})', "date", 11),
            Col(f"bm_ret_{code}", "calc",
                f'=IF(OR([@bm_base_{code}]="",[@ret_{code}]=""),"",IF(OR([@base_{code}]-[@bm_base_{code}]>ctl_MaxBaseGap,'
                f'[@nav_date]-[@bm_end_date]>ctl_MaxBaseGap),"",{bmval("[@bm_end_date]")}/{bmval(f"[@bm_base_{code}]")}-1))',
                "pct", 9),
        ]
    b.tbl("07_NAV_Calc", TableSpec("tbl_NavCalc", cols, nrows=CAP_FUND))


# =============================================================== 08 Position calc
def sheet_poscalc(b: Book):
    ws = b.ws["08_Position_Calc"]
    intro(ws, "08_Position_Calc — ポジション計算（06 の各行を同じ順序でミラー）",
          "Ticker 解析、Spot 取得、Strike Distance、ITM/OTM、満期、Greeks、Premium、スナップショット判定、Portfolio 照合。",
          "06_Position_Raw / 05_Market_Input / 03_Strategy_Map（数式のみ）。",
          "自動計算（容量 %d 行。06 の行数が超えると 99_Check C16 が DATA REQUIRED）。" % CAP_POS,
          "Putの符号付き Strike Distance＝Strike/Spot−1。Spot欠損・原資産不一致時は計算しない。Greeks/Premium は定義確認なしに合算しない。")
    G = '=IF([@src_row]="","",{})'
    R = lambda x: f"INDEX(tbl_PosRaw[{x}],[@src_row])"
    CP = lambda x: G.format(f'IF({R(x)}="","",{R(x)})')
    S = "[@slash]"
    cols = [
        Col("n", "calc", "=ROW()-ROW(§[#Headers])", "int", 5),
        Col("src_row", "calc",
            '=IF([@n]>ROWS(tbl_PosRaw[strategy_data_key]),"",IF(INDEX(tbl_PosRaw[strategy_data_key],[@n])="","",[@n]))',
            "int", 6),
        Col("strategy_data_key", "calc", CP("strategy_data_key"), w=15),
        Col("source_type", "calc", CP("source_type"), w=11),
        Col("position_date", "calc", CP("position_date"), "date", 11),
        Col("file_name", "calc", CP("file_name"), w=18),
        Col("is_latest_file", "calc", CP("is_latest_file"), "int", 6),
        Col("use_row", "calc", CP("use_row"), "int", 6),
        Col("ticker", "calc", CP("BloombergTicker"), w=26),
        Col("date_of_entry", "calc", CP("date_of_entry"), "date", 11),
        Col("leg_id", "calc", CP("leg_id"), w=7),
        Col("tk", "calc", G.format("TRIM([@ticker])"), w=8),
        Col("slash", "calc", '=IF(OR([@src_row]="",[@tk]=""),"",IFERROR(FIND("/",[@tk]),0))', "int", 5),
        Col("parse_ok", "calc",
            '=IF([@slash]="","",IFERROR(AND(' + S + '>3,MID([@tk],' + S + '+3,1)="/",MID([@tk],' + S + '-3,1)=" ",'
            'VALUE(MID([@tk],' + S + '-2,2))>=1,VALUE(MID([@tk],' + S + '-2,2))<=12,'
            'VALUE(MID([@tk],' + S + '+1,2))>=1,VALUE(MID([@tk],' + S + '+1,2))<=31,'
            'OR(MID([@tk],' + S + '+7,1)="P",MID([@tk],' + S + '+7,1)="C"),'
            'ISNUMBER(VALUE(MID([@tk],' + S + '+8,FIND(" ",[@tk]&" ",' + S + '+8)-' + S + '-8)))),FALSE))', w=7),
        Col("pc_parsed", "calc", '=IF([@parse_ok]=TRUE,MID([@tk],' + S + '+7,1),"")', w=6),
        Col("strike_parsed", "calc",
            '=IF([@parse_ok]=TRUE,VALUE(MID([@tk],' + S + '+8,FIND(" ",[@tk]&" ",' + S + '+8)-' + S + '-8)),"")', "num", 10),
        Col("expiry_parsed", "calc",
            '=IF([@parse_ok]=TRUE,DATE(2000+VALUE(MID([@tk],' + S + '+4,2)),VALUE(MID([@tk],' + S + '-2,2)),'
            'VALUE(MID([@tk],' + S + '+1,2))),"")', "date", 11),
        Col("undl_parsed", "calc",
            '=IF([@parse_ok]=TRUE,TRIM(LEFT([@tk],' + S + '-3))&" "&TRIM(MID([@tk],FIND(" ",[@tk]&" ",' + S + '+8)+1,50)),"")',
            w=13),
        Col("underlying_raw", "calc", CP("underlying_raw"), w=11),
        Col("put_call_raw", "calc", CP("put_call_raw"), w=6),
        Col("strike_raw", "calc", CP("strike_raw"), "num", 10),
        Col("maturity_raw", "calc", CP("maturity_raw"), "date", 11),
        Col("pc_used", "calc", G.format('IF([@put_call_raw]<>"",UPPER(LEFT([@put_call_raw],1)),[@pc_parsed])'), w=6),
        Col("strike_used", "calc", G.format("IF(ISNUMBER([@strike_raw]),[@strike_raw],[@strike_parsed])"), "num", 10),
        Col("undl_used", "calc", G.format('IF([@underlying_raw]<>"",[@underlying_raw],[@undl_parsed])'), w=13),
        Col("maturity_used", "calc", G.format("IF(ISNUMBER([@maturity_raw]),[@maturity_raw],[@expiry_parsed])"), "date", 11),
        Col("parse_conflict", "calc",
            G.format('TRIM(IF(AND([@put_call_raw]<>"",[@pc_parsed]<>""),IF([@pc_used]<>[@pc_parsed],"P/C ",""),"")'
                     '&IF(AND(ISNUMBER([@strike_raw]),ISNUMBER([@strike_parsed])),IF(ABS([@strike_raw]-[@strike_parsed])>0.000001,"Strike ",""),"")'
                     '&IF(AND([@underlying_raw]<>"",[@undl_parsed]<>""),IF([@underlying_raw]<>[@undl_parsed],"Underlying ",""),"")'
                     '&IF(AND(ISNUMBER([@maturity_raw]),ISNUMBER([@expiry_parsed])),IF([@maturity_raw]<>[@expiry_parsed],"Maturity",""),""))'),
            w=12),
        Col("spot_series", "calc",
            '=IF(OR([@src_row]="",[@undl_used]=""),"",IF(COUNTIF(tbl_UndlMap[map_key],[@undl_used]&"|Yes")>0,'
            'INDEX(tbl_UndlMap[spot_series_id],MATCH([@undl_used]&"|Yes",tbl_UndlMap[map_key],0)),[@undl_used]))', w=12),
        Col("spot_mkt_date", "calc",
            '=IF(OR([@spot_series]="",NOT(ISNUMBER([@position_date]))),"",IF(COUNTIFS(tbl_Mkt[series_id],[@spot_series],'
            'tbl_Mkt[valid],1,tbl_Mkt[data_date],"<="&[@position_date])=0,"",MAXIFS(tbl_Mkt[data_date],tbl_Mkt[series_id],'
            '[@spot_series],tbl_Mkt[valid],1,tbl_Mkt[data_date],"<="&[@position_date])))', "date", 11),
        Col("spot_mkt", "calc",
            f'=IF(OR([@src_row]="",[@spot_mkt_date]=""),"",INDEX(tbl_Mkt[value],MATCH([@spot_series]&"|"&{DK("[@spot_mkt_date]")},tbl_Mkt[mkt_key],0)))',
            "num", 10),
        Col("spot_raw", "calc", CP("spot_raw"), "num", 10),
        Col("spot_underlying_raw", "calc", CP("spot_underlying_raw"), w=11),
        Col("spot_used", "calc", G.format("IF(ISNUMBER([@spot_raw]),[@spot_raw],[@spot_mkt])"), "num", 10),
        Col("spot_source", "calc",
            G.format('IF(ISNUMBER([@spot_raw]),"Position file",IF(ISNUMBER([@spot_mkt]),"05_Market_Input",""))'), w=12),
        Col("spot_gap_days", "calc",
            G.format('IF([@spot_source]="05_Market_Input",[@position_date]-[@spot_mkt_date],IF([@spot_source]="Position file",0,""))'),
            "int", 6),
        Col("undl_check", "calc",
            G.format('IF(ISNUMBER([@spot_raw]),IF([@spot_underlying_raw]="","Unverified",IF(OR([@spot_underlying_raw]=[@undl_used],'
                     '[@spot_underlying_raw]=[@spot_series]),"OK","MISMATCH")),IF(ISNUMBER([@spot_mkt]),"OK",""))'), w=10),
        Col("strat_type", "calc",
            G.format('IF(COUNTIF(tbl_FundMap[resolved_strategy_data_key],[@strategy_data_key])=0,"Unmapped",'
                     'INDEX(tbl_FundMap[strategy_type],MATCH([@strategy_data_key],tbl_FundMap[resolved_strategy_data_key],0)))'),
            w=10),
        Col("is_option", "calc", G.format('OR([@pc_used]="P",[@pc_used]="C")'), w=7),
        Col("calc_status", "calc",
            G.format('IF(NOT([@is_option]),IF([@strat_type]="STEP","Parse Failed / Put-Call missing","N/A (non-option)"),'
                     'IF([@undl_check]="MISMATCH","STOP: Spot/Strike underlying mismatch",'
                     'IF(ISNUMBER(SEARCH("Underlying",[@parse_conflict])),"STOP: Underlying raw vs ticker conflict",'
                     'IF([@parse_conflict]<>"","Check Required: raw vs ticker conflict",'
                     'IF(NOT(ISNUMBER([@strike_used])),"Strike Missing",IF(NOT(ISNUMBER([@spot_used])),"Spot Missing",'
                     'IF([@spot_gap_days]>thr_SpotGap,"Spot Stale",'
                     'IF([@undl_check]="Unverified","OK (spot underlying unverified)","OK"))))))))'), w=24),
        Col("calc_ok", "calc", G.format('LEFT([@calc_status],2)="OK"'), w=6),
        Col("strike_distance", "calc", '=IF([@calc_ok]=TRUE,[@strike_used]/[@spot_used]-1,"")', "pct", 9),
        Col("strike_distance_bps", "calc", '=IF([@calc_ok]=TRUE,[@strike_distance]*10000,"")', "bps", 10),
        Col("moneyness", "calc",
            '=IF([@calc_ok]=TRUE,IF([@pc_used]="P",IF([@spot_used]<[@strike_used],"ITM",IF([@spot_used]>[@strike_used],"OTM","ATM")),'
            'IF([@spot_used]>[@strike_used],"ITM",IF([@spot_used]<[@strike_used],"OTM","ATM"))),"")', w=7),
        Col("strike_buffer_abs", "calc", '=IF([@calc_ok]=TRUE,ABS([@strike_distance]),"")', "pct", 9),
        Col("otm_buffer", "calc",
            '=IF([@moneyness]="OTM",[@strike_buffer_abs],IF([@moneyness]="ATM",0,""))', "pct", 9),
        Col("buffer_display", "calc",
            '=IF([@calc_ok]=TRUE,TEXT([@strike_buffer_abs],"0.0%")&" ("&[@moneyness]&")","")', w=13),
        Col("days_to_maturity", "calc",
            '=IF(AND(ISNUMBER([@maturity_used]),ISNUMBER([@position_date])),[@maturity_used]-[@position_date],"")', "int", 7),
        Col("weight", "calc", CP("weight"), "num4", 8),
        Col("notional", "calc", CP("notional"), "amt", 12),
        Col("delta_reported", "calc", CP("delta_reported"), "num4", 8),
        Col("gamma_reported", "calc", CP("gamma_reported"), "num4", 8),
        Col("vega_reported", "calc", CP("vega_reported"), "num4", 8),
        Col("theta_reported", "calc", CP("theta_reported"), "num4", 8),
        Col("greeks_basis", "calc", CP("greeks_basis"), w=13),
        Col("delta_wadj", "calc",
            G.format('IF(NOT(AND(ISNUMBER([@delta_reported]),ISNUMBER([@weight]))),"",IF([@greeks_basis]="Per unit (weight-scalable)",'
                     '[@delta_reported]*[@weight],IF([@greeks_basis]="Position total",[@delta_reported],"")))'), "num4", 9),
        Col("vega_wadj", "calc",
            G.format('IF(NOT(AND(ISNUMBER([@vega_reported]),ISNUMBER([@weight]))),"",IF([@greeks_basis]="Per unit (weight-scalable)",'
                     '[@vega_reported]*[@weight],IF([@greeks_basis]="Position total",[@vega_reported],"")))'), "num4", 9),
        Col("greeks_note", "calc",
            G.format('IF(OR([@greeks_basis]="",[@greeks_basis]="Unknown"),"Greeks定義未確認：Weight調整・合算不可","")'), w=18),
        Col("premium_received", "calc", CP("premium_received"), "num", 10),
        Col("premium_paid", "calc", CP("premium_paid"), "num", 10),
        Col("mtm_value", "calc", CP("mtm_value"), "num", 10),
        Col("cumulative_premium", "calc", CP("cumulative_premium"), "num", 10),
        Col("position_key", "calc",
            G.format('[@strategy_data_key]&"|"&[@ticker]&"|"&IF(ISNUMBER([@date_of_entry]),TEXT([@date_of_entry],"yyyy-mm-dd"),'
                     '"NoEntryDate")&IF([@leg_id]<>"","|"&[@leg_id],"")'), w=34),
        Col("snap_dup", "calc",
            G.format('IF([@use_row]=1,COUNTIFS(§[position_key],[@position_key],§[use_row],1,§[position_date],[@position_date]),"")'),
            "int", 6),
        Col("uniqueness", "calc",
            '=IF([@snap_dup]="","",IF([@snap_dup]>1,"Check Required: non-unique key","OK"))', w=14),
        Col("latest_snap", "calc",
            G.format('IF(NOT(ISNUMBER(ctl_EvalDate)),"",IF(COUNTIFS(§[strategy_data_key],[@strategy_data_key],§[use_row],1,'
                     '§[position_date],"<="&ctl_EvalDate)=0,"",MAXIFS(§[position_date],§[strategy_data_key],[@strategy_data_key],'
                     '§[use_row],1,§[position_date],"<="&ctl_EvalDate)))'), "date", 11),
        Col("prev_snap", "calc",
            '=IF(OR([@src_row]="",[@latest_snap]=""),"",IF(COUNTIFS(§[strategy_data_key],[@strategy_data_key],§[use_row],1,'
            '§[position_date],"<"&[@latest_snap])=0,"",MAXIFS(§[position_date],§[strategy_data_key],[@strategy_data_key],'
            '§[use_row],1,§[position_date],"<"&[@latest_snap])))', "date", 11),
        Col("snap_role", "calc",
            G.format('IF([@use_row]<>1,"Not used (superseded / Portfolio)",IF([@latest_snap]="","After eval date / no eval date",'
                     'IF([@position_date]=[@latest_snap],"Current",IF([@prev_snap]="",IF([@position_date]>[@latest_snap],'
                     '"After eval date / no eval date","History"),IF([@position_date]=[@prev_snap],"Previous",'
                     'IF([@position_date]>[@latest_snap],"After eval date / no eval date","History"))))))'), w=16),
        Col("port_recon", "calc",
            G.format('IF(OR([@source_type]<>"Composition",[@use_row]<>1),"",IF(COUNTIFS(tbl_PosRaw[strategy_data_key],'
                     '[@strategy_data_key],tbl_PosRaw[position_date],[@position_date],tbl_PosRaw[source_type],"Portfolio",'
                     'tbl_PosRaw[is_latest_file],1)=0,"No Portfolio file",IF(COUNTIFS(§[position_key],[@position_key],'
                     '§[source_type],"Portfolio",§[is_latest_file],1,§[position_date],[@position_date])>0,"Matched","Not in Portfolio")))'),
            w=14),
        Col("port_weight_diff", "calc",
            '=IF([@port_recon]<>"Matched","",IF(ISNUMBER([@weight]),[@weight]-SUMIFS(§[weight],§[position_key],[@position_key],'
            '§[source_type],"Portfolio",§[is_latest_file],1,§[position_date],[@position_date]),""))', "num4", 9),
    ]
    b.tbl("08_Position_Calc", TableSpec("tbl_PosCalc", cols, nrows=CAP_POS))


# =============================================================== 09 Position change
def sheet_poschg(b: Book):
    ws = b.ws["09_Position_Change"]
    intro(ws, "09_Position_Change — 前回スナップショットとの差分",
          "同一 Strategy Data Key の直近2有効スナップショット（Current / Previous）を Position Key で比較。",
          "08_Position_Calc（数式のみ）。左＝レッグ別、右＝ファンド別サマリー。",
          "自動計算。表のフィルターで change_type を絞り込んで確認。",
          "Position Key = strategy_data_key|BloombergTicker|Date of Entry(|leg_id)。非一意は Check Required。ロールは Potential Roll（要確認）としてのみ表示。")
    G = '=IF([@src]="","",{})'
    C = lambda x: G.format(f"INDEX(tbl_PosCalc[{x}],[@src])")
    cols = [
        Col("n", "calc", "=ROW()-ROW(§[#Headers])", "int", 5),
        Col("src", "calc",
            '=IF(OR(INDEX(tbl_PosCalc[snap_role],[@n])="Current",INDEX(tbl_PosCalc[snap_role],[@n])="Previous"),[@n],"")', "int", 5),
        Col("snap_role", "calc", C("snap_role"), w=9),
        Col("strategy_data_key", "calc", C("strategy_data_key"), w=15),
        Col("position_date", "calc", C("position_date"), "date", 11),
        Col("position_key", "calc", C("position_key"), w=34),
        Col("ticker", "calc", C("ticker"), w=24),
        Col("undl_used", "calc", C("undl_used"), w=12),
        Col("pc_used", "calc", C("pc_used"), w=5),
        Col("strike_used", "calc", C("strike_used"), "num", 10),
        Col("maturity_used", "calc", C("maturity_used"), "date", 11),
        Col("weight", "calc", C("weight"), "num4", 8),
        Col("weight_num", "calc", G.format("IF(ISNUMBER([@weight]),1,0)"), "int", 5),
        Col("delta_reported", "calc", C("delta_reported"), "num4", 8),
        Col("delta_num", "calc", G.format("IF(ISNUMBER([@delta_reported]),1,0)"), "int", 5),
        Col("uniqueness", "calc", C("uniqueness"), w=12),
        Col("has_prev", "calc", G.format('INDEX(tbl_PosCalc[prev_snap],[@src])<>""'), w=6),
        Col("change_type", "calc",
            G.format('IF([@uniqueness]<>"OK","Check Required",IF([@snap_role]="Current",IF(NOT([@has_prev]),'
                     '"N/A (no previous snapshot)",IF(COUNTIFS(§[position_key],[@position_key],§[snap_role],"Previous")>0,'
                     '"Existing","New")),IF(COUNTIFS(§[position_key],[@position_key],§[snap_role],"Current")>0,"(Previous — same key in Current)",'
                     'IF(AND(ISNUMBER([@maturity_used]),[@maturity_used]<=INDEX(tbl_PosCalc[latest_snap],[@src])),'
                     '"Matured / Removed (Matured)","Matured / Removed (Removed)"))))'), w=24),
        Col("is_exit", "calc", G.format('IF(LEFT([@change_type],17)="Matured / Removed",1,0)'), "int", 5),
        Col("prev_weight", "calc",
            '=IF([@change_type]<>"Existing","",IF(COUNTIFS(§[position_key],[@position_key],§[snap_role],"Previous",§[weight_num],1)=1,'
            'SUMIFS(§[weight],§[position_key],[@position_key],§[snap_role],"Previous"),""))', "num4", 8),
        Col("weight_chg", "calc",
            '=IF(AND(ISNUMBER([@weight]),ISNUMBER([@prev_weight])),[@weight]-[@prev_weight],"")', "num4", 8),
        Col("wchg_flag", "calc", '=IF(ISNUMBER([@weight_chg]),IF(ABS([@weight_chg])>0.0000001,1,0),0)', "int", 5),
        Col("prev_delta", "calc",
            '=IF([@change_type]<>"Existing","",IF(COUNTIFS(§[position_key],[@position_key],§[snap_role],"Previous",§[delta_num],1)=1,'
            'SUMIFS(§[delta_reported],§[position_key],[@position_key],§[snap_role],"Previous"),""))', "num4", 8),
        Col("delta_chg", "calc",
            '=IF(AND(ISNUMBER([@delta_reported]),ISNUMBER([@prev_delta])),[@delta_reported]-[@prev_delta],"")', "num4", 8),
        Col("roll_flag", "calc",
            G.format('IF([@change_type]="New",IF(COUNTIFS(§[strategy_data_key],[@strategy_data_key],§[is_exit],1,§[undl_used],'
                     '[@undl_used],§[pc_used],[@pc_used])>0,"Potential Roll (要確認)",""),IF([@is_exit]=1,IF(COUNTIFS('
                     '§[strategy_data_key],[@strategy_data_key],§[change_type],"New",§[undl_used],[@undl_used],§[pc_used],[@pc_used])>0,'
                     '"Potential Roll source (要確認)",""),""))'), w=22),
    ]
    b.tbl("09_Position_Change", TableSpec("tbl_PosChg", cols, nrows=CAP_POS))
    left = len(cols) + 2
    ws.cell(6, left, "ファンド別ポジションサマリー（Strategy Data Key 単位の値を各ファンドに表示）").font = F_SECT
    ws.cell(6, 1, "レッグ別差分（Current / Previous 行のみ表示）").font = F_SECT
    G = '=IF([@fund_code]="","",{})'
    K = '=IF([@strategy_data_key]="","",{})'
    PK = "tbl_PosCalc[strategy_data_key],[@strategy_data_key]"
    CUR = PK + ',tbl_PosCalc[snap_role],"Current"'
    XK = "tbl_PosChg[strategy_data_key],[@strategy_data_key]"
    sm = [
        Col("n", "calc", "=ROW()-ROW(§[#Headers])", "int", 5),
        Col("fund_code", "calc", "=INDEX(tbl_NavCalc[fund_code],[@n])", w=11),
        Col("strategy_type", "calc", G.format("INDEX(tbl_NavCalc[strategy_type],[@n])"), w=11),
        Col("mapping_status", "calc",
            G.format('IF(INDEX(tbl_NavCalc[map_row],[@n])=0,"Mapping Required: no 03 row",'
                     'INDEX(tbl_FundMap[mapping_status],INDEX(tbl_NavCalc[map_row],[@n])))'), w=24),
        Col("strategy_data_key", "calc",
            G.format('IF(LEFT([@mapping_status],6)<>"Mapped","",INDEX(tbl_FundMap[resolved_strategy_data_key],INDEX(tbl_NavCalc[map_row],[@n])))'),
            w=15),
        Col("latest_snap", "calc",
            K.format(f'IF(COUNTIFS({CUR})=0,"",MAXIFS(tbl_PosCalc[position_date],{CUR}))'), "date", 11),
        Col("prev_snap", "calc",
            K.format(f'IF(COUNTIFS({PK},tbl_PosCalc[snap_role],"Previous")=0,"",MAXIFS(tbl_PosCalc[position_date],{PK},tbl_PosCalc[snap_role],"Previous"))'),
            "date", 11),
        Col("pos_stale_bd", "calc",
            '=IF(OR([@latest_snap]="",NOT(ISNUMBER(ctl_EvalDate))),"",NETWORKDAYS.INTL([@latest_snap],ctl_EvalDate,1,tbl_Holiday[holiday_date])-1)',
            "int", 6),
        Col("pos_status", "calc",
            G.format('IF([@strategy_data_key]="","N/A (not mapped)",IF([@latest_snap]="","Missing",IF([@pos_stale_bd]>=thr_PosStaleHigh,'
                     '"Stale (High)",IF([@pos_stale_bd]>=thr_PosStaleMed,"Stale (Medium)","OK"))))'), w=13),
        Col("n_current", "calc", K.format(f"COUNTIFS({CUR})"), "int", 6),
        Col("n_option", "calc", K.format(f"COUNTIFS({CUR},tbl_PosCalc[is_option],TRUE)"), "int", 6),
        Col("n_calc_ok", "calc", K.format(f"COUNTIFS({CUR},tbl_PosCalc[calc_ok],TRUE)"), "int", 6),
        Col("n_calc_issue", "calc",
            K.format(f'IF([@strategy_type]="STEP",[@n_current]-[@n_calc_ok],COUNTIFS({CUR},tbl_PosCalc[is_option],TRUE,tbl_PosCalc[calc_ok],FALSE))'),
            "int", 6),
        Col("n_stop", "calc", K.format(f'COUNTIFS({CUR},tbl_PosCalc[calc_status],"STOP*")'), "int", 6),
        Col("n_itm", "calc", K.format(f'COUNTIFS({CUR},tbl_PosCalc[moneyness],"ITM")'), "int", 6),
        Col("n_atm", "calc", K.format(f'COUNTIFS({CUR},tbl_PosCalc[moneyness],"ATM")'), "int", 6),
        Col("n_otm", "calc", K.format(f'COUNTIFS({CUR},tbl_PosCalc[moneyness],"OTM")'), "int", 6),
        Col("min_otm_buffer", "calc",
            K.format(f'IF(COUNTIFS({CUR},tbl_PosCalc[otm_buffer],">=0")=0,"",MINIFS(tbl_PosCalc[otm_buffer],{CUR},tbl_PosCalc[otm_buffer],">=0"))'),
            "pct", 9),
        Col("min_days_to_mat", "calc",
            K.format(f'IF(COUNTIFS({CUR},tbl_PosCalc[days_to_maturity],">=0")=0,"",MINIFS(tbl_PosCalc[days_to_maturity],{CUR},tbl_PosCalc[days_to_maturity],">=0"))'),
            "int", 7),
        Col("next_maturity", "calc",
            K.format(f'IF(COUNTIFS({CUR},tbl_PosCalc[days_to_maturity],">=0")=0,"",MINIFS(tbl_PosCalc[maturity_used],{CUR},tbl_PosCalc[days_to_maturity],">=0"))'),
            "date", 11),
        Col("n_mat_within_low", "calc",
            K.format(f'COUNTIFS({CUR},tbl_PosCalc[days_to_maturity],">=0",tbl_PosCalc[days_to_maturity],"<="&thr_Mat_Low)'), "int", 6),
        Col("n_new", "calc", K.format(f'COUNTIFS({XK},tbl_PosChg[change_type],"New")'), "int", 6),
        Col("n_existing", "calc", K.format(f'COUNTIFS({XK},tbl_PosChg[change_type],"Existing")'), "int", 6),
        Col("n_exit_matured", "calc",
            K.format(f'COUNTIFS({XK},tbl_PosChg[change_type],"Matured / Removed (Matured)")'), "int", 6),
        Col("n_exit_removed", "calc",
            K.format(f'COUNTIFS({XK},tbl_PosChg[change_type],"Matured / Removed (Removed)")'), "int", 6),
        Col("n_potential_roll", "calc", K.format(f'COUNTIFS({XK},tbl_PosChg[roll_flag],"Potential Roll (要確認)")'), "int", 6),
        Col("n_check_unique", "calc",
            K.format(f'COUNTIFS({PK},tbl_PosCalc[snap_role],"Current",tbl_PosCalc[uniqueness],"Check Required*")'), "int", 6),
        Col("n_weight_changed", "calc", K.format(f"COUNTIFS({XK},tbl_PosChg[wchg_flag],1)"), "int", 6),
        Col("delta_wadj_sum", "calc",
            K.format(f'IF([@n_current]=0,"",'
                     f'IF(SUMPRODUCT((tbl_PosCalc[strategy_data_key]=[@strategy_data_key])*(tbl_PosCalc[snap_role]="Current")*'
                     f'ISNUMBER(tbl_PosCalc[delta_wadj]))<>[@n_current],"",SUMIFS(tbl_PosCalc[delta_wadj],{CUR})))'),
            "num4", 9),
    ]
    b.tbl("09_Position_Change", TableSpec("tbl_ChgSum", sm, left=left, nrows=CAP_FUND))


# =============================================================== 10 Alerts
def sheet_alert(b: Book):
    ws = b.ws["10_Alert"]
    intro(ws, "10_Alert — アラート（Control の閾値で自動判定）",
          "ファンド×アラート種別ごとに実績値・閾値・重要度・理由・データソース・確認状況を保持する。",
          "07/09/03/02/01（数式）。確認状況・担当者・解消状況は右端の tbl_AlertReview（黄色）に alert_id 単位で入力。",
          "上表（ワイド計算）→ 下表 tbl_Alert（1アラート1行）。フィルターで severity=High 等を抽出。",
          "データ欠損は独立アラート（A01/A07/A10/A12/A14/A15）。欠損がある限り総合ステータスは GREEN になりません。")
    NC = lambda x: f"INDEX(tbl_NavCalc[{x}],[@n])"
    CS = lambda x: f"INDEX(tbl_ChgSum[{x}],[@n])"
    G = '=IF([@fund_code]="","",{})'

    def ladder(v, h, m, l):
        return f'IF({v}>={h},"High",IF({v}>={m},"Medium",IF({v}>={l},"Low","OK")))'

    defs = {
        "A01": (NC("nav_status"),
                f'IF({NC("nav_status")}="OK","OK",IF({NC("nav_status")}="Stale (Medium)","Medium","High"))',
                f'IF({NC("nav_status")}="OK","NAV基準日 "&TEXT({NC("nav_date")},"yyyy-mm-dd"),"DATA: NAV "&{NC("nav_status")}'
                f'&IF({NC("nav_date")}="",""," (NAV基準日 "&TEXT({NC("nav_date")},"yyyy-mm-dd")&", "&{NC("nav_stale_bd")}&"営業日経過)"))'),
        "A02": (NC("ret_1D"),
                f'IF({NC("nav_date")}="","N/A",IF({NC("ret_1D")}="","Low",{ladder("ABS(" + NC("ret_1D") + ")", "thr_D1_High", "thr_D1_Med", "thr_D1_Low")}))',
                f'IF({NC("nav_date")}="","NAV欠損（A01参照）",IF({NC("ret_1D")}="","DATA: 1D基準NAVなし/ギャップ超過","1D NAVリターン "&TEXT({NC("ret_1D")},"+0.00%;-0.00%")))'),
        "A03": (NC("ret_1W"),
                f'IF({NC("nav_date")}="","N/A",IF({NC("ret_1W")}="","Low",{ladder("ABS(" + NC("ret_1W") + ")", "thr_W1_High", "thr_W1_Med", "thr_W1_Low")}))',
                f'IF({NC("nav_date")}="","NAV欠損（A01参照）",IF({NC("ret_1W")}="","DATA: 1W基準NAVなし/ギャップ超過","1W NAVリターン "&TEXT({NC("ret_1W")},"+0.00%;-0.00%")))'),
        "A04": (NC("cur_dd"),
                f'IF({NC("nav_date")}="","N/A",{ladder("-MIN(" + NC("cur_dd") + ",0)", "thr_DD_High", "thr_DD_Med", "thr_DD_Low")})',
                f'IF({NC("nav_date")}="","NAV欠損（A01参照）","高値からの下落率 "&TEXT({NC("cur_dd")},"0.00%")&"（価格NAVベース）")'),
        "A05": (NC("aum_chg_1w"),
                f'IF({NC("nav_date")}="","N/A",IF({NC("aum_chg_1w")}="","Low",IF(ABS({NC("aum_chg_1w")})>=thr_AUM_High,"High",'
                f'IF(ABS({NC("aum_chg_1w")})>=thr_AUM_Med,"Medium","OK"))))',
                f'IF({NC("nav_date")}="","NAV欠損（A01参照）",IF({NC("aum_chg_1w")}="","DATA: AUM欠損","AUM 1W変化 "&TEXT({NC("aum_chg_1w")},"+0.0%;-0.0%")'
                f'&"（推定フロー・公式フローではない）"))'),
        "A06": (CS("mapping_status"),
                f'IF({CS("mapping_status")}="Mapped","OK",IF(LEFT({CS("mapping_status")},6)="Review","Medium","High"))',
                CS("mapping_status")),
        "A07": (CS("pos_status"),
                f'IF({CS("pos_status")}="OK","OK",IF({CS("pos_status")}="Stale (Medium)","Medium",IF(LEFT({CS("pos_status")},3)="N/A","N/A","High")))',
                f'IF({CS("pos_status")}="OK","Position基準日 "&TEXT({CS("latest_snap")},"yyyy-mm-dd"),"DATA: Position "&{CS("pos_status")})'),
        "A08": (CS("n_itm"),
                f'IF(OR({NC("strategy_type")}<>"STEP",{CS("latest_snap")}=""),"N/A",IF({CS("n_itm")}>0,"High","OK"))',
                f'IF({NC("strategy_type")}<>"STEP","STEP以外",IF({CS("latest_snap")}="","Positionなし（A07）","ITMレッグ "&{CS("n_itm")}&"本（ITM≠実現損失）"))'),
        "A09": (CS("min_otm_buffer"),
                f'IF(OR({NC("strategy_type")}<>"STEP",{CS("min_otm_buffer")}=""),"N/A",IF({CS("min_otm_buffer")}<=thr_Buf_High,"High",'
                f'IF({CS("min_otm_buffer")}<=thr_Buf_Med,"Medium",IF({CS("min_otm_buffer")}<=thr_Buf_Low,"Low","OK"))))',
                f'IF({CS("min_otm_buffer")}="","OTMレッグの算出値なし","最小OTMバッファ "&TEXT({CS("min_otm_buffer")},"0.0%")&" (OTM)")'),
        "A10": (CS("n_calc_issue"),
                f'IF({CS("latest_snap")}="","N/A",IF({CS("n_calc_issue")}>0,"High","OK"))',
                f'IF({CS("latest_snap")}="","Positionなし（A07）","算出不可レッグ "&{CS("n_calc_issue")}&"本（うちSTOP "&{CS("n_stop")}&"本）")'),
        "A11": (CS("min_days_to_mat"),
                f'IF({CS("min_days_to_mat")}="","N/A",IF({CS("min_days_to_mat")}<=thr_Mat_Med,"Medium",IF({CS("min_days_to_mat")}<=thr_Mat_Low,"Low","OK")))',
                f'IF({CS("min_days_to_mat")}="","満期データなし","最短満期まで "&{CS("min_days_to_mat")}&"日")'),
        "A12": (CS("n_check_unique"),
                f'IF({CS("latest_snap")}="","N/A",IF({CS("n_check_unique")}>0,"Medium","OK"))',
                f'"非一意レッグ "&N({CS("n_check_unique")})&"本"'),
        "A13": (NC("bm_label"),
                f'IF({NC("bm_label")}="Benchmark (confirmed)",IF({NC("bm_ret_1W")}="","Low","OK"),"Low")',
                f'IF({NC("bm_label")}="Benchmark (confirmed)",IF({NC("bm_ret_1W")}="","DATA: 指数データ欠損/日付不一致","OK"),{NC("bm_label")})'),
        "A14": ("INDEX(tbl_Fund[dup_count],[@n])",
                'IF(INDEX(tbl_Fund[dup_count],[@n])>1,"High","OK")',
                '"fund_code 出現回数 "&INDEX(tbl_Fund[dup_count],[@n])'),
        "A15": ('COUNTIFS(tbl_HoldAgg[fund_code],[@fund_code],tbl_HoldAgg[agg_status],"Check*")',
                'IF(COUNTIFS(tbl_HoldAgg[fund_code],[@fund_code],tbl_HoldAgg[agg_status],"Check*")>0,"Medium","OK")',
                '"要確認保有行 "&COUNTIFS(tbl_HoldAgg[fund_code],[@fund_code],tbl_HoldAgg[agg_status],"Check*")'),
    }
    cols = [Col("n", "calc", "=ROW()-ROW(§[#Headers])", "int", 5),
            Col("fund_code", "calc", "=INDEX(tbl_NavCalc[fund_code],[@n])", w=11)]
    for code, *_ in ALERTS:
        v, s, r = defs[code]
        cols += [Col(f"{code}_val", "calc", G.format(v), w=10),
                 Col(f"{code}_sev", "calc", G.format(s), w=8),
                 Col(f"{code}_reason", "calc", G.format(r), w=22)]
    sevsum = lambda lvl: "+".join(f'([@{c}_sev]="{lvl}")' for c, *_ in ALERTS)
    cols += [
        Col("n_high", "calc", G.format(sevsum("High")), "int", 6),
        Col("n_medium", "calc", G.format(sevsum("Medium")), "int", 6),
        Col("n_low", "calc", G.format(sevsum("Low")), "int", 6),
        Col("overall_status", "calc",
            G.format('IF([@n_high]>0,"RED",IF([@n_medium]>0,"AMBER",IF([@n_low]>0,"YELLOW","GREEN")))'), w=10),
        Col("n_high_unreviewed", "calc",
            G.format('COUNTIFS(tbl_Alert[fund_code],[@fund_code],tbl_Alert[severity],"High",tbl_Alert[review_status],"Unreviewed")'),
            "int", 7),
    ]
    b.tbl("10_Alert", TableSpec("tbl_AlertCalc", cols, nrows=CAP_FUND))
    ws.cell(6, 1, "ワイド計算（1ファンド1行）").font = F_SECT
    top2 = 7 + CAP_FUND + 4
    ws.cell(top2 - 1, 1, "アラート一覧（1ファンド×1アラート種別＝1行）").font = F_SECT
    ch = lambda field: "CHOOSE([@ti]," + ",".join(f"tbl_AlertCalc[{c}_{field}]" for c, *_ in ALERTS) + ")"
    lst = lambda i: ",".join(f'"{a[i]}"' for a in ALERTS)
    G2 = '=IF([@fund_code]="","",{})'
    long = [
        Col("n", "calc", "=ROW()-ROW(§[#Headers])", "int", 5),
        Col("slot", "calc", f"=INT(([@n]-1)/{T})+1", "int", 5),
        Col("ti", "calc", f"=MOD([@n]-1,{T})+1", "int", 4),
        Col("fund_code", "calc", "=INDEX(tbl_AlertCalc[fund_code],[@slot])", w=11),
        Col("alert_code", "calc", G2.format(f"CHOOSE([@ti],{lst(0)})"), w=6),
        Col("alert_name", "calc", G2.format(f"CHOOSE([@ti],{lst(1)})"), w=18),
        Col("category", "calc", G2.format(f"CHOOSE([@ti],{lst(2)})"), w=12),
        Col("data_source", "calc", G2.format(f"CHOOSE([@ti],{lst(3)})"), w=20),
        Col("actual_value", "calc", G2.format(f"INDEX({ch('val')},[@slot])"), w=12),
        Col("threshold", "calc", G2.format("CHOOSE([@ti]," + ",".join(a[4] for a in ALERTS) + ")"), w=30),
        Col("severity", "calc", G2.format(f"INDEX({ch('sev')},[@slot])"), w=8),
        Col("reason", "calc", G2.format(f"INDEX({ch('reason')},[@slot])"), w=36),
        Col("eval_date", "calc", G2.format('IF(ISNUMBER(ctl_EvalDate),ctl_EvalDate,"")'), "date", 11),
        Col("judged_at", "calc", G2.format("NOW()"), "dt", 15),
        Col("alert_id", "calc",
            G2.format('[@fund_code]&"|"&[@alert_code]&"|"&IF(ISNUMBER(ctl_EvalDate),TEXT(ctl_EvalDate,"yyyymmdd"),"NOEVAL")'), w=24),
        Col("review_status", "calc",
            G2.format('IF(COUNTIF(tbl_AlertReview[alert_id],[@alert_id])=0,"Unreviewed",'
                      'IF(INDEX(tbl_AlertReview[review_status],MATCH([@alert_id],tbl_AlertReview[alert_id],0))="","Unreviewed",'
                      'INDEX(tbl_AlertReview[review_status],MATCH([@alert_id],tbl_AlertReview[alert_id],0))))'), w=12),
        Col("reviewer", "calc",
            G2.format('IF(COUNTIF(tbl_AlertReview[alert_id],[@alert_id])=0,"",INDEX(tbl_AlertReview[reviewer],'
                      'MATCH([@alert_id],tbl_AlertReview[alert_id],0))&"")'), w=10),
        Col("resolution", "calc",
            G2.format('IF(COUNTIF(tbl_AlertReview[alert_id],[@alert_id])=0,"Open",IF(INDEX(tbl_AlertReview[resolution],'
                      'MATCH([@alert_id],tbl_AlertReview[alert_id],0))="","Open",INDEX(tbl_AlertReview[resolution],'
                      'MATCH([@alert_id],tbl_AlertReview[alert_id],0))))'), w=10),
    ]
    b.tbl("10_Alert", TableSpec("tbl_Alert", long, top=top2, nrows=CAP_FUND * T))
    left = len(long) + 2
    rv = [Col("alert_id", "input", w=24), Col("review_status", "input", w=12, dv=DV["arev"]),
          Col("reviewer", "input", w=10), Col("reviewed_at", "input", fmt="date", w=11),
          Col("resolution", "input", w=10, dv=DV["ares"]), Col("note", "input", w=24)]
    b.tbl("10_Alert", TableSpec("tbl_AlertReview", rv, top=top2, left=left, rows=b.rows("alertreview")))
    ws.cell(top2 - 1, left, "確認記録（alert_id をコピーして入力。評価日が変わると新しい alert_id）").font = F_SECT
    add_status_cf(ws, f"A8:CZ{7 + CAP_FUND}")
    add_status_cf(ws, f"K{top2 + 1}:K{top2 + CAP_FUND * T}")
    ws.freeze_panes = "C8"


# =============================================================== 11 Coverage
def sheet_cov(b: Book):
    ws = b.ws["11_Update_Coverage"]
    intro(ws, "11_Update_Coverage — ファンド別データ更新状況・データゲート",
          "各ファンドの NAV / Mapping / Position / 指数 / アラートの充足状況と、配信前のデータゲートを一覧化。",
          "07/09/10/01（数式のみ）。",
          "自動計算。missing_items と data_gate を確認し、欠損元シートを修正する。",
          "data_gate：DATA REQUIRED＝必須データ欠損・計算不能、REVIEW REQUIRED＝確認未了、READY＝データ面の条件充足（配信可否は 12/99 で別判定）。")
    NC = lambda x: f"INDEX(tbl_NavCalc[{x}],[@n])"
    CS = lambda x: f"INDEX(tbl_ChgSum[{x}],[@n])"
    AC = lambda x: f"INDEX(tbl_AlertCalc[{x}],[@n])"
    G = '=IF([@fund_code]="","",{})'
    cols = [
        Col("n", "calc", "=ROW()-ROW(§[#Headers])", "int", 5),
        Col("fund_code", "calc", "=INDEX(tbl_NavCalc[fund_code],[@n])", w=11),
        Col("fund_name", "calc", G.format(NC("fund_name")), w=24),
        Col("strategy_type", "calc", G.format(NC("strategy_type")), w=11),
        Col("investor_count", "calc", G.format("INDEX(tbl_Fund[investor_count],[@n])"), "int", 7),
        Col("total_amount", "calc", G.format("INDEX(tbl_Fund[total_amount],[@n])"), "amt", 15),
        Col("mapping_status", "calc", G.format(CS("mapping_status")), w=24),
        Col("nav_date", "calc", G.format(NC("nav_date")), "date", 11),
        Col("nav_status", "calc", G.format(NC("nav_status")), w=13),
        Col("pos_snap_date", "calc", G.format(CS("latest_snap")), "date", 11),
        Col("pos_status", "calc", G.format(CS("pos_status")), w=13),
        Col("pos_calc_issues", "calc", G.format(CS("n_calc_issue")), "int", 7),
        Col("bm_label", "calc", G.format(NC("bm_label")), w=22),
        Col("alert_overall", "calc", G.format(AC("overall_status")), w=9),
        Col("high_unreviewed", "calc", G.format(AC("n_high_unreviewed")), "int", 7),
        Col("missing_items", "calc",
            G.format('TEXTJOIN(", ",TRUE,IF(NOT(ISNUMBER(ctl_EvalDate)),"評価日",""),IF([@nav_status]="Missing","NAV",""),'
                     'IF(LEFT([@nav_status],5)="Stale","NAV(遅延)",""),IF(' + AC("A06_sev") + '<>"OK","Mapping",""),'
                     'IF([@pos_status]="Missing","Position",""),IF(LEFT([@pos_status],5)="Stale","Position(遅延)",""),'
                     'IF(N([@pos_calc_issues])>0,"Position計算("&[@pos_calc_issues]&")",""),'
                     'IF(' + AC("A12_sev") + '="Medium","Position Key非一意",""),'
                     'IF(' + AC("A14_sev") + '="High","fund_code重複",""),IF(' + AC("A15_sev") + '<>"OK","保有データ",""),'
                     'IF(' + AC("A13_sev") + '<>"OK","参照指数",""))'), w=30),
        Col("coverage_status", "calc",
            G.format('IF([@missing_items]="","Complete",IF(AND([@nav_status]="Missing",OR([@pos_status]="Missing",'
                     'LEFT([@pos_status],3)="N/A")),"Missing","Partial"))'), w=10),
        Col("data_gate", "calc",
            G.format('IF(OR(NOT(ISNUMBER(ctl_EvalDate)),' + AC("A01_sev") + '="High",' + AC("A06_sev") + '="High",'
                     + AC("A07_sev") + '="High",' + AC("A10_sev") + '="High",' + AC("A14_sev") + '="High",'
                     + AC("A15_sev") + '<>"OK"),"DATA REQUIRED",'
                     'IF(OR(' + AC("A01_sev") + '="Medium",' + AC("A06_sev") + '="Medium",' + AC("A07_sev") + '="Medium",'
                     + AC("A12_sev") + '<>"OK",[@high_unreviewed]>0,ctl_ThrConfirmed<>"Yes"),"REVIEW REQUIRED","READY"))'),
            w=16),
    ]
    # A12 sev may be N/A when no snapshot -> treat N/A as not blocking
    cols[-1].f = cols[-1].f.replace(AC("A12_sev") + '<>"OK"', AC("A12_sev") + '="Medium"')
    b.tbl("11_Update_Coverage", TableSpec("tbl_Cov", cols, nrows=CAP_FUND))
    add_status_cf(ws, f"N8:R{7 + CAP_FUND}")


# =============================================================== 12 Distribution
def sheet_dist(b: Book):
    ws = b.ws["12_Distribution_List"]
    intro(ws, "12_Distribution_List — 配信管理（投資家単位 ＋ 投資家×ファンド明細）",
          "投資家ごとの保有ファンド・Amount・宛先・配信方式（個別/統合）・配信ステータス・レビュー状況を管理。",
          "左表 tbl_Dist：営業入力（黄色）＋計算（青）。右表 tbl_DistDetail：02集約ビューの各行の配信判定（数式）。",
          "投資家を追加したら左表に end_investor_bdr_id を1行追加し宛先等を入力。メールアドレス不明は空欄のまま（DATA REQUIRED 表示）。",
          "Excel からの自動メール送信は実装していません。配信は READY の行のみ、final_text を手動で使用してください。")
    D = "tbl_DistDetail[end_investor_bdr_id],[@end_investor_bdr_id]"
    G = '=IF([@end_investor_bdr_id]="","",{})'
    HR = "INDEX(tbl_HoldRaw[{}],MATCH([@end_investor_bdr_id],tbl_HoldRaw[end_investor_bdr_id],0))"
    cols = [
        Col("end_investor_bdr_id", "input", w=14), Col("to_email", "input", w=24), Col("cc_email", "input", w=24),
        Col("sales_rep_override", "input", w=12), Col("distribution_mode", "input", w=11, dv=DV["mode"]),
        Col("last_sent_date", "input", fmt="date", w=11), Col("next_scheduled_date", "input", fmt="date", w=11),
        Col("distribution_status", "input", w=11, dv=DV["diststat"]), Col("review_status", "input", w=11, dv=DV["rev"]),
        Col("note", "input", w=16),
        Col("in_holdings", "calc",
            G.format('IF(COUNTIF(tbl_HoldRaw[end_investor_bdr_id],[@end_investor_bdr_id])>0,"Yes","No")'), w=8),
        Col("investor_name", "calc", G.format('IF([@in_holdings]="No","",' + HR.format("end_investor_name") + '&"")'), w=22),
        Col("sales_rep", "calc",
            G.format('IF([@sales_rep_override]<>"",[@sales_rep_override],IF([@in_holdings]="No","",' + HR.format("sales_rep") + '&""))'),
            w=12),
        Col("fund_count", "calc", G.format(f"COUNTIFS({D})"), "int", 6),
        Col("fund_list", "calc",
            G.format('TEXTJOIN(", ",TRUE,IF(tbl_DistDetail[end_investor_bdr_id]=[@end_investor_bdr_id],tbl_DistDetail[fund_code],""))'),
            w=26, arr=True),
        Col("amount_total", "calc",
            G.format('IF(COUNTIFS(tbl_HoldAgg[end_investor_bdr_id],[@end_investor_bdr_id],tbl_HoldAgg[amount_total],"")>0,"",'
                     'SUMIFS(tbl_HoldAgg[amount_num],tbl_HoldAgg[end_investor_bdr_id],[@end_investor_bdr_id]))'), "amt", 15),
        Col("update_target_funds", "calc",
            G.format('TEXTJOIN(", ",TRUE,IF((tbl_DistDetail[end_investor_bdr_id]=[@end_investor_bdr_id])*'
                     '(tbl_DistDetail[fund_data_gate]<>"DATA REQUIRED"),tbl_DistDetail[fund_code],""))'), w=22, arr=True),
        Col("n_ready", "calc", G.format(f'COUNTIFS({D},tbl_DistDetail[gate],"READY")'), "int", 6),
        Col("n_review", "calc", G.format(f'COUNTIFS({D},tbl_DistDetail[gate],"REVIEW REQUIRED")'), "int", 6),
        Col("n_data", "calc", G.format(f'COUNTIFS({D},tbl_DistDetail[gate],"DATA REQUIRED")'), "int", 6),
        Col("email_status", "calc", G.format('IF([@to_email]="","宛先未登録","OK")'), w=10),
        Col("distribution_gate", "calc",
            G.format('IF([@fund_count]=0,"DATA REQUIRED",IF([@distribution_mode]="Combined",IF([@n_data]>0,"DATA REQUIRED",'
                     'IF([@n_review]>0,"REVIEW REQUIRED","READY")),IF([@distribution_mode]="Individual",'
                     'IF([@n_ready]>0,"READY",IF([@n_review]>0,"REVIEW REQUIRED","DATA REQUIRED")),"REVIEW REQUIRED")))'), w=16),
        Col("gate_note", "calc",
            G.format('IF([@fund_count]=0,"02に保有なし",IF([@distribution_mode]="","配信方式未選択",'
                     'IF([@distribution_mode]="Individual","個別配信：READY "&[@n_ready]&" / REVIEW "&[@n_review]&" / DATA "&[@n_data]&"（READY行のみ配信可）",'
                     '"統合配信：全ファンドREADYで配信可")))'), w=36),
    ]
    b.tbl("12_Distribution_List", TableSpec("tbl_Dist", cols, rows=b.rows("dist")))
    left = len(cols) + 2
    ws.cell(6, left, "投資家×ファンド配信判定（02 集約ビューと同じ行順・数式）").font = F_SECT
    ws.cell(6, 1, "投資家別 配信管理").font = F_SECT
    G = '=IF([@holding_key]="","",{})'
    HA = lambda x: f'IF(INDEX(tbl_HoldAgg[{x}],[@n])="","",INDEX(tbl_HoldAgg[{x}],[@n]))'
    CID = ('[@end_investor_bdr_id]&"|"&[@fund_code]&"|"&IF(ISNUMBER(ctl_EvalDate),TEXT(ctl_EvalDate,"yyyymmdd"),"NOEVAL")'
           '&"|N"&IF(ISNUMBER([@nav_date]),TEXT([@nav_date],"yyyymmdd"),"NA")&"|P"&IF(ISNUMBER([@pos_date]),TEXT([@pos_date],"yyyymmdd"),"NA")')
    det = [
        Col("n", "calc", "=ROW()-ROW(§[#Headers])", "int", 5),
        Col("holding_key", "calc", "=" + HA("holding_key"), w=20),
        Col("end_investor_bdr_id", "calc", G.format(HA("end_investor_bdr_id")), w=14),
        Col("fund_code", "calc", G.format(HA("fund_code")), w=11),
        Col("amount_total", "calc", G.format(HA("amount_total")), "amt", 15),
        Col("holding_status", "calc", G.format(HA("agg_status")), w=20),
        Col("fund_row", "calc",
            G.format("IF(COUNTIF(tbl_Cov[fund_code],[@fund_code])=0,0,MATCH([@fund_code],tbl_Cov[fund_code],0))"), "int", 5),
        Col("fund_data_gate", "calc",
            G.format('IF([@fund_row]=0,"DATA REQUIRED",INDEX(tbl_Cov[data_gate],[@fund_row]))'), w=16),
        Col("nav_date", "calc", G.format('IF([@fund_row]=0,"",INDEX(tbl_NavCalc[nav_date],[@fund_row]))'), "date", 11),
        Col("pos_date", "calc", G.format('IF([@fund_row]=0,"",INDEX(tbl_ChgSum[latest_snap],[@fund_row]))'), "date", 11),
        Col("comment_id", "calc", G.format(CID), w=40),
        Col("comment_status", "calc",
            G.format('IF(COUNTIF(tbl_CommentReview[comment_id],[@comment_id])=0,"No review record",'
                     'INDEX(tbl_CommentReview[review_status],MATCH([@comment_id],tbl_CommentReview[comment_id],0))&"")'), w=14),
        Col("final_text_ready", "calc",
            G.format('IF(COUNTIF(tbl_CommentReview[comment_id],[@comment_id])<>1,"No",IF(COUNTIFS(tbl_CommentReview[comment_id],'
                     '[@comment_id],tbl_CommentReview[review_status],"Approved",tbl_CommentReview[reviewed_text],"<>",'
                     'tbl_CommentReview[approver],"<>")=1,"Yes","No"))'), w=8),
        Col("dist_row", "calc",
            G.format("IF(COUNTIF(tbl_Dist[end_investor_bdr_id],[@end_investor_bdr_id])=0,0,"
                     "MATCH([@end_investor_bdr_id],tbl_Dist[end_investor_bdr_id],0))"), "int", 5),
        Col("email_present", "calc",
            G.format('IF([@dist_row]=0,"No",IF(INDEX(tbl_Dist[to_email],[@dist_row])="","No","Yes"))'), w=7),
        Col("distribution_mode", "calc",
            G.format('IF([@dist_row]=0,"",INDEX(tbl_Dist[distribution_mode],[@dist_row])&"")'), w=10),
        Col("gate", "calc",
            G.format('IF(OR(ctl_GlobalGate="DATA REQUIRED",[@holding_status]<>"OK",[@fund_data_gate]="DATA REQUIRED",'
                     '[@email_present]<>"Yes"),"DATA REQUIRED",IF(OR(ctl_GlobalGate="REVIEW REQUIRED",'
                     '[@fund_data_gate]="REVIEW REQUIRED",[@final_text_ready]<>"Yes"),"REVIEW REQUIRED","READY"))'), w=16),
        Col("gate_reason", "calc",
            G.format('TEXTJOIN(" / ",TRUE,IF(ctl_GlobalGate<>"READY","全体:"&ctl_GlobalGate,""),IF([@holding_status]<>"OK",'
                     '[@holding_status],""),IF([@fund_data_gate]<>"READY","ファンド:"&[@fund_data_gate],""),'
                     'IF([@email_present]<>"Yes","宛先未登録",""),IF([@final_text_ready]<>"Yes","コメント未承認",""))'), w=40),
    ]
    b.tbl("12_Distribution_List", TableSpec("tbl_DistDetail", det, left=left, nrows=CAP_HOLD))
    add_status_cf(ws, "A8:AZ2000")


# =============================================================== view helpers
def put(ws, addr, value, kind="calc", fmt=None, font=None, name=None, b=None, arr=False, wrap=False):
    if isinstance(value, str) and value.startswith("="):
        value = xlfn(value)
        if arr:
            ws[addr] = ArrayFormula(addr, value)
        else:
            ws[addr] = value
    else:
        ws[addr] = value
    c = ws[addr]
    if kind:
        c.fill = FILL[kind]
        c.border = BORDER
    c.font = font or F_BODY
    if fmt:
        c.number_format = FMT.get(fmt, fmt)
    c.alignment = Alignment(vertical="top", wrap_text=wrap)
    if name and b:
        col = "".join(ch for ch in addr if ch.isalpha())
        row = "".join(ch for ch in addr if ch.isdigit())
        b.name(name, ws.title, f"${col}${row}")
    return c


def label(ws, addr, text, sect=False):
    c = ws[addr]
    c.value = text
    c.font = F_SECT if sect else F_NOTE
    if not sect:
        c.fill = FILL["sect"]
        c.border = BORDER
    c.alignment = Alignment(vertical="top", wrap_text=True)
    return c


def DSH(table, col, row="ctl_SelRow"):
    return (f'=IF({row}=0,"—",IF(INDEX({table}[{col}],{row})="","—",INDEX({table}[{col}],{row})))')


def P(c):
    return f'IF(ISNUMBER({c}),TEXT({c},"+0.00%;-0.00%;0.00%"),"[データなし]")'


def Dt(c):
    return f'IF(ISNUMBER({c}),TEXT({c},"yyyy/mm/dd"),"[データなし]")'


# =============================================================== 13 Dashboard
def sheet_dashboard(b: Book):
    ws = b.ws["13_Dashboard"]
    ws.sheet_view.showGridLines = False
    ws["B1"] = "QIS Fund Monitoring — Sales Dashboard（1投資家 × 1ファンド）"
    ws["B1"].font = Font(name="Yu Gothic", size=14, bold=True, color="1F2933")
    ws["B2"] = ("目的：顧客向け定期アップデートの確認用。データソース：07/09/10/11/14（数式のみ）。更新方法：00_Control で評価日・投資家・"
                "fund_code を選択。注意：数値は Comment Output と同一セルを参照。『—』はデータなし（0ではない）。")
    ws["B2"].font = F_NOTE
    ws["B2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("B2:M3")
    for col, w in zip("ABCDEFGHIJKLMN", [3, 20, 11, 11, 11, 11, 11, 11, 11, 11, 13, 11, 13, 3]):
        ws.column_dimensions[col].width = w

    # ---- header block (rows 5-8)
    hdr1 = [
        ("B", "投資家名", '=IF(ctl_SelInvestor="","（未選択）",IF(COUNTIF(tbl_HoldRaw[end_investor_bdr_id],ctl_SelInvestor)=0,'
                        '"（02に該当投資家なし）",INDEX(tbl_HoldRaw[end_investor_name],MATCH(ctl_SelInvestor,tbl_HoldRaw[end_investor_bdr_id],0))&""))',
         None, "dsh_InvName"),
        ("D", "ファンド名", DSH("tbl_NavCalc", "fund_name"), None, "dsh_FundName"),
        ("F", "Fund Code", '=IF(ctl_SelFund="","（未選択）",ctl_SelFund)', None, "dsh_FundCode"),
        ("H", "Strategy Type", DSH("tbl_NavCalc", "strategy_type"), None, "dsh_StratType"),
        ("J", "Index", DSH("tbl_NavCalc", "index_name"), None, "dsh_Index"),
        ("L", "評価日", '=IF(ISNUMBER(ctl_EvalDate),ctl_EvalDate,"未設定")', "date", "dsh_EvalDate"),
    ]
    hdr2 = [
        ("B", "NAV基準日", DSH("tbl_NavCalc", "nav_date"), "date", "dsh_NavDate"),
        ("D", "Position基準日", DSH("tbl_ChgSum", "latest_snap"), "date", "dsh_PosDate"),
        ("F", "データ更新ステータス", '=IF(ctl_SelRow=0,"—",INDEX(tbl_Cov[coverage_status],ctl_SelRow)&" / "&INDEX(tbl_Cov[data_gate],ctl_SelRow))',
         None, "dsh_DataStatus"),
        ("H", "アラート総合", DSH("tbl_AlertCalc", "overall_status"), None, "dsh_Alert"),
        ("J", "配信判定", "=ctl_SelGate", None, "dsh_Gate"),
        ("L", "保有Amount", '=IF(ctl_SelHoldRow=0,"保有なし",IF(INDEX(tbl_HoldAgg[amount_total],ctl_SelHoldRow)="","要確認",'
                          'INDEX(tbl_HoldAgg[amount_total],ctl_SelHoldRow)))', "amt", "dsh_Amount"),
    ]
    for row, items in [(5, hdr1), (7, hdr2)]:
        for col, lab, f, fmt, nm in items:
            nxt = chr(ord(col) + 1)
            label(ws, f"{col}{row}", lab)
            ws.merge_cells(f"{col}{row}:{nxt}{row}")
            put(ws, f"{col}{row + 1}", f, fmt=fmt, font=F_BOLD, name=nm, b=b)
            ws.merge_cells(f"{col}{row + 1}:{nxt}{row + 1}")
    add_status_cf(ws, "B6:M8")

    # ---- performance (rows 10-16)
    label(ws, "B10", "NAV・パフォーマンス（ファンド固有NAV時系列／基準日以前の直近有効NAV）", sect=True)
    pcols = ["1D", "1W", "MTD", "3M", "YTD", "1Y", "SI"]
    label(ws, "B11", "項目")
    for j, code in enumerate(pcols):
        label(ws, f"{chr(ord('C') + j)}11", code)
    rowsdef = [
        (12, "Fund NAVリターン", lambda c: DSH("tbl_NavCalc", f"ret_{c}"), "pct", "dsh_Ret_"),
        (13, "基準NAV日", lambda c: DSH("tbl_NavCalc", "first_nav_date" if c == "SI" else f"base_{c}"), "date", "dsh_Base_"),
        (14, "参照指数リターン（同一期間）", lambda c: '="—"' if c == "SI" else DSH("tbl_NavCalc", f"bm_ret_{c}"), "pct", "dsh_Bm_"),
        (15, "差（Fund − Benchmark）", lambda c: '="—"' if c == "SI" else
         f'=IF(AND(dsh_BmLabel="Benchmark (confirmed)",ISNUMBER(dsh_Ret_{c}),ISNUMBER(dsh_Bm_{c})),dsh_Ret_{c}-dsh_Bm_{c},"—")',
         "pct", "dsh_Rel_"),
        (16, "Total Return（定義確認済のみ）", lambda c: DSH("tbl_NavCalc", f"tr_{c}") if c in ("YTD", "1Y", "SI") else '="—"',
         "pct", "dsh_TR_"),
    ]
    for r, lab, fn, fmt, prefix in rowsdef:
        label(ws, f"B{r}", lab)
        for j, code in enumerate(pcols):
            put(ws, f"{chr(ord('C') + j)}{r}", fn(code), fmt=fmt, name=f"{prefix}{code}", b=b)
    put(ws, "B17", DSH("tbl_NavCalc", "bm_label"), name="dsh_BmLabel", b=b, font=F_NOTE)
    put(ws, "C17", DSH("tbl_NavCalc", "bm_series"), name="dsh_BmSeries", b=b, font=F_NOTE)
    put(ws, "D17", DSH("tbl_NavCalc", "tr_status"), name="dsh_TRStatus", b=b, font=F_NOTE)
    ws.merge_cells("D17:I17")
    stats = [
        (11, "NAV", DSH("tbl_NavCalc", "nav"), "nav", "dsh_Nav"),
        (12, "Peak NAV", DSH("tbl_NavCalc", "peak_nav"), "nav", "dsh_Peak"),
        (13, "高値からの下落率", DSH("tbl_NavCalc", "cur_dd"), "pct", "dsh_CurDD"),
        (14, "最大ドローダウン", DSH("tbl_NavCalc", "max_dd"), "pct", "dsh_MaxDD"),
        (15, "AUM（公式）", DSH("tbl_NavCalc", "aum"), "amt", "dsh_AUM"),
        (16, "AUM 1W変化", DSH("tbl_NavCalc", "aum_chg_1w"), "pct", "dsh_AUMChg"),
        (17, "推定資金流出入1W（非公式）", DSH("tbl_NavCalc", "est_flow_1w"), "amt", "dsh_EstFlow"),
        (18, "YTD分配金（受領データ合計）", DSH("tbl_NavCalc", "dist_ytd"), "nav", "dsh_DistYTD"),
        (19, "通貨", DSH("tbl_NavCalc", "currency"), None, "dsh_Ccy"),
    ]
    for r, lab, f, fmt, nm in stats:
        label(ws, f"K{r}", lab)
        ws.merge_cells(f"K{r}:L{r}")
        put(ws, f"M{r}", f, fmt=fmt, name=nm, b=b)

    # ---- positions (rows 21-37)
    put(ws, "B21", '="ポジション状況（Strategy Data Key："&IF(ctl_SelRow=0,"—",INDEX(tbl_ChgSum[strategy_data_key],ctl_SelRow)&"")&"）"',
        kind=None, font=F_SECT)
    put(ws, "A21", '=IF(ctl_SelRow=0,"",INDEX(tbl_ChgSum[strategy_data_key],ctl_SelRow)&"")', kind=None, font=F_HELP,
        name="dsh_StratKey", b=b)
    summ = [
        ("B", "前回基準日", DSH("tbl_ChgSum", "prev_snap"), "date", "dsh_PrevPos"),
        ("C", "レッグ数", DSH("tbl_ChgSum", "n_current"), "int", "dsh_NLegs"),
        ("D", "ITM", DSH("tbl_ChgSum", "n_itm"), "int", "dsh_NITM"),
        ("E", "ATM", DSH("tbl_ChgSum", "n_atm"), "int", "dsh_NATM"),
        ("F", "OTM", DSH("tbl_ChgSum", "n_otm"), "int", "dsh_NOTM"),
        ("G", "最小OTMバッファ", DSH("tbl_ChgSum", "min_otm_buffer"), "pct", "dsh_MinBuf"),
        ("H", "最短満期(日)", DSH("tbl_ChgSum", "min_days_to_mat"), "int", "dsh_MinDays"),
        ("I", "新規", DSH("tbl_ChgSum", "n_new"), "int", "dsh_NNew"),
        ("J", "満期/削除", '=IF(ctl_SelRow=0,"—",IF(INDEX(tbl_ChgSum[strategy_data_key],ctl_SelRow)="","—",'
                         'INDEX(tbl_ChgSum[n_exit_matured],ctl_SelRow)&" / "&INDEX(tbl_ChgSum[n_exit_removed],ctl_SelRow)))', None,
         "dsh_NExit"),
        ("K", "Potential Roll", DSH("tbl_ChgSum", "n_potential_roll"), "int", "dsh_NRoll"),
        ("L", "算出不可", DSH("tbl_ChgSum", "n_calc_issue"), "int", "dsh_NIssue"),
        ("M", "Position状態", DSH("tbl_ChgSum", "pos_status"), None, "dsh_PosStatus"),
    ]
    for col, lab, f, fmt, nm in summ:
        label(ws, f"{col}22", lab)
        put(ws, f"{col}23", f, fmt=fmt, name=nm, b=b)
    for nm, col in [("dsh_NExitMat", "n_exit_matured"), ("dsh_NExitRem", "n_exit_removed"), ("dsh_NExisting", "n_existing"),
                    ("dsh_NCheck", "n_check_unique"), ("dsh_NextMat", "next_maturity"), ("dsh_NMatLow", "n_mat_within_low"),
                    ("dsh_NStop", "n_stop")]:
        pass
    hidden = [("dsh_NExitMat", "n_exit_matured", None), ("dsh_NExitRem", "n_exit_removed", None),
              ("dsh_NExisting", "n_existing", None), ("dsh_NCheck", "n_check_unique", None),
              ("dsh_NextMat", "next_maturity", "date"), ("dsh_NMatLow", "n_mat_within_low", None),
              ("dsh_NStop", "n_stop", None)]
    for k, (nm, col, fmt) in enumerate(hidden):
        put(ws, f"P{40 + k}", DSH("tbl_ChgSum", col), fmt=fmt, name=nm, b=b, font=F_HELP)
        ws[f"O{40 + k}"] = nm
        ws[f"O{40 + k}"].font = F_HELP
    legh = ["Ticker", "P/C", "Strike", "Spot", "Strike Dist.", "Moneyness", "Buffer（表示）", "Maturity", "残日数",
            "Weight", "Delta(報告値)", "状態 / 変更"]
    legc = ["ticker", "pc_used", "strike_used", "spot_used", "strike_distance", "moneyness", "buffer_display",
            "maturity_used", "days_to_maturity", "weight", "delta_reported", None]
    legf = [None, None, "num", "num", "pct", None, None, "date", "int", "num4", "num4", None]
    for j, h in enumerate(legh):
        label(ws, f"{chr(ord('B') + j)}25", h)
    for k in range(1, N_LEGS + 1):
        r = 25 + k
        put(ws, f"A{r}",
            f'=IF(OR(dsh_StratKey="",{k}>COUNTIFS(tbl_PosCalc[strategy_data_key],dsh_StratKey,tbl_PosCalc[snap_role],"Current")),"",'
            f'SMALL(IF((tbl_PosCalc[strategy_data_key]=dsh_StratKey)*(tbl_PosCalc[snap_role]="Current"),tbl_PosCalc[n]),{k}))',
            kind=None, font=F_HELP, arr=True)
        for j, (c, fmt) in enumerate(zip(legc, legf)):
            col = chr(ord("B") + j)
            if c:
                f = f'=IF($A{r}="","",INDEX(tbl_PosCalc[{c}],$A{r}))'
            else:
                f = (f'=IF($A{r}="","",IF(INDEX(tbl_PosCalc[calc_ok],$A{r})=TRUE,"","["&INDEX(tbl_PosCalc[calc_status],$A{r})&"] ")'
                     f'&INDEX(tbl_PosChg[change_type],$A{r})&IF(INDEX(tbl_PosChg[roll_flag],$A{r})="",""," "&INDEX(tbl_PosChg[roll_flag],$A{r})))')
            put(ws, f"{col}{r}", f, fmt=fmt)
    r = 26 + N_LEGS
    put(ws, f"B{r}", f'=IF(N(dsh_NLegs)>{N_LEGS},"※ 表示は先頭{N_LEGS}レッグのみ。全件は 08_Position_Calc を参照。","")',
        kind=None, font=F_NOTE)
    put(ws, f"B{r + 1}", '="※ Strike Dist.＝Strike/Spot−1（符号付き）。Buffer は絶対値で ITM/OTM を併記。ITM は評価時点の状態であり実現損失を意味しません。'
                         'Greeks は報告値（単位・定義は 06 greeks_basis）でファンド全体のリスク量ではありません。"', kind=None, font=F_NOTE)
    ws.merge_cells(f"B{r + 1}:M{r + 1}")
    ws.row_dimensions[r + 1].height = 26
    ws[f"B{r + 1}"].alignment = Alignment(wrap_text=True, vertical="top")

    # ---- summary (bottom)
    s0 = r + 3
    put(ws, f"B{s0}", "営業向けサマリー（14_Comment_Output のドラフトと同一。承認前は顧客送付不可）", kind=None, font=F_SECT)
    for k, (lab, nm) in enumerate([("今週のポイント", "cmt_Points"), ("ファンドの状況", "cmt_Fund"),
                                   ("現在のポジション", "cmt_Position"), ("データ注意事項", "cmt_Notes")]):
        rr = s0 + 1 + k
        label(ws, f"B{rr}", lab)
        put(ws, f"C{rr}", f"={nm}", wrap=True)
        ws.merge_cells(f"C{rr}:M{rr}")
        ws.row_dimensions[rr].height = 58
    last = s0 + 4

    # ---- chart data (columns R..W)
    ws["R4"] = "Chart data（NAV と参照指数、起点=100。#N/A は欠損＝描画しない）"
    ws["R4"].font = F_NOTE
    for j, h in enumerate(["i", "date", "NAV", "Ref index", "Fund (=100)", "Ref (=100)"]):
        label(ws, f"{chr(ord('R') + j)}5", h)
    put(ws, "R3", ('=IF(OR(ctl_SelFund="",NOT(ISNUMBER(dsh_NavDate))),0,COUNTIFS(tbl_NAV[fund_code],ctl_SelFund,tbl_NAV[valid],1,'
                   'tbl_NAV[nav_date],"<="&dsh_NavDate))'), name="dsh_ChartCnt", b=b, fmt="int")
    first, lastr = 6, 6 + N_CHART - 1
    BMS = 'IF(dsh_BmSeries="—","",dsh_BmSeries)'
    for i in range(1, N_CHART + 1):
        r = 5 + i
        put(ws, f"R{r}", i, kind=None, font=F_HELP)
        put(ws, f"S{r}",
            f'=IF({N_CHART + 1 - i}>dsh_ChartCnt,NA(),LARGE(IF((tbl_NAV[fund_code]=ctl_SelFund)*(tbl_NAV[valid]=1)*'
            f'(tbl_NAV[nav_date]<=dsh_NavDate),tbl_NAV[nav_date]),{N_CHART + 1 - i}))', fmt="date", arr=True)
        put(ws, f"T{r}", f'=IF(ISNA(S{r}),NA(),INDEX(tbl_NAV[nav],MATCH(ctl_SelFund&"|"&{DK(f"S{r}")},tbl_NAV[nav_key],0)))', fmt="nav")
        cnt = f'COUNTIFS(tbl_Mkt[series_id],{BMS},tbl_Mkt[valid],1,tbl_Mkt[data_date],"<="&S{r})'
        mx = f'MAXIFS(tbl_Mkt[data_date],tbl_Mkt[series_id],{BMS},tbl_Mkt[valid],1,tbl_Mkt[data_date],"<="&S{r})'
        put(ws, f"U{r}",
            f'=IF(ISNA(S{r}),NA(),IF({BMS}="",NA(),IF({cnt}=0,NA(),IF(S{r}-{mx}>ctl_MaxBaseGap,NA(),'
            f'INDEX(tbl_Mkt[value],MATCH({BMS}&"|"&{DK(mx)},tbl_Mkt[mkt_key],0))))))', fmt="num")
        put(ws, f"V{r}", f'=IF(OR(ISNA(T{r}),dsh_ChartBase=0),NA(),IF(R{r}<dsh_ChartBase,NA(),T{r}/INDEX($T${first}:$T${lastr},dsh_ChartBase)*100))',
            fmt="num")
        put(ws, f"W{r}", f'=IF(OR(ISNA(U{r}),dsh_ChartCommon=0),NA(),IF(R{r}<dsh_ChartCommon,NA(),U{r}/INDEX($U${first}:$U${lastr},dsh_ChartCommon)*100))',
            fmt="num")
    put(ws, "T3", f'=IFERROR(MATCH(1,ISNUMBER(T{first}:T{lastr})*ISNUMBER(U{first}:U{lastr}),0),0)', name="dsh_ChartCommon",
        b=b, arr=True, fmt="int")
    put(ws, "U3", f'=IF(dsh_ChartCommon>0,dsh_ChartCommon,IFERROR(MATCH(TRUE,ISNUMBER(T{first}:T{lastr}),0),0))',
        name="dsh_ChartBase", b=b, arr=True, fmt="int")
    ws["S3"] = "count / common start / base start"
    ws["S3"].font = F_HELP
    for col, w in zip("RSTUVW", [4, 11, 11, 11, 10, 10]):
        ws.column_dimensions[col].width = w
    ch = LineChart()
    ch.title = "NAV vs 参照指数（起点=100）"
    ch.height, ch.width = 7.2, 16
    ch.y_axis.title = "Index (=100)"
    ch.x_axis.number_format = "yyyy-mm-dd"
    ch.add_data(Reference(ws, min_col=22, min_row=5, max_row=lastr), titles_from_data=True)
    ch.add_data(Reference(ws, min_col=23, min_row=5, max_row=lastr), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=19, min_row=first, max_row=lastr))
    ws.add_chart(ch, f"B{last + 2}")
    put(ws, f"B{last + 17}", '="参照指数ラベル："&dsh_BmLabel&"（Benchmark (confirmed) 以外はファンドのベンチマークとして扱わない）"',
        kind=None, font=F_NOTE)
    ws.print_area = f"B1:M{last + 18}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True


# =============================================================== 14 Comment
def sheet_comment(b: Book):
    ws = b.ws["14_Comment_Output"]
    intro(ws, "14_Comment_Output — 営業コメント（ドラフト → reviewed_text → 承認 → final_text）",
          "計算セル（Dashboard と同一）と文章テンプレートからドラフトを自動生成。承認済の reviewed_text のみ final_text として使用可能。",
          "13_Dashboard の名前付きセル（dsh_*）／下表 tbl_CommentReview（営業入力）。",
          "① ドラフト確認 → ② comment_id を下表にコピーし reviewed_text に修正文を貼付 → ③ 承認者が review_status=Approved・approver を入力。",
          "データ欠損時は推測で文章を補完せず【データ不足】と表示。市場見通しは記載しない。comment_id に評価日・NAV/Position基準日を含むためデータ更新後は再承認が必要。")
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 22
    for col in "CDEFGHIJ":
        ws.column_dimensions[col].width = 14
    NL = "CHAR(10)"
    STEP = 'dsh_StratType="STEP"'
    HASPOS = "ISNUMBER(dsh_PosDate)"
    HASNAV = "ISNUMBER(dsh_Nav)"
    sections = [
        ("【件名】", "cmt_Subject",
         '=IF(dsh_FundName="—","【データ不足】対象ファンド未選択/未登録",dsh_FundName&"（"&ctl_SelFund&"）定期アップデート（評価日 "&'
         + Dt("ctl_EvalDate") + '&"）")'),
        ("【評価日】", "cmt_EvalLine",
         '="評価日 "&' + Dt("ctl_EvalDate") + '&" ／ NAV基準日 "&' + Dt("dsh_NavDate") + '&" ／ ポジション基準日 "&' + Dt("dsh_PosDate")),
        ("■今週のポイント", "cmt_Points",
         f'=IF(NOT({HASNAV}),"【データ不足】NAVデータが未取得のため記載できません。","・NAV（"&' + Dt("dsh_NavDate")
         + '&"時点）は "&TEXT(dsh_Nav,"#,##0.0000")&"、前週比 "&' + P("dsh_Ret_1W") + '&"、年初来 "&' + P("dsh_Ret_YTD")
         + f'&"（NAVベースの価格リターン）。")&IF(AND({STEP},{HASPOS}),{NL}&"・ポジション（"&' + Dt("dsh_PosDate")
         + '&"時点）：ITMレッグ "&dsh_NITM&"本、最小OTMバッファ "&IF(ISNUMBER(dsh_MinBuf),TEXT(dsh_MinBuf,"0.0%")&"（OTM）","[算出不可]")&"。","")'
         f'&IF(N(dsh_NIssue)>0,{NL}&"・一部レッグ（"&dsh_NIssue&"本）はデータ不足等のため指標を算出していません。","")'),
        ("■市場環境", "cmt_Market",
         '=IF(OR(dsh_BmLabel="Not set",dsh_BmLabel="—"),"参照指数が未設定のため記載していません。","・"&dsh_BmSeries&"：同期間（"&'
         + Dt("dsh_Base_1W") + '&"→"&' + Dt("dsh_NavDate") + '&"）"&' + P("dsh_Bm_1W")
         + '&"、年初来 "&' + P("dsh_Bm_YTD") + '&"（05_Market_Input）。"&IF(dsh_BmLabel="Benchmark (confirmed)","",'
         f'{NL}&"※通貨・ヘッジ条件が未確認の参考指数であり、ファンドのベンチマークではありません。"))&{NL}&"・市場見通しは記載していません。"'),
        ("■ファンドの状況", "cmt_Fund",
         f'=IF(NOT({HASNAV}),"【データ不足】NAVデータが未取得のため記載できません。","・NAVリターン：1D "&' + P("dsh_Ret_1D")
         + '&" / 1W "&' + P("dsh_Ret_1W") + '&" / MTD "&' + P("dsh_Ret_MTD") + '&" / YTD "&' + P("dsh_Ret_YTD")
         + '&" / 1Y "&' + P("dsh_Ret_1Y") + f'&"。"&{NL}&"・高値からの下落率 "&' + P("dsh_CurDD")
         + '&"、最大ドローダウン "&' + P("dsh_MaxDD") + f'&"（価格NAVベース）。"&{NL}&"・AUM（公式）"&IF(ISNUMBER(dsh_AUM),'
         'TEXT(dsh_AUM,"#,##0")&" "&dsh_Ccy,"[データなし]")&IF(ISNUMBER(dsh_EstFlow),"、推定資金流出入（1W）"&TEXT(dsh_EstFlow,"+#,##0;-#,##0")'
         '&"（AUMとNAVからの推定値であり公式フローではありません）","")&"。"&IF(ISNUMBER(dsh_TR_YTD),'
         f'{NL}&"・分配金込みリターン（定義確認済）：YTD "&' + P("dsh_TR_YTD") + ',""))'),
        ("■現在のポジション", "cmt_Position",
         f'=IF(NOT({HASPOS}),"【データ不足】ポジションデータが未取得のため記載できません。",IF(NOT({STEP}),"・ポジション基準日 "&'
         + Dt("dsh_PosDate") + '&"：構成レッグ "&dsh_NLegs&"本（STEP以外の詳細分析は未設定）。","・ポジション基準日 "&' + Dt("dsh_PosDate")
         + '&"：レッグ "&dsh_NLegs&"本（ITM "&dsh_NITM&" / ATM "&dsh_NATM&" / OTM "&dsh_NOTM&"）。"&'
         f'{NL}&"・最小OTMバッファ "&IF(ISNUMBER(dsh_MinBuf),TEXT(dsh_MinBuf,"0.0%")&"（OTM）","[算出不可]")&"。ITMは評価時点のスポットとストライクの位置関係であり、実現損失を意味するものではありません。"'
         f'&IF(N(dsh_NIssue)>0,{NL}&"・算出不可レッグ "&dsh_NIssue&"本（Spot欠損・原資産不一致等。08_Position_Calc 参照）。","")))'),
        ("■前回からの変更点", "cmt_Changes",
         f'=IF(NOT({HASPOS}),"【データ不足】ポジションデータが未取得のため記載できません。",IF(NOT(ISNUMBER(dsh_PrevPos)),'
         '"・前回スナップショットがないため比較していません。","・前回（"&' + Dt("dsh_PrevPos")
         + '&"）比：新規 "&dsh_NNew&"本、継続 "&dsh_NExisting&"本、満期 "&dsh_NExitMat&"本、削除 "&dsh_NExitRem&"本。"'
         f'&IF(N(dsh_NRoll)>0,{NL}&"・ロールの可能性があるレッグ "&dsh_NRoll&"本（運用会社未確認のため、ロールとは断定しません）。","")'
         f'&IF(N(dsh_NCheck)>0,{NL}&"・Position Keyが一意でないレッグ "&dsh_NCheck&"本は比較対象外（要確認）。","")))'),
        ("■今後の注目点", "cmt_Watch",
         f'=IF(NOT({HASPOS}),"・ポジションデータ取得後に記載します。","・"&thr_Mat_Low&"日以内に満期を迎えるレッグ "&dsh_NMatLow&"本"'
         '&IF(ISNUMBER(dsh_NextMat),"（最短満期 "&TEXT(dsh_NextMat,"yyyy/mm/dd")&"）","")&"。")'
         f'&{NL}&"・本欄は保有ポジションの日程等の事実のみを記載し、市場見通しは含みません。"'),
        ("■データに関する注意事項", "cmt_Notes",
         f'=TEXTJOIN({NL},TRUE,"・評価日 "&' + Dt("ctl_EvalDate") + '&" ／ NAV基準日 "&' + Dt("dsh_NavDate") + '&" ／ ポジション基準日 "&'
         + Dt("dsh_PosDate") + '&"。",IF(ctl_SelRow=0,"・対象ファンドが Live Fund List にありません。",IF(INDEX(tbl_Cov[missing_items],ctl_SelRow)="","",'
         '"・未充足／要確認データ："&INDEX(tbl_Cov[missing_items],ctl_SelRow))),'
         '"・リターンはファンドNAV（分配金控除前の価格）ベースであり、Index（指数）リターンとは異なります。",'
         'IF(ISNUMBER(dsh_TR_YTD),"","・分配金込みリターンは nav_with_distribution の定義未確認のため記載していません。"),'
         'IF(dsh_BmLabel="Benchmark (confirmed)","","・参照指数は通貨・ヘッジ条件未確認（またはデータなし）のため比較は参考扱い。"),'
         'IF(ctl_SelRow=0,"",IF(INDEX(tbl_FundMap[multiplier_used],MAX(1,INDEX(tbl_NavCalc[map_row],ctl_SelRow)))="",'
         '"・fund_exposure_multiplier 未確認のため、ファンド固有のエクスポージャーは算出していません。","")),'
         '"・推定資金流出入は公式データではありません。",IF(ctl_ThrConfirmed="Yes","","・アラート閾値は暫定値（元仕様書と未照合）。"))'),
    ]
    r = 7
    for lab, nm, f in sections:
        label(ws, f"B{r}", lab)
        put(ws, f"C{r}", f, name=nm, b=b, wrap=True)
        ws.merge_cells(f"C{r}:J{r}")
        ws.row_dimensions[r].height = 15 if lab.startswith("【") else 62
        r += 1
    r += 1
    joined = "=" + f"&CHAR(10)&CHAR(10)&".join(f'"{lab}"&CHAR(10)&{nm}' for lab, nm, _ in sections)
    label(ws, f"B{r}", "draft_text（全文）")
    put(ws, f"C{r}", joined, name="cmt_Draft", b=b, wrap=True)
    ws.merge_cells(f"C{r}:J{r}")
    ws.row_dimensions[r].height = 120
    r += 2
    CID = ('=ctl_SelInvestor&"|"&ctl_SelFund&"|"&IF(ISNUMBER(ctl_EvalDate),TEXT(ctl_EvalDate,"yyyymmdd"),"NOEVAL")&"|N"&'
           'IF(ISNUMBER(dsh_NavDate),TEXT(dsh_NavDate,"yyyymmdd"),"NA")&"|P"&IF(ISNUMBER(dsh_PosDate),TEXT(dsh_PosDate,"yyyymmdd"),"NA")')
    meta = [
        ("comment_id（下表にコピー）", CID, "cmt_ID", None),
        ("review行", '=IF(COUNTIF(tbl_CommentReview[comment_id],cmt_ID)=0,0,MATCH(cmt_ID,tbl_CommentReview[comment_id],0))', "cmt_RvRow", "int"),
        ("reviewed_text", '=IF(cmt_RvRow=0,"（未登録）",INDEX(tbl_CommentReview[reviewed_text],cmt_RvRow)&"")', "cmt_Reviewed", None),
        ("review_status", '=IF(cmt_RvRow=0,"No review record",INDEX(tbl_CommentReview[review_status],cmt_RvRow)&"")', "cmt_Status", None),
        ("final_text 使用可否",
         '=IF(COUNTIF(tbl_CommentReview[comment_id],cmt_ID)<>1,"No",IF(COUNTIFS(tbl_CommentReview[comment_id],cmt_ID,'
         'tbl_CommentReview[review_status],"Approved",tbl_CommentReview[reviewed_text],"<>",tbl_CommentReview[approver],"<>")=1,"Yes","No"))',
         "cmt_FinalReady", None),
        ("final_text（承認済のみ）", '=IF(cmt_FinalReady="Yes",INDEX(tbl_CommentReview[reviewed_text],cmt_RvRow),"")', "cmt_Final", None),
        ("配信判定（99_Check）", "=ctl_SelGate", "cmt_Gate", None),
    ]
    for lab, f, nm, fmt in meta:
        label(ws, f"B{r}", lab)
        put(ws, f"C{r}", f, name=nm, b=b, wrap=True, fmt=fmt)
        ws.merge_cells(f"C{r}:J{r}")
        if nm in ("cmt_Reviewed", "cmt_Final"):
            ws.row_dimensions[r].height = 90
        r += 1
    add_status_cf(ws, f"C{r - 1}:C{r - 1}")
    ws.print_area = f"B1:J{r - 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    r += 2
    ws.cell(r - 1, 2, "レビュー・承認記録（comment_id 単位。行追加可）").font = F_SECT
    rv = [Col("comment_id", "input", w=40), Col("reviewed_text", "input", w=60), Col("reviewer", "input", w=12),
          Col("review_status", "input", w=11, dv=DV["crev"]), Col("approver", "input", w=12),
          Col("approved_date", "input", fmt="date", w=11), Col("qa_review_status", "input", w=11, dv=DV["crev"]),
          Col("note", "input", w=20),
          Col("dup_count", "calc", '=IF([@comment_id]="","",COUNTIF(§[comment_id],[@comment_id]))', "int", 6)]
    b.tbl("14_Comment_Output", TableSpec("tbl_CommentReview", rv, top=r, left=2, rows=b.rows("commentreview")))
    ws.freeze_panes = None


# =============================================================== 15 QA
def sheet_qa(b: Book):
    ws = b.ws["15_QA_Output"]
    intro(ws, "15_QA_Output — 顧客想定Q&A（ドラフト）",
          "顧客からの想定質問への回答案を、Dashboard と同じ計算セルから生成する。",
          "13_Dashboard（dsh_*）。レビュー状況は 14 tbl_CommentReview[qa_review_status]。",
          "Dashboard で対象を選択すると自動更新。回答案は営業がレビューしてから使用。",
          "要因分析・見通しはデータがないため回答しない（運用会社確認）。ITM≠実現損失、推定フロー≠公式フロー。")
    for col, w in zip("ABCDEF", [5, 30, 70, 26, 16, 14]):
        ws.column_dimensions[col].width = w
    for j, h in enumerate(["No", "想定質問", "回答案（ドラフト）", "根拠データ", "ステータス", "QAレビュー"]):
        c = ws.cell(7, j + 1, h)
        c.font, c.fill, c.border = F_HEAD, FILL["head"], BORDER
    HASNAV = "ISNUMBER(dsh_Nav)"
    HASPOS = "ISNUMBER(dsh_PosDate)"
    qa = [
        ("直近のパフォーマンスは？",
         f'=IF(NOT({HASNAV}),"【データ不足】NAV未取得のため回答できません。","NAV基準日 "&' + Dt("dsh_NavDate") + '&" 時点で、1W "&'
         + P("dsh_Ret_1W") + '&"、MTD "&' + P("dsh_Ret_MTD") + '&"、YTD "&' + P("dsh_Ret_YTD") + '&"（NAVベースの価格リターン、分配金控除前）です。")',
         "07_NAV_Calc ret_*", f'=IF({HASNAV},"OK","DATA REQUIRED")'),
        ("値動きの要因は？", '="本ワークブックには要因分析データが含まれていないため、推測による回答は行いません。運用会社に確認のうえ回答してください。"',
         "なし", '="要運用会社確認"'),
        ("ストライクまでの距離は？",
         f'=IF(dsh_StratType<>"STEP","STEP以外のため該当なし。",IF(NOT({HASPOS}),"【データ不足】ポジション未取得。","ポジション基準日 "&'
         + Dt("dsh_PosDate") + '&" 時点で、ITM "&dsh_NITM&"本、OTMの最小バッファは "&IF(ISNUMBER(dsh_MinBuf),TEXT(dsh_MinBuf,"0.0%"),"[算出不可]")&"です。"))',
         "08_Position_Calc", f'=IF(dsh_StratType<>"STEP","N/A",IF(AND({HASPOS},N(dsh_NIssue)=0),"OK","DATA REQUIRED"))'),
        ("ITMになると損失が確定するのか？",
         '="ITMは評価時点のスポットがストライクを下回る（Put）／上回る（Call）状態を示すもので、実現損失の確定を意味しません。最終的な損益は満期時の水準や商品条件によります。"',
         "定義説明", '="OK"'),
        ("前回からポジションは変わったか？",
         f'=IF(NOT({HASPOS}),"【データ不足】ポジション未取得。",IF(NOT(ISNUMBER(dsh_PrevPos)),"前回スナップショットがないため比較できません。",'
         '"新規 "&dsh_NNew&"本、満期/削除 "&dsh_NExit&"本です。"&IF(N(dsh_NRoll)>0,"ロールの可能性があるレッグがありますが、運用会社未確認のため断定できません。","")))',
         "09_Position_Change", f'=IF(AND({HASPOS},ISNUMBER(dsh_PrevPos)),"OK","DATA REQUIRED")'),
        ("次の満期はいつか？",
         f'=IF(NOT({HASPOS}),"【データ不足】ポジション未取得。",IF(ISNUMBER(dsh_NextMat),"最短の満期は "&TEXT(dsh_NextMat,"yyyy/mm/dd")&'
         '"（基準日から "&dsh_MinDays&"日）です。","満期データがありません。"))', "08_Position_Calc maturity_used",
         f'=IF(AND({HASPOS},ISNUMBER(dsh_NextMat)),"OK","DATA REQUIRED")'),
        ("資金流出入は？",
         '=IF(ISNUMBER(dsh_EstFlow),"AUMとNAVからの推定では1Wで "&TEXT(dsh_EstFlow,"+#,##0;-#,##0")&" です（推定値であり公式フローではありません）。",'
         '"【データ不足】AUMデータが不足しているため回答できません。")', "07 est_flow_1w（推定）",
         '=IF(ISNUMBER(dsh_EstFlow),"OK（推定）","DATA REQUIRED")'),
        ("ベンチマークと比べて？",
         '=IF(dsh_BmLabel="Benchmark (confirmed)",IF(ISNUMBER(dsh_Rel_1W),"同期間の1W差は "&TEXT(dsh_Rel_1W,"+0.00%;-0.00%")&" です。",'
         '"【データ不足】指数データが不足しています。"),"確認済のベンチマークが設定されていないため、比較は行いません（参考指数は通貨・ヘッジ条件未確認）。")',
         "07 bm_ret_*", '=IF(dsh_BmLabel="Benchmark (confirmed)",IF(ISNUMBER(dsh_Rel_1W),"OK","DATA REQUIRED"),"REVIEW REQUIRED")'),
    ]
    for i, (q, a, basis, st) in enumerate(qa, start=1):
        r = 7 + i
        put(ws, f"A{r}", i, kind=None)
        put(ws, f"B{r}", q, kind="calc", wrap=True)
        put(ws, f"C{r}", a, wrap=True)
        put(ws, f"D{r}", basis, wrap=True)
        put(ws, f"E{r}", st)
        put(ws, f"F{r}", '=IF(cmt_RvRow=0,"Unreviewed",IF(INDEX(tbl_CommentReview[qa_review_status],cmt_RvRow)="","Unreviewed",'
                         'INDEX(tbl_CommentReview[qa_review_status],cmt_RvRow)))')
        ws.row_dimensions[r].height = 48
    add_status_cf(ws, "E8:E20")
    ws.print_area = f"A1:F{7 + len(qa)}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# =============================================================== 99 Check
def sheet_check(b: Book):
    ws = b.ws["99_Check"]
    intro(ws, "99_Check — 配信前チェック",
          "必須データ・マッピング・レビュー・承認・宛先を確認し、READY / REVIEW REQUIRED / DATA REQUIRED を判定する。",
          "全シート（数式）。",
          "Status が OK 以外の行の Error Reason・Affected Fund・Required Action を確認し、元シートを修正。",
          "必須条件が未達の場合 READY は表示されません。Global-Gate と Selected の結果が最終判定に反映（Global-Info / Selected-Detail は原因特定用）。")
    TJ = lambda cond, val: f'TEXTJOIN(", ",TRUE,IF({cond},{val},""))'
    SR = "ctl_SelRow"
    checks = [
        # id, category, description, scope, class, result, reason, affected, action, arr_affected
        ("C01", "Control", "評価日が入力されている", "Global-Gate", "DATA",
         '=IF(ISNUMBER(ctl_EvalDate),"PASS","FAIL")', '=IF(ISNUMBER(ctl_EvalDate),"","評価日未入力")', '="(全体)"',
         "00_Control C8 に評価日を入力"),
        ("C02", "Control", "アラート閾値が元仕様書と照合済み", "Global-Gate", "REVIEW",
         '=IF(ctl_ThrConfirmed="Yes","PASS","FAIL")', '=IF(ctl_ThrConfirmed="Yes","","閾値は暫定値のまま")', '="(全体)"',
         "元仕様書の閾値を 00_Control に入力し C12=Yes"),
        ("C06", "Capacity", "02 集約ビュー容量 ≥ holding_key 数", "Global-Gate", "DATA",
         '=IF(N(MAX(tbl_HoldRaw[key_rank]))<=ROWS(tbl_HoldAgg[n]),"PASS","FAIL")',
         '=IF(N(MAX(tbl_HoldRaw[key_rank]))<=ROWS(tbl_HoldAgg[n]),"","holding_key "&MAX(tbl_HoldRaw[key_rank])&" 件 > 容量 "&ROWS(tbl_HoldAgg[n]))',
         '="(全体)"', "tbl_HoldAgg と tbl_DistDetail の行を下方向に拡張（build_workbook.py CAP_HOLD を増やして再生成）"),
        ("C16", "Capacity", "08/09 ポジション容量 ≥ 06 行数", "Global-Gate", "DATA",
         '=IF(ROWS(tbl_PosRaw[strategy_data_key])<=ROWS(tbl_PosCalc[n]),"PASS","FAIL")',
         '=IF(ROWS(tbl_PosRaw[strategy_data_key])<=ROWS(tbl_PosCalc[n]),"","06 行数 "&ROWS(tbl_PosRaw[strategy_data_key])&" > 容量 "&ROWS(tbl_PosCalc[n]))',
         '="(全体)"', "tbl_PosCalc / tbl_PosChg を拡張、または古い履歴を別ブックへ退避"),
        ("C17", "Capacity", "07/09/10/11 ファンド容量 ≥ Live Fund 行数", "Global-Gate", "DATA",
         '=IF(ROWS(tbl_Fund[fund_code])<=ROWS(tbl_NavCalc[n]),"PASS","FAIL")',
         '=IF(ROWS(tbl_Fund[fund_code])<=ROWS(tbl_NavCalc[n]),"","Live Fund "&ROWS(tbl_Fund[fund_code])&" 行 > 容量 "&ROWS(tbl_NavCalc[n]))',
         '="(全体)"', "CAP_FUND を増やして再生成"),
        ("C03", "Live Fund", "fund_code の重複がない", "Global-Info", "DATA",
         '=IF(COUNTIF(tbl_Fund[dup_flag],"DUPLICATE")=0,"PASS","FAIL")',
         '=IF(COUNTIF(tbl_Fund[dup_flag],"DUPLICATE")=0,"","重複行 "&COUNTIF(tbl_Fund[dup_flag],"DUPLICATE")&" 行（自動削除していません）")',
         "=" + TJ('tbl_Fund[dup_flag]="DUPLICATE"', "tbl_Fund[fund_code]"), "Live Fund List の原本を確認し、正しい1行を残すかを担当部署と確認"),
        ("C04", "Holdings", "保有の fund_code が Live Fund List に存在", "Global-Info", "REVIEW",
         '=IF(COUNTIF(tbl_HoldRaw[fund_in_master],"No")=0,"PASS","FAIL")',
         '=IF(COUNTIF(tbl_HoldRaw[fund_in_master],"No")=0,"",COUNTIF(tbl_HoldRaw[fund_in_master],"No")&" 行が Live Fund List にない")',
         "=" + TJ('tbl_HoldRaw[fund_in_master]="No"', "tbl_HoldRaw[fund_code]"), "償還済/対象外ファンドか確認。Live なら 01 に追加"),
        ("C05", "Holdings", "Amount が数値", "Global-Info", "DATA",
         '=IF(COUNTIF(tbl_HoldRaw[amount_check],"Blank")+COUNTIF(tbl_HoldRaw[amount_check],"Not numeric")=0,"PASS","FAIL")',
         '=IF(COUNTIF(tbl_HoldRaw[amount_check],"Blank")+COUNTIF(tbl_HoldRaw[amount_check],"Not numeric")=0,"",'
         '(COUNTIF(tbl_HoldRaw[amount_check],"Blank")+COUNTIF(tbl_HoldRaw[amount_check],"Not numeric"))&" 行の Amount が空欄/非数値")',
         "=" + TJ('(tbl_HoldRaw[amount_check]="Blank")+(tbl_HoldRaw[amount_check]="Not numeric")', "tbl_HoldRaw[holding_key]"),
         "元ファイルの Amount を確認し正しい値で再貼付（0で埋めない）"),
        ("C07", "Distribution", "保有投資家が 12 配信管理に登録済み", "Global-Info", "REVIEW",
         '=IF(COUNTIFS(tbl_HoldRaw[inv_first],1,tbl_HoldRaw[in_dist_list],"No")=0,"PASS","FAIL")',
         '=IF(COUNTIFS(tbl_HoldRaw[inv_first],1,tbl_HoldRaw[in_dist_list],"No")=0,"",COUNTIFS(tbl_HoldRaw[inv_first],1,'
         'tbl_HoldRaw[in_dist_list],"No")&" 投資家が未登録")',
         "=" + TJ('(tbl_HoldRaw[inv_first]=1)*(tbl_HoldRaw[in_dist_list]="No")', "tbl_HoldRaw[end_investor_bdr_id]"),
         "12 tbl_Dist に end_investor_bdr_id を追加し宛先等を入力"),
        ("C08", "Mapping", "全 Live Fund に 03 マッピング行がある", "Global-Info", "DATA",
         '=IF(COUNTIF(tbl_Fund[fund_map_rows],0)=0,"PASS","FAIL")',
         '=IF(COUNTIF(tbl_Fund[fund_map_rows],0)=0,"",COUNTIF(tbl_Fund[fund_map_rows],0)&" ファンドが 03 未登録")',
         "=" + TJ("tbl_Fund[fund_map_rows]=0", "tbl_Fund[fund_code]"), "03 tbl_FundMap に fund_code 行を追加"),
        ("C09", "Mapping", "Mapping Required のファンドがない", "Global-Info", "DATA",
         '=IF(COUNTIF(tbl_ChgSum[mapping_status],"Mapping Required*")=0,"PASS","FAIL")',
         '=IF(COUNTIF(tbl_ChgSum[mapping_status],"Mapping Required*")=0,"",COUNTIF(tbl_ChgSum[mapping_status],"Mapping Required*")&" ファンド")',
         "=" + TJ('LEFT(tbl_ChgSum[mapping_status],16)="Mapping Required"', "tbl_ChgSum[fund_code]"),
         "03 で explicit / Index承認 / 手動キーのいずれかを設定"),
        ("C10", "Mapping", "Review Required のマッピングがない", "Global-Info", "REVIEW",
         '=IF(COUNTIF(tbl_ChgSum[mapping_status],"Review*")=0,"PASS","FAIL")',
         '=IF(COUNTIF(tbl_ChgSum[mapping_status],"Review*")=0,"",COUNTIF(tbl_ChgSum[mapping_status],"Review*")&" ファンド（同一Indexに複数Key等）")',
         "=" + TJ('LEFT(tbl_ChgSum[mapping_status],6)="Review"', "tbl_ChgSum[fund_code]"),
         "03 mapping_status の理由を確認し、explicit key または承認者を設定"),
        ("C22", "Mapping", "Index マッピングに重複行がない", "Global-Info", "REVIEW",
         '=IF(COUNTIF(tbl_IndexMap[dup_count],">1")=0,"PASS","FAIL")', '=IF(COUNTIF(tbl_IndexMap[dup_count],">1")=0,"","重複行あり")',
         "=" + TJ('tbl_IndexMap[dup_count]=2', "tbl_IndexMap[index_name]"), "tbl_IndexMap の重複行を整理"),
        ("C11", "NAV", "NAV の fund_code × nav_date 重複がない", "Global-Info", "DATA",
         '=IF(COUNTIF(tbl_NAV[row_status],"Duplicate date")=0,"PASS","FAIL")',
         '=IF(COUNTIF(tbl_NAV[row_status],"Duplicate date")=0,"",COUNTIF(tbl_NAV[row_status],"Duplicate date")&" 行（重複日は計算から除外中）")',
         "=" + TJ('tbl_NAV[row_status]="Duplicate date"', "tbl_NAV[nav_key]"), "正しいNAVを確認し、誤った行を削除（元ファイルと照合のうえ）"),
        ("C12", "NAV", "NAV の欠損・非数値・0以下がない", "Global-Info", "REVIEW",
         '=IF(COUNTIF(tbl_NAV[row_status],"NAV missing*")+COUNTIF(tbl_NAV[row_status],"NAV<=0")=0,"PASS","FAIL")',
         '=IF(COUNTIF(tbl_NAV[row_status],"NAV missing*")+COUNTIF(tbl_NAV[row_status],"NAV<=0")=0,"","無効NAV行あり（計算から除外中）")',
         "=" + TJ('(tbl_NAV[row_status]="NAV missing / non-numeric")+(tbl_NAV[row_status]="NAV<=0")', "tbl_NAV[nav_key]"),
         "元ファイルを確認"),
        ("C13", "Market", "市場データの重複がない", "Global-Info", "REVIEW",
         '=IF(COUNTIF(tbl_Mkt[dup_count],">1")=0,"PASS","FAIL")', '=IF(COUNTIF(tbl_Mkt[dup_count],">1")=0,"","重複行あり（除外中）")',
         "=" + TJ("tbl_Mkt[valid]=0", "tbl_Mkt[mkt_key]"), "05 の重複行を確認"),
        ("C14", "Position", "旧版ファイル（Superseded）の保持状況", "Global-Info", "INFO",
         '="INFO"', '=COUNTIF(tbl_PosRaw[version_note],"Superseded*")&" 行を旧版として保持（計算不使用）"', '="(全体)"', "対応不要（履歴保持）"),
        ("C15", "Position", "Current スナップショットの Position Key が一意", "Global-Info", "REVIEW",
         '=IF(COUNTIFS(tbl_PosCalc[snap_role],"Current",tbl_PosCalc[uniqueness],"Check*")=0,"PASS","FAIL")',
         '=IF(COUNTIFS(tbl_PosCalc[snap_role],"Current",tbl_PosCalc[uniqueness],"Check*")=0,"",COUNTIFS(tbl_PosCalc[snap_role],"Current",'
         'tbl_PosCalc[uniqueness],"Check*")&" レッグが非一意（同一Ticker・Entry Date）")',
         "=" + TJ('(tbl_PosCalc[snap_role]="Current")*(LEFT(tbl_PosCalc[uniqueness],5)="Check")', "tbl_PosCalc[position_key]"),
         "06 leg_id に追加識別子を入力"),
        ("C18", "Position", "Spot/Strike の原資産不一致（STOP）がない", "Global-Info", "DATA",
         '=IF(COUNTIFS(tbl_PosCalc[snap_role],"Current",tbl_PosCalc[calc_status],"STOP*")=0,"PASS","FAIL")',
         '=IF(COUNTIFS(tbl_PosCalc[snap_role],"Current",tbl_PosCalc[calc_status],"STOP*")=0,"","計算停止レッグあり")',
         "=" + TJ('(tbl_PosCalc[snap_role]="Current")*(LEFT(tbl_PosCalc[calc_status],4)="STOP")', "tbl_PosCalc[position_key]"),
         "原資産・Spot の出所を確認（05 tbl_UndlMap / 06 spot_underlying_raw）"),
        ("C19", "Position", "Current レッグの Spot 欠損がない", "Global-Info", "DATA",
         '=IF(COUNTIFS(tbl_PosCalc[snap_role],"Current",tbl_PosCalc[calc_status],"Spot Missing")+COUNTIFS(tbl_PosCalc[snap_role],"Current",'
         'tbl_PosCalc[calc_status],"Spot Stale")=0,"PASS","FAIL")', '="Spot 欠損/遅延レッグの有無"',
         "=" + TJ('(tbl_PosCalc[snap_role]="Current")*((tbl_PosCalc[calc_status]="Spot Missing")+(tbl_PosCalc[calc_status]="Spot Stale"))',
                  "tbl_PosCalc[position_key]"), "05 に原資産 Spot を追加"),
        ("C20", "Position", "STEP レッグの Ticker 解析失敗がない", "Global-Info", "DATA",
         '=IF(COUNTIFS(tbl_PosCalc[snap_role],"Current",tbl_PosCalc[calc_status],"Parse Failed*")=0,"PASS","FAIL")',
         '="Ticker 解析失敗レッグの有無"',
         "=" + TJ('(tbl_PosCalc[snap_role]="Current")*(LEFT(tbl_PosCalc[calc_status],5)="Parse")', "tbl_PosCalc[position_key]"),
         "06 の put_call_raw / strike_raw / maturity_raw を原本から補記"),
        ("C21", "Position", "Composition と Portfolio の照合", "Global-Info", "REVIEW",
         '=IF(COUNTIFS(tbl_PosCalc[snap_role],"Current",tbl_PosCalc[port_recon],"Not in Portfolio")=0,"PASS","FAIL")',
         '="Portfolio に存在しない Composition レッグの有無"',
         "=" + TJ('(tbl_PosCalc[snap_role]="Current")*(tbl_PosCalc[port_recon]="Not in Portfolio")', "tbl_PosCalc[position_key]"),
         "運用会社に照合"),
        ("C23", "Coverage", "DATA REQUIRED のファンド一覧", "Global-Info", "DATA",
         '=IF(COUNTIF(tbl_Cov[data_gate],"DATA REQUIRED")=0,"PASS","FAIL")',
         '=COUNTIF(tbl_Cov[data_gate],"DATA REQUIRED")&" ファンド"',
         "=" + TJ('tbl_Cov[data_gate]="DATA REQUIRED"', "tbl_Cov[fund_code]"), "11_Update_Coverage の missing_items を確認"),
        ("S01", "Selected", "対象投資家が対象ファンドを保有し、集約が正常", "Selected", "DATA",
         '=IF(ctl_SelHoldRow=0,"FAIL",IF(INDEX(tbl_HoldAgg[agg_status],ctl_SelHoldRow)="OK","PASS","FAIL"))',
         '=IF(ctl_SelHoldRow=0,"02 に holding_key "&ctl_SelHoldKey&" がない",IF(INDEX(tbl_HoldAgg[agg_status],ctl_SelHoldRow)="OK","",'
         'INDEX(tbl_HoldAgg[agg_status],ctl_SelHoldRow)))', "=ctl_SelFund&\"\"", "00_Control の選択、02 の明細を確認"),
        ("S02", "Selected", "対象ファンドのデータゲート（11）", "Selected",
         f'=IF({SR}=0,"DATA",IF(INDEX(tbl_Cov[data_gate],{SR})="REVIEW REQUIRED","REVIEW","DATA"))',
         f'=IF({SR}=0,"FAIL",IF(INDEX(tbl_Cov[data_gate],{SR})="READY","PASS","FAIL"))',
         f'=IF({SR}=0,"Live Fund List にない",IF(INDEX(tbl_Cov[data_gate],{SR})="READY","",INDEX(tbl_Cov[data_gate],{SR})&"："&'
         f'INDEX(tbl_Cov[missing_items],{SR})&IF(INDEX(tbl_Cov[high_unreviewed],{SR})>0," / 未確認Highアラート "&INDEX(tbl_Cov[high_unreviewed],{SR}),"")))',
         "=ctl_SelFund&\"\"", "11_Update_Coverage / 10_Alert で原因を確認し元データを修正・アラートを確認"),
        ("S03", "Selected", "承認済コメント（final_text）がある", "Selected", "REVIEW",
         '=IF(cmt_FinalReady="Yes","PASS","FAIL")', '=IF(cmt_FinalReady="Yes","","review_status="&cmt_Status)', "=ctl_SelFund&\"\"",
         "14 でreviewed_text を登録し承認者が Approved に変更"),
        ("S04", "Selected", "宛先メールアドレスが登録済み", "Selected", "DATA",
         '=IF(ctl_SelDistRow=0,"FAIL",IF(INDEX(tbl_Dist[to_email],ctl_SelDistRow)="","FAIL","PASS"))',
         '=IF(ctl_SelDistRow=0,"12 に投資家未登録",IF(INDEX(tbl_Dist[to_email],ctl_SelDistRow)="","to_email 空欄",""))',
         "=ctl_SelFund&\"\"", "12 tbl_Dist に宛先を入力（不明な場合は空欄のまま＝配信不可）"),
        ("D01", "Selected-Detail", "NAV 状態", "Selected-Detail", "DATA",
         f'=IF({SR}=0,"FAIL",IF(INDEX(tbl_NavCalc[nav_status],{SR})="OK","PASS","FAIL"))',
         f'=IF({SR}=0,"",INDEX(tbl_NavCalc[nav_status],{SR}))', "=ctl_SelFund&\"\"", "04 に最新NAVを追加"),
        ("D02", "Selected-Detail", "マッピング状態", "Selected-Detail", "DATA",
         f'=IF({SR}=0,"FAIL",IF(INDEX(tbl_ChgSum[mapping_status],{SR})="Mapped","PASS","FAIL"))',
         f'=IF({SR}=0,"",INDEX(tbl_ChgSum[mapping_status],{SR}))', "=ctl_SelFund&\"\"", "03 を確認"),
        ("D03", "Selected-Detail", "ポジション状態", "Selected-Detail", "DATA",
         f'=IF({SR}=0,"FAIL",IF(INDEX(tbl_ChgSum[pos_status],{SR})="OK","PASS","FAIL"))',
         f'=IF({SR}=0,"",INDEX(tbl_ChgSum[pos_status],{SR}))', "=ctl_SelFund&\"\"", "06 に最新ポジションを追加"),
        ("D04", "Selected-Detail", "Position 計算不能レッグなし", "Selected-Detail", "DATA",
         f'=IF({SR}=0,"FAIL",IF(N(INDEX(tbl_ChgSum[n_calc_issue],{SR}))=0,"PASS","FAIL"))',
         f'=IF({SR}=0,"",N(INDEX(tbl_ChgSum[n_calc_issue],{SR}))&" レッグ")', "=ctl_SelFund&\"\"", "08 calc_status を確認"),
        ("D05", "Selected-Detail", "High アラートの確認済み", "Selected-Detail", "REVIEW",
         f'=IF({SR}=0,"FAIL",IF(INDEX(tbl_AlertCalc[n_high_unreviewed],{SR})=0,"PASS","FAIL"))',
         f'=IF({SR}=0,"","未確認 High "&INDEX(tbl_AlertCalc[n_high_unreviewed],{SR})&" 件")', "=ctl_SelFund&\"\"",
         "10 tbl_AlertReview に alert_id と確認結果を入力"),
        ("D06", "Selected-Detail", "分配金込みリターン定義（TR）", "Selected-Detail", "INFO",
         '="INFO"', "=dsh_TRStatus", "=ctl_SelFund&\"\"", "03 nav_dist_definition を運用会社に確認"),
        ("D07", "Selected-Detail", "fund_exposure_multiplier 確認", "Selected-Detail", "INFO",
         '="INFO"', f'=IF({SR}=0,"",IF(INDEX(tbl_FundMap[multiplier_used],MAX(1,INDEX(tbl_NavCalc[map_row],{SR})))="","未確認（計算に不使用）","確認済"))',
         "=ctl_SelFund&\"\"", "根拠資料を確認し 03 に入力"),
        ("D08", "Selected-Detail", "参照指数の扱い", "Selected-Detail", "INFO", '="INFO"', "=dsh_BmLabel", "=ctl_SelFund&\"\"",
         "通貨・ヘッジ条件を確認し 03 benchmark_basis_confirmed"),
    ]
    hdr = ["Check ID", "Check Category", "Check Description", "Scope", "Gate Class", "Result", "Status", "Error Reason",
           "Affected Fund", "Required Action", "Owner", "Last Checked"]
    top = 12
    for j, h in enumerate(hdr):
        c = ws.cell(top, j + 1, h)
        c.font, c.fill, c.border = F_HEAD, FILL["head"], BORDER
    for col, w in zip("ABCDEFGHIJKL", [8, 14, 40, 14, 9, 8, 17, 44, 30, 44, 12, 16]):
        ws.column_dimensions[col].width = w
    for i, (cid, cat, desc, scope, cls, res, reason, aff, action) in enumerate(checks):
        r = top + 1 + i
        put(ws, f"A{r}", cid, kind="calc")
        put(ws, f"B{r}", cat, kind="calc")
        put(ws, f"C{r}", desc, kind="calc", wrap=True)
        put(ws, f"D{r}", scope, kind="calc")
        put(ws, f"E{r}", cls, kind="calc")
        put(ws, f"F{r}", res)
        put(ws, f"G{r}", f'=IF(F{r}="PASS","OK",IF(F{r}="INFO","INFO",IF(E{r}="DATA","DATA REQUIRED","REVIEW REQUIRED")))')
        put(ws, f"H{r}", reason, wrap=True)
        put(ws, f"I{r}", aff, wrap=True, arr=aff.startswith("=TEXTJOIN"))
        put(ws, f"J{r}", action, kind="calc", wrap=True)
        put(ws, f"K{r}", '=IF(ctl_Owner="","(未設定)",ctl_Owner)')
        put(ws, f"L{r}", "=NOW()", fmt="dt")
    last = top + len(checks)
    b.name("chk_Scope", "99_Check", f"$D${top + 1}:$D${last}")
    b.name("chk_Status", "99_Check", f"$G${top + 1}:$G${last}")
    label(ws, "A7", "全体ゲート（Global-Gate）")
    ws.merge_cells("A7:B7")
    put(ws, "C7", '=IF(COUNTIFS(chk_Scope,"Global-Gate",chk_Status,"DATA REQUIRED")>0,"DATA REQUIRED",'
                  'IF(COUNTIFS(chk_Scope,"Global-Gate",chk_Status,"REVIEW REQUIRED")>0,"REVIEW REQUIRED","READY"))',
        name="chk_GlobalGate", b=b, font=F_BOLD)
    label(ws, "A8", "最終判定（選択中の投資家×ファンド）")
    ws.merge_cells("A8:B8")
    put(ws, "C8", '=IF(OR(chk_GlobalGate="DATA REQUIRED",COUNTIFS(chk_Scope,"Selected",chk_Status,"DATA REQUIRED")>0),"DATA REQUIRED",'
                  'IF(OR(chk_GlobalGate="REVIEW REQUIRED",COUNTIFS(chk_Scope,"Selected",chk_Status,"REVIEW REQUIRED")>0),'
                  '"REVIEW REQUIRED","READY"))', name="chk_FinalGate", b=b, font=F_BOLD)
    label(ws, "A9", "対象")
    ws.merge_cells("A9:B9")
    put(ws, "C9", '=ctl_SelInvestor&" × "&ctl_SelFund')
    label(ws, "A10", "未解決件数（DATA / REVIEW）")
    ws.merge_cells("A10:B10")
    put(ws, "C10", '=COUNTIF(chk_Status,"DATA REQUIRED")&" / "&COUNTIF(chk_Status,"REVIEW REQUIRED")')
    add_status_cf(ws, f"C7:C8")
    add_status_cf(ws, f"G{top + 1}:G{last}")
    ws.freeze_panes = f"A{top + 1}"
    return last


def build(out: Path, fx=None):
    b = Book(fx)
    sheet_control(b)
    sheet_fund(b)
    sheet_holdings(b)
    sheet_map(b)
    sheet_nav(b)
    sheet_market(b)
    sheet_posraw(b)
    sheet_navcalc(b)
    sheet_poscalc(b)
    sheet_poschg(b)
    sheet_alert(b)
    sheet_cov(b)
    sheet_dist(b)
    sheet_dashboard(b)
    sheet_comment(b)
    sheet_qa(b)
    sheet_check(b)
    tabc = {"00": "1C3F6E", "01": "7F7F7F", "02": "7F7F7F", "03": "BF9000", "04": "7F7F7F", "05": "7F7F7F",
            "06": "7F7F7F", "13": "2E75B6", "14": "2E75B6", "99": "C00000"}
    for name, ws in b.ws.items():
        ws.sheet_properties.tabColor = tabc.get(name[:2], "9DC3E6")
    b.wb.calculation = CalcProperties(fullCalcOnLoad=True)
    b.wb.active = 13
    out.parent.mkdir(parents=True, exist_ok=True)
    b.wb.save(out)
    return out


if __name__ == "__main__":
    if "--fixtures" in sys.argv:
        from fixtures import fixture_data
        p = build(ROOT / "dist" / "QIS_Fund_Monitoring_TESTDATA.xlsx", fixture_data())
    else:
        p = build(ROOT / "dist" / "QIS_Fund_Monitoring.xlsx")
    print("wrote", p)
