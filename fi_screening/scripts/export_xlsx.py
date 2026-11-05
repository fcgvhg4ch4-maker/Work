"""fi_screening/data/*.csv から 12 シートのターゲット分析ブックを生成する。

    python3 -I fi_screening/scripts/export_xlsx.py

- 正本は CSV。ブックは毎回作り直す（手入力は CSV 側に入れてから再生成する）。
- Sales Coverage Input / Sales Action Plan はブック上で直接追記してよい（社内情報。CSV へは戻さない運用も可）。
- 構成比・前期比・パーセンタイル・スコア・順位はすべて Excel 数式（Excel 2010 以降の標準関数のみ）。
"""
import csv
import pathlib

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = pathlib.Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
OUT = ROOT / "dist" / "FI_Target_Screening.xlsx"

UNIVERSE_ASOF = "2026-10-01"
FIN_ASOF = "FY2025（2026年3月期）を最新期とする3期"
P0, P1, P2 = "FY2025(2026/3期)", "FY2024(2025/3期)", "FY2023(2024/3期)"
FD_MAX = 30000  # Financial Data の参照上限行（追記余地）

# ---------------------------------------------------------------- style
JP = "Yu Gothic"
F_TITLE = Font(name=JP, size=14, bold=True, color="1F2933")
F_HEAD = Font(name=JP, size=9, bold=True, color="FFFFFF")
F_SUBH = Font(name=JP, size=9, bold=True, color="1F2933")
F_BODY = Font(name=JP, size=9, color="1F2933")
F_NOTE = Font(name=JP, size=9, color="52606D")
F_WARN = Font(name=JP, size=9, bold=True, color="9B2C2C")
FILL_HEAD = PatternFill("solid", fgColor="1C3F6E")
FILL_CONF = PatternFill("solid", fgColor="7B341E")   # 社内情報シートの見出し色
FILL_SUB = PatternFill("solid", fgColor="E2EAF4")
FILL_INPUT = PatternFill("solid", fgColor="FFF8E1")
FILL_CALC = PatternFill("solid", fgColor="F1F5F9")
FILL_NG = PatternFill("solid", fgColor="FDE2E1")
FILL_NA = PatternFill("solid", fgColor="EDEDED")
FILL_OK = PatternFill("solid", fgColor="E3F4E8")
_t = Side(style="thin", color="D5DAE1")
BORDER = Border(left=_t, right=_t, top=_t, bottom=_t)
WRAP = Alignment(vertical="top", wrap_text=True)
HEAD_AL = Alignment(vertical="center", wrap_text=True)


def read(name):
    p = DATA / name
    if not p.exists():
        return []
    with open(p, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def num(v):
    if v in ("", None):
        return None
    try:
        return float(v)
    except ValueError:
        return v


def title(ws, text, sub, conf=False):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    ws["A2"] = sub
    ws["A2"].font = F_WARN if conf else F_NOTE
    ws.sheet_properties.tabColor = "7B341E" if conf else "1C3F6E"


def header(ws, row, labels, widths=None, conf=False):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_CONF if conf else FILL_HEAD, HEAD_AL, BORDER
    ws.row_dimensions[row].height = 42
    for i, w in enumerate(widths or [], 1):
        ws.column_dimensions[L(i)].width = w
    ws.freeze_panes = ws.cell(row=row + 1, column=3)


def body(ws, r, c, v, fill=None, fmt=None):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font, cell.border = F_BODY, BORDER
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    return cell


def finish(ws, hdr_row, last_row, last_col):
    ws.auto_filter.ref = f"A{hdr_row}:{L(last_col)}{max(last_row, hdr_row + 1)}"


# ---------------------------------------------------------------- data
uni = read("universe.csv")
src = read("sources.csv") + read("auto_sources.csv") + read("manual_sources.csv")
fin = read("manual_financial.csv") + read("financial_data.csv") + read("auto_financial.csv")  # 先勝ち：手入力が優先
pol = read("policy_evidence.csv")
qual = read("qual_inputs.csv")
sel = {r["fi_code"]: r for r in read("selection.csv")}
prof = read("profiles.csv")
batch = read("batch_log.csv")
cat_of = {u["fi_code"]: u["category"] for u in uni}
N = len(uni)

wb = Workbook()
wb.remove(wb.active)

# ================================================================ 1. Institution Universe
ws = wb.create_sheet("1_Institution Universe")
title(ws, "Institution Universe（調査母集団）",
      f"母集団基準日 {UNIVERSE_ASOF}｜主キー=金融機関コード｜信用組合は対象外（追加は行を足すだけで全シートに波及）")
cols = [("fi_code", "金融機関コード", 9), ("legal_name", "正式名称", 22), ("short_name", "略称", 14),
        ("category", "業態", 11), ("segment", "分析セグメント（数式）", 14), ("association", "所属協会", 14),
        ("hq_pref", "本店都道府県", 9), ("region", "地域", 9), ("hq_address", "本店（本部）所在地", 30),
        ("website", "Webサイト", 26), ("corporate_number", "法人番号", 14), ("corporate_number_status", "法人番号ステータス", 13),
        ("holdco", "持株会社", 14), ("holdco_status", "持株会社ステータス", 13), ("reorg_history", "再編履歴", 34),
        ("continuity_break", "時系列の非連続", 18), ("reorg_planned", "再編予定", 18), ("universe_asof", "母集団基準日", 10),
        ("member_source", "名簿出典ID", 15), ("code_source", "コード出典ID", 15), ("research_status", "調査ステータス", 22),
        ("note", "備考", 30)]
HR = 4
header(ws, HR, [c[1] for c in cols], [c[2] for c in cols])
UROW = {}
for i, u in enumerate(uni):
    r = HR + 1 + i
    UROW[u["fi_code"]] = r
    for j, (k, _, _) in enumerate(cols, 1):
        if k == "segment":
            # 信金は FY2025 預金残高（百万円）で規模区分：大≥1兆円, 中≥3,000億円, 小<3,000億円
            dep = f"IFERROR(INDEX('3_Financial Data'!$D$2:$D${FD_MAX},MATCH($A{r}&\"|預金|{P0}\",'3_Financial Data'!$N$2:$N${FD_MAX},0)),\"\")"
            v = (f'=IF(D{r}="信用金庫",IF({dep}="","信金（規模未判定）",IF({dep}>=Dashboard_LargeTh,"信金・大",'
                 f'IF({dep}>=Dashboard_MidTh,"信金・中","信金・小"))),IF(D{r}="信金中央金庫","信金中金","地銀・第二地銀"))')
            body(ws, r, j, v, FILL_CALC)
        else:
            body(ws, r, j, u.get(k, ""), None)
        ws.cell(row=r, column=j).alignment = WRAP if k in ("reorg_history", "note", "hq_address") else Alignment(vertical="top")
ULAST = HR + N
finish(ws, HR, ULAST, len(cols))
U_HR = HR
UR = lambda col: f"'1_Institution Universe'!${col}${U_HR+1}:${col}${ULAST}"  # noqa: E731

# ================================================================ 2. Source Register
ws = wb.create_sheet("2_Source Register")
title(ws, "Source Register（出典台帳）", "Tier1=金融機関一次資料／Tier2=公的機関・業界団体／Tier3=補完。全数値は資料IDでここに紐付く")
scol = [("source_id", "資料ID", 22), ("fi_code", "金融機関コード", 9), ("doc_type", "資料種別", 14), ("title", "資料名", 48),
        ("period", "対象期間", 16), ("publisher", "発行者", 14), ("url", "URL", 50), ("published_date", "公表日", 11),
        ("accessed_date", "確認日", 11), ("tier", "Tier", 6), ("basis", "単体/連結", 8), ("notes", "備考", 40)]
header(ws, 4, [c[1] for c in scol], [c[2] for c in scol])
for i, s in enumerate(src):
    for j, (k, _, _) in enumerate(scol, 1):
        c = body(ws, 5 + i, j, s.get(k, ""))
        if k == "url" and s.get(k, "").startswith("http"):
            c.hyperlink = s[k]
finish(ws, 4, 4 + len(src), len(scol))

# ================================================================ 3. Financial Data (long)
ws = wb.create_sheet("3_Financial Data")
fcol = [("fi_code", "金融機関コード", 9), ("fi_name", "金融機関名", 14), ("metric", "指標名", 26), ("value", "数値", 13),
        ("unit", "単位", 7), ("period", "対象期間", 16), ("basis", "単体/連結", 7), ("source_id", "資料ID", 20),
        ("page", "ページ／所在", 30), ("status", "データステータス", 13), ("accessed_date", "確認日", 10),
        ("formula", "推計式（ESTIMATEDのみ）", 30), ("note", "備考", 40)]
labels = [c[1] for c in fcol] + ["照合キー（数式）", "業態"]
header(ws, 1, labels, [c[2] for c in fcol] + [30, 10])
ws.freeze_panes = "D2"
for i, f in enumerate(fin):
    r = 2 + i
    for j, (k, _, _) in enumerate(fcol, 1):
        v = num(f.get(k)) if k == "value" else f.get(k, "")
        body(ws, r, j, v, fmt="#,##0;[Red]-#,##0" if k == "value" and isinstance(v, float) and abs(v) >= 100 else None)
    body(ws, r, 14, f'=A{r}&"|"&C{r}&"|"&F{r}', FILL_CALC)
    body(ws, r, 15, cat_of.get(f["fi_code"], ""))
FD_LAST = 1 + len(fin)
finish(ws, 1, FD_LAST, 15)
st_rng = f"J2:J{FD_MAX}"
for txt, fill in [("NOT YET CHECKED", FILL_NA), ("NOT DISCLOSED", FILL_SUB), ("CONFLICT", FILL_NG), ("ESTIMATED", FILL_INPUT), ("VERIFIED", FILL_OK)]:
    ws.conditional_formatting.add(st_rng, CellIsRule(operator="equal", formula=[f'"{txt}"'], fill=fill))


def fd(row, metric_cell, period_cell):
    """Financial Data から値を引く（無ければ空文字）。"""
    return (f"IFERROR(INDEX('3_Financial Data'!$D$2:$D${FD_MAX},MATCH($A{row}&\"|\"&{metric_cell}&\"|\"&{period_cell},"
            f"'3_Financial Data'!$N$2:$N${FD_MAX},0)),\"\")")


def fd_status(row, metric, period):
    return (f"IFERROR(INDEX('3_Financial Data'!$J$2:$J${FD_MAX},MATCH($A{row}&\"|{metric}|{period}\","
            f"'3_Financial Data'!$N$2:$N${FD_MAX},0)),\"NOT YET CHECKED\")")


def ident_cols(ws, r, code):
    ur = UROW[code]
    body(ws, r, 1, code)
    body(ws, r, 2, f"='1_Institution Universe'!C{ur}")
    body(ws, r, 3, f"='1_Institution Universe'!D{ur}")
    body(ws, r, 4, f"='1_Institution Universe'!E{ur}", FILL_CALC)
    body(ws, r, 5, f"='1_Institution Universe'!H{ur}")


IDENT = [("金融機関コード", 9), ("略称", 14), ("業態", 10), ("分析セグメント", 13), ("地域", 9)]

# ================================================================ 4. Securities Portfolio (wide)
ws = wb.create_sheet("4_Securities Portfolio")
title(ws, "Securities Portfolio（有価証券ポートフォリオ・単体）",
      "値は 3_Financial Data から数式で参照（行3=指標名、行4=期）。単位: 百万円。比率・前期比は数式。空欄=未取得（ステータスは3_Financial Dataで確認）")
raw_cols = []  # (metric, period)
for m in ["総資産", "預金", "譲渡性預金", "貸出金", "有価証券", "国債", "地方債", "短期社債", "社債", "株式",
          "その他の証券（推計・外国証券等を含む上位集合）", "外国証券", "その他有価証券評価差額", "純資産", "国債等債券売却損",
          "自己資本比率"]:
    raw_cols.append((m, P0))
for m in ["総資産", "預金", "貸出金", "有価証券", "外国証券", "その他の証券（推計・外国証券等を含む上位集合）", "国債等債券売却損"]:
    raw_cols.append((m, P1))
for m in ["有価証券", "外国証券", "その他の証券（推計・外国証券等を含む上位集合）"]:
    raw_cols.append((m, P2))
SP = {}  # (metric, period) -> column letter
c0 = len(IDENT) + 1
for k, (m, p) in enumerate(raw_cols):
    SP[(m, p)] = L(c0 + k)
calc = [  # (label, formula template using {r} and SP letters)
    ("有価証券／総資産", "=IFERROR({有価証券}{r}/{総資産}{r},\"\")"),
    ("預貸率（貸出金／(預金+NCD)）", "=IFERROR({貸出金}{r}/({預金}{r}+{譲渡性預金}{r}),\"\")"),
    ("有価証券／預金", "=IFERROR({有価証券}{r}/{預金}{r},\"\")"),
    ("国債構成比", "=IFERROR({国債}{r}/{有価証券}{r},\"\")"),
    ("地方債構成比", "=IFERROR({地方債}{r}/{有価証券}{r},\"\")"),
    ("社債構成比", "=IFERROR(({社債}{r}+{短期社債}{r})/{有価証券}{r},\"\")"),
    ("株式構成比", "=IFERROR({株式}{r}/{有価証券}{r},\"\")"),
    ("その他の証券（推計）構成比", "=IFERROR({その他の証券（推計・外国証券等を含む上位集合）}{r}/{有価証券}{r},\"\")"),
    ("外国証券／有価証券", "=IF({外国証券}{r}=\"\",\"\",IFERROR({外国証券}{r}/{有価証券}{r},\"\"))"),
    ("評価差額／純資産", "=IFERROR({その他有価証券評価差額}{r}/{純資産}{r},\"\")"),
    ("有価証券 前期比", "=IFERROR({有価証券}{r}/{有価証券@P1}{r}-1,\"\")"),
    ("外国証券 前期比", "=IF(OR({外国証券}{r}=\"\",{外国証券@P1}{r}=\"\"),\"\",IFERROR({外国証券}{r}/{外国証券@P1}{r}-1,\"\"))"),
    ("外国証券 2期累計変化", "=IF(OR({外国証券}{r}=\"\",{外国証券@P2}{r}=\"\"),\"\",IFERROR({外国証券}{r}/{外国証券@P2}{r}-1,\"\"))"),
    ("その他の証券（推計）前期比", "=IFERROR({その他の証券（推計・外国証券等を含む上位集合）}{r}/{その他の証券（推計・外国証券等を含む上位集合）@P1}{r}-1,\"\")"),
    ("外国証券ステータス(FY2025)", None),
    ("時系列の非連続", None),
]


def fill_tpl(t, r):
    out = t
    for (m, p), col in sorted(SP.items(), key=lambda x: -len(x[0][0])):
        tag = "{" + m + ("" if p == P0 else ("@P1" if p == P1 else "@P2")) + "}"
        out = out.replace(tag, col)
    return out.replace("{r}", str(r))


hdr = [h for h, _ in IDENT] + [f"{m}" for m, p in raw_cols] + [c[0] for c in calc]
HR = 5
header(ws, HR, hdr, [w for _, w in IDENT] + [12] * len(raw_cols) + [11] * len(calc))
for k, (m, p) in enumerate(raw_cols):
    col = c0 + k
    ws.cell(row=3, column=col, value=m).font = F_SUBH
    ws.cell(row=4, column=col, value=p).font = F_NOTE
    ws.cell(row=HR, column=col, value=f"{m}\n{p[:6]}")
SP_ROW = {}
for i, u in enumerate(uni):
    r = HR + 1 + i
    SP_ROW[u["fi_code"]] = r
    ident_cols(ws, r, u["fi_code"])
    for k in range(len(raw_cols)):
        col = c0 + k
        body(ws, r, col, "=" + fd(r, f"{L(col)}$3", f"{L(col)}$4"), fmt="#,##0;[Red]-#,##0")
    for j, (lab, tpl) in enumerate(calc):
        col = c0 + len(raw_cols) + j
        if lab.startswith("外国証券ステータス"):
            body(ws, r, col, "=" + fd_status(r, "外国証券", P0), FILL_CALC)
        elif lab == "時系列の非連続":
            body(ws, r, col, f"='1_Institution Universe'!P{UROW[u['fi_code']]}&\"\"", FILL_CALC)
        else:
            body(ws, r, col, fill_tpl(tpl, r), FILL_CALC, "0.0%")
SP_LAST = HR + N
CALC_COL = {lab: L(c0 + len(raw_cols) + j) for j, (lab, _) in enumerate(calc)}
finish(ws, HR, SP_LAST, len(hdr))
ws.freeze_panes = "C6"


SP_HR = HR


def spref(colletter):
    return f"'4_Securities Portfolio'!${colletter}${SP_HR+1}:${colletter}${SP_LAST}"


def spcell(code, colletter):
    return f"'4_Securities Portfolio'!{colletter}{SP_ROW[code]}"


# ================================================================ 5. Foreign Bond Analysis
ws = wb.create_sheet("5_Foreign Bond Analysis")
title(ws, "Foreign Bond Analysis（外債関連）",
      "外国証券≠外国債券≠米国債。外国証券の減少要因（売却/償還/為替/時価）は資料に明記がある場合のみ記載。Type判定は数式（複数該当可）")
fb_metrics = [("外国証券", P0), ("外国証券", P1), ("外国証券", P2), ("外国債券", P0), ("米国債", P0),
              ("外貨建債券", P0), ("為替ヘッジ比率", P0), ("ヘッジ付外債", P0), ("ヘッジなし外債", P0), ("外債評価損益", P0)]
fb_text = ["通貨構成", "ヘッジコストに関する開示", "外債の売却・償還動向（開示記述）", "外債の再投資方針（開示記述）"]
pct_cols = ["外国証券 業態内pct", "外国証券比率 業態内pct", "有価証券 業態内pct"]
types = ["Type A 外債規模大", "Type B 外債縮小", "Type C 外債積極化", "Type D 規模大・外債比率低", "Type E ファンド活用"]
hdr = [h for h, _ in IDENT] + [f"{m}\n{p[:6]}" for m, p in fb_metrics] + ["外国証券ステータス"] + fb_text + pct_cols + types
HR = 5
header(ws, HR, hdr, [w for _, w in IDENT] + [12] * len(fb_metrics) + [13] + [26] * len(fb_text) + [10] * 3 + [11] * 5)
c0 = len(IDENT) + 1
for k, (m, p) in enumerate(fb_metrics):
    ws.cell(row=3, column=c0 + k, value=m).font = F_SUBH
    ws.cell(row=4, column=c0 + k, value=p).font = F_NOTE
qmap = {}
for q in qual:
    qmap.setdefault(q["fi_code"], {})[q["item"]] = q
FB_ROW = {}
cA = {}  # header label -> letter
for j, h in enumerate(hdr, 1):
    cA[h] = L(j)
fx = spref(SP["外国証券", P0])
for i, u in enumerate(uni):
    r = HR + 1 + i
    code = u["fi_code"]
    FB_ROW[code] = r
    ident_cols(ws, r, code)
    for k in range(len(fb_metrics)):
        col = c0 + k
        body(ws, r, col, "=" + fd(r, f"{L(col)}$3", f"{L(col)}$4"), fmt="#,##0;[Red]-#,##0")
    col = c0 + len(fb_metrics)
    body(ws, r, col, "=" + fd_status(r, "外国証券", P0), FILL_CALC)
    for t in fb_text:
        col += 1
        body(ws, r, col, qmap.get(code, {}).get(t, {}).get("text", ""), FILL_INPUT).alignment = WRAP
    seg = f"$D{r}"
    segR = f"$D${HR+1}:$D${HR+N}"
    for lab, metric_col in zip(pct_cols, [SP["外国証券", P0], CALC_COL["外国証券／有価証券"], SP["有価証券", P0]]):
        col += 1
        v = spcell(code, metric_col)
        rng = spref(metric_col)
        body(ws, r, col, f'=IF(ISNUMBER({v}),IF(COUNTIFS({segR},{seg},{rng},">-1E+300")<2,"",'
                         f'COUNTIFS({segR},{seg},{rng},"<"&{v})/(COUNTIFS({segR},{seg},{rng},">-1E+300")-1)),"")', FILL_CALC, "0%")
    pfx, pratio, psec = cA[pct_cols[0]], cA[pct_cols[1]], cA[pct_cols[2]]
    chg2 = spcell(code, CALC_COL["外国証券 2期累計変化"])
    ratio_pct = f"{pratio}{r}"
    tA = f'=IF(AND({pfx}{r}="",{pratio}{r}=""),"保留",IF(OR(AND(ISNUMBER({pfx}{r}),{pfx}{r}>=0.8),AND(ISNUMBER({pratio}{r}),{pratio}{r}>=0.8)),"該当",""))'
    tB = f'=IF({chg2}="","保留",IF({chg2}<=-Dashboard_TypeChg,"該当",""))'
    tC = f'=IF({chg2}="","保留",IF({chg2}>=Dashboard_TypeChg,"該当",""))&IF(ISNUMBER(SEARCH("拡大",{cA["外債の再投資方針（開示記述）"]}{r}))," （方針:拡大）","")'
    tD = f'=IF(OR({psec}{r}="",{ratio_pct}=""),"保留",IF(AND({psec}{r}>=0.7,{ratio_pct}<=0.4),"該当",""))'
    tE = "__TYPE_E__"
    for v in (tA, tB, tC, tD, tE):
        col += 1
        body(ws, r, col, v, FILL_CALC)
FB_LAST = HR + N
finish(ws, HR, FB_LAST, len(hdr))
ws.freeze_panes = "C6"
FB_TYPEE_COL = len(hdr)
for t in types:
    ws.conditional_formatting.add(f"{cA[t]}{HR+1}:{cA[t]}{FB_LAST}", CellIsRule(operator="equal", formula=['"該当"'], fill=FILL_OK))

# ================================================================ 6. Investment Policy Evidence
ws = wb.create_sheet("6_Investment Policy Evidence")
title(ws, "Investment Policy Evidence（運用方針の開示記述）",
      "原文抜粋は資料の記載どおり。要約・方向性は分析者の判断（推論）。スコアB群・A3/A4・D群の根拠IDはここを参照")
pcol = [("evidence_id", "証拠ID", 12), ("fi_code", "金融機関コード", 9), ("fi_name", "金融機関名", 14), ("category", "カテゴリ", 16),
        ("quote", "原文抜粋（事実）", 60), ("summary", "要約（分析者）", 36), ("direction", "方向性", 9),
        ("source_id", "資料ID", 20), ("page", "ページ", 8), ("published_date", "公表日", 10), ("accessed_date", "確認日", 10)]
header(ws, 4, [c[1] for c in pcol], [c[2] for c in pcol])
for i, p in enumerate(pol):
    for j, (k, _, _) in enumerate(pcol, 1):
        body(ws, 5 + i, j, p.get(k, "")).alignment = WRAP
dv = DataValidation(type="list", formula1='"外債方針,米国金利見通し,為替ヘッジ方針,ポートフォリオ再構築,収益性向上,債券リスク管理,投信・ファンド活用,外部運用"', allow_blank=True)
ws.add_data_validation(dv)
dv.add(f"D5:D{5+max(len(pol),0)+500}")
finish(ws, 4, 4 + max(len(pol), 1), len(pcol))

# ================================================================ 7. Target Scoring
ws = wb.create_sheet("7_Target Scoring")
title(ws, "Target Scoring（スコア算出）",
      "黄=手入力（定性評価：満点以内の点数。空欄=評価保留（未調査）、\"ND\"=評価保留（非開示））／灰=数式。評価保留は0点にせず分母から除外。配点は行3（Dashboardのパラメータ）")
items = [  # code, label, max name, kind, source
    ("A1", "A1 外国証券規模", "W_A1", "pct", pct_cols[0]),
    ("A2", "A2 外国証券比率", "W_A2", "pct", pct_cols[1]),
    ("A3", "A3 外債投資方針の開示", "W_A3", "input", None),
    ("A4", "A4 為替ヘッジ利用", "W_A4", "input", None),
    ("B1", "B1 外債運用見直しの開示", "W_B1", "input", None),
    ("B2", "B2 再投資ニーズ示唆", "W_B2", "input", None),
    ("B3", "B3 ヘッジコスト問題意識", "W_B3", "input", None),
    ("B4", "B4 金利・為替リスク（評価損/純資産・逆順位）", "W_B4", "pct_rev", "評価差額／純資産"),
    ("C1", "C1 有価証券運用規模", "W_C1", "pct", pct_cols[2]),
    ("C2", "C2 自己資本比率", "W_C2", "pct", "自己資本比率"),
    ("C3", "C3 運用余力（1−預貸率）", "W_C3", "pct_rev", "預貸率（貸出金／(預金+NCD)）"),
    ("D1", "D1 投信・ファンド利用", "W_D1", "input", None),
    ("D2", "D2 外部運用会社の活用", "W_D2", "input", None),
    ("D3", "D3 複雑商品の投資経験", "W_D3", "input", None),
]
hdr = [h for h, _ in IDENT]
for code, lab, w, kind, _ in items:
    if kind == "input":
        hdr += [f"{lab}\n入力", f"{code} 得点", f"{code} 根拠ID"]
    else:
        hdr += [f"{lab}\npct", f"{code} 得点"]
hdr += ["確認済み得点", "確認済み満点", "参考スコア(100換算)", "データ充足率", "本ランキング対象", "全国順位", "業態内順位", "A群", "B群", "C群", "D群"]
HR = 5
header(ws, HR, hdr, [w for _, w in IDENT] + [9] * (len(hdr) - len(IDENT)))
TS = {h: L(j) for j, h in enumerate(hdr, 1)}
for code, lab, w, kind, _ in items:  # 行3に満点（パラメータ参照）
    sc = TS[f"{code} 得点"]
    ws[f"{sc}3"] = f"={w}"
    ws[f"{sc}3"].font = F_SUBH
    ws[f"{sc}4"] = "満点"
    ws[f"{sc}4"].font = F_NOTE
segR = f"$D${HR+1}:$D${HR+N}"
TS_ROW = {}
score_cols = []
for i, u in enumerate(uni):
    r = HR + 1 + i
    code_ = u["fi_code"]
    TS_ROW[code_] = r
    ident_cols(ws, r, code_)
    qv = qmap.get(code_, {})
    for code, lab, w, kind, srcname in items:
        sc = TS[f"{code} 得点"]
        if kind == "input":
            ic = TS[f"{lab}\n入力"]
            q = qv.get(code, {})
            g = q.get("grade", "")
            body(ws, r, ws[f"{ic}1"].column, num(g) if g not in ("", "ND") else (g or None), FILL_INPUT)
            body(ws, r, ws[f"{sc}1"].column, f'=IF(ISNUMBER({ic}{r}),MIN({ic}{r},{sc}$3),"")', FILL_CALC, "0.0")
            body(ws, r, ws[f"{TS[f'{code} 根拠ID']}1"].column, q.get("evidence_ids", ""), FILL_INPUT)
        else:
            pc = TS[f"{lab}\npct"]
            if kind == "pct" and srcname in pct_cols:
                f = f"='5_Foreign Bond Analysis'!{cA[srcname]}{FB_ROW[code_]}"
            else:
                colL = SP[(srcname, P0)] if srcname == "自己資本比率" else CALC_COL[srcname]
                v, rng = spcell(code_, colL), spref(colL)
                cmp_ = '">"&' if kind == "pct_rev" else '"<"&'
                f = (f'=IF(ISNUMBER({v}),IF(COUNTIFS({segR},$D{r},{rng},">-1E+300")<2,"",'
                     f'COUNTIFS({segR},$D{r},{rng},{cmp_}{v})/(COUNTIFS({segR},$D{r},{rng},">-1E+300")-1)),"")')
            body(ws, r, ws[f"{pc}1"].column, f, FILL_CALC, "0%")
            body(ws, r, ws[f"{sc}1"].column,
                 f'=IF(ISNUMBER({pc}{r}),{sc}$3*IF({pc}{r}>=0.8,1,IF({pc}{r}>=0.6,0.7,IF({pc}{r}>=0.4,0.5,IF({pc}{r}>=0.2,0.3,0.1)))),"")',
                 FILL_CALC, "0.0")
    scs = [TS[f"{c} 得点"] for c, *_ in items]
    got = "+".join(f"N({c}{r})" for c in scs)
    mx = "+".join(f"IF(ISNUMBER({c}{r}),{c}$3,0)" for c in scs)
    body(ws, r, ws[f"{TS['確認済み得点']}1"].column, f"={got}", FILL_CALC, "0.0")
    body(ws, r, ws[f"{TS['確認済み満点']}1"].column, f"={mx}", FILL_CALC, "0.0")
    body(ws, r, ws[f"{TS['参考スコア(100換算)']}1"].column,
         f'=IF({TS["確認済み満点"]}{r}=0,"",{TS["確認済み得点"]}{r}/{TS["確認済み満点"]}{r}*100)', FILL_CALC, "0.0")
    body(ws, r, ws[f"{TS['データ充足率']}1"].column,
         f'={TS["確認済み満点"]}{r}/SUM(${scs[0]}$3:${scs[-1]}$3)'.replace(f"SUM(${scs[0]}$3:${scs[-1]}$3)", "(" + "+".join(f"{c}$3" for c in scs) + ")"),
         FILL_CALC, "0%")
    body(ws, r, ws[f"{TS['本ランキング対象']}1"].column,
         f'=IF(AND($C{r}<>"信金中央金庫",{TS["データ充足率"]}{r}>=Dashboard_MinFill,ISNUMBER({TS["参考スコア(100換算)"]}{r})),1,0)', FILL_CALC)
    for grp, codes in [("A群", ["A1", "A2", "A3", "A4"]), ("B群", ["B1", "B2", "B3", "B4"]), ("C群", ["C1", "C2", "C3"]), ("D群", ["D1", "D2", "D3"])]:
        body(ws, r, ws[f"{TS[grp]}1"].column, "=" + "+".join(f"N({TS[c+' 得点']}{r})" for c in codes), FILL_CALC, "0.0")
TS_LAST = HR + N
el, scv = TS["本ランキング対象"], TS["参考スコア(100換算)"]
for i in range(N):
    r = HR + 1 + i
    body(ws, r, ws[f"{TS['全国順位']}1"].column,
         f'=IF({el}{r}=1,COUNTIFS(${el}${HR+1}:${el}${TS_LAST},1,${scv}${HR+1}:${scv}${TS_LAST},">"&{scv}{r})+1,"")', FILL_CALC)
    body(ws, r, ws[f"{TS['業態内順位']}1"].column,
         f'=IF({el}{r}=1,COUNTIFS(${el}${HR+1}:${el}${TS_LAST},1,$D${HR+1}:$D${TS_LAST},$D{r},${scv}${HR+1}:${scv}${TS_LAST},">"&{scv}{r})+1,"")', FILL_CALC)
finish(ws, HR, TS_LAST, len(hdr))
ws.freeze_panes = "C6"

# Type E（ファンド活用）= D1 または D2 が加点
wsfb = wb["5_Foreign Bond Analysis"]
for code_, r in FB_ROW.items():
    tr = TS_ROW[code_]
    d1, d2 = TS["D1 得点"], TS["D2 得点"]
    wsfb.cell(row=r, column=FB_TYPEE_COL,
              value=f"=IF(AND('7_Target Scoring'!{d1}{tr}=\"\",'7_Target Scoring'!{d2}{tr}=\"\"),\"保留\",IF(N('7_Target Scoring'!{d1}{tr})+N('7_Target Scoring'!{d2}{tr})>0,\"該当\",\"\"))")

# ================================================================ 8. Priority Ranking
ws = wb.create_sheet("8_Priority Ranking")
title(ws, "Priority Ranking（優先順位）",
      "順位は 7_Target Scoring の数式結果。充足率が閾値未満の先は『データ不足候補』として順位を付けない。選定段階・理由は分析者入力（data/selection.csv）")
hdr = [h for h, _ in IDENT] + ["参考スコア", "データ充足率", "本ランキング対象", "全国順位", "業態内順位", "Type A", "Type B", "Type C", "Type D", "Type E",
                               "選定段階（30/15/10）", "選定理由", "除外・保留理由"]
HR = 5
header(ws, HR, hdr, [w for _, w in IDENT] + [10, 10, 9, 8, 8, 8, 8, 12, 8, 8, 12, 40, 40])
for i, u in enumerate(uni):
    r = HR + 1 + i
    code_ = u["fi_code"]
    ident_cols(ws, r, code_)
    tr, fr = TS_ROW[code_], FB_ROW[code_]
    vals = [f"='7_Target Scoring'!{TS['参考スコア(100換算)']}{tr}", f"='7_Target Scoring'!{TS['データ充足率']}{tr}",
            f"='7_Target Scoring'!{TS['本ランキング対象']}{tr}", f"='7_Target Scoring'!{TS['全国順位']}{tr}",
            f"='7_Target Scoring'!{TS['業態内順位']}{tr}"] + [f"='5_Foreign Bond Analysis'!{cA[t]}{fr}" for t in types]
    fmts = ["0.0", "0%", None, None, None] + [None] * 5
    for j, (v, fm) in enumerate(zip(vals, fmts)):
        body(ws, r, 6 + j, v, FILL_CALC, fm)
    s = sel.get(code_, {})
    body(ws, r, 16, s.get("stage", ""), FILL_INPUT)
    body(ws, r, 17, s.get("reason", ""), FILL_INPUT).alignment = WRAP
    body(ws, r, 18, s.get("exclusion", ""), FILL_INPUT).alignment = WRAP
finish(ws, HR, HR + N, len(hdr))
ws.freeze_panes = "C6"

# ================================================================ 9. Bank Profiles
ws = wb.create_sheet("9_Bank Profiles")
title(ws, "Bank Profiles（上位候補の提案仮説）",
      "区分: 事実（出典付）／推論（根拠と確信度）／営業仮説。事実以外は顧客確認前の仮説であり購入意欲を示すものではない")
pcol = [("fi_code", "金融機関コード", 9), ("fi_name", "金融機関名", 14), ("section", "項目", 22), ("kind", "区分", 8),
        ("content", "内容", 80), ("confidence", "確信度", 7), ("evidence_ids", "根拠ID", 18), ("updated", "更新日", 10)]
header(ws, 4, [c[1] for c in pcol], [c[2] for c in pcol])
for i, p in enumerate(prof):
    for j, (k, _, _) in enumerate(pcol, 1):
        c = body(ws, 5 + i, j, p.get(k, ""))
        c.alignment = WRAP
        if k == "kind":
            c.fill = {"事実": FILL_OK, "推論": FILL_SUB, "営業仮説": FILL_INPUT}.get(p.get(k), None) or PatternFill()
finish(ws, 4, 4 + max(len(prof), 1), len(pcol))

# ================================================================ 10. Sales Coverage Input (社内)
ws = wb.create_sheet("10_Sales Coverage Input")
title(ws, "Sales Coverage Input（社内情報・営業入力）",
      "【社内限・Confidential】公開情報シートとは別管理。M1月次ファンドフローは個別金融機関の売買が特定できない限り個別先の根拠に使わない", conf=True)
hdr = [h for h, _ in IDENT] + ["営業担当者", "既存カバレッジ状況", "既存ファンド購入実績", "過去の商談履歴", "ファンド投資の可否",
                               "現在の外債投資ニーズ", "次回面談予定", "営業担当者の評価", "情報源（社内）", "機密区分", "更新日", "更新者"]
header(ws, 4, hdr, [w for _, w in IDENT] + [12, 12, 14, 30, 12, 24, 11, 10, 14, 12, 10, 10], conf=True)
for i, u in enumerate(uni):
    r = 5 + i
    ident_cols(ws, r, u["fi_code"])
    for j in range(6, len(hdr) + 1):
        body(ws, r, j, "社内限" if hdr[j - 1] == "機密区分" else None, FILL_INPUT)
vals = {"既存カバレッジ状況": '"カバー中,未カバー,過去カバー,他社メイン"', "既存ファンド購入実績": '"あり（本商品）,あり（他ファンド）,なし,不明"',
        "ファンド投資の可否": '"可,条件付き可,不可,不明"', "営業担当者の評価": '"A,B,C,D"', "機密区分": '"社内限,厳秘"'}
for lab, lst in vals.items():
    j = hdr.index(lab) + 1
    d = DataValidation(type="list", formula1=lst, allow_blank=True)
    ws.add_data_validation(d)
    d.add(f"{L(j)}5:{L(j)}{4+N}")
d = DataValidation(type="date", operator="greaterThan", formula1="DATE(2020,1,1)", allow_blank=True)
ws.add_data_validation(d)
d.add(f"{L(hdr.index('次回面談予定')+1)}5:{L(hdr.index('次回面談予定')+1)}{4+N}")
finish(ws, 4, 4 + N, len(hdr))
ws.freeze_panes = "C5"
# M1 フローデータの受け皿
m1 = wb.create_sheet("10b_M1 Fund Flow")
title(m1, "M1 月次ファンドフロー（受領後に貼付）",
      "【社内限】先にデータ定義（対象範囲・粒度・個別金融機関識別の可否）を記入。識別不可のデータは業態・地域合計としてのみ利用", conf=True)
for i, (k, v) in enumerate([("データ提供元", "M1"), ("対象範囲", ""), ("粒度（月次/業態/地域/個別先）", ""),
                            ("個別金融機関の識別可否", "未確認"), ("受領日", ""), ("利用制限", "個別先の投資行動の根拠に使用しない（識別可能な場合を除く）")]):
    body(m1, 4 + i, 1, k, FILL_SUB)
    body(m1, 4 + i, 2, v, FILL_INPUT)
m1.column_dimensions["A"].width = 30
m1.column_dimensions["B"].width = 60
header(m1, 11, ["年月", "集計単位", "金融機関コード（識別可能な場合のみ）", "ファンド名", "購入額", "解約額", "純流入", "単位", "備考"], [10, 12, 18, 24, 12, 12, 12, 8, 30], conf=True)
m1.freeze_panes = "A12"

# ================================================================ 11. Sales Action Plan
ws = wb.create_sheet("11_Sales Action Plan")
title(ws, "Sales Action Plan（営業アクション）", "【社内限】営業担当者が追記。仮説IDは 9_Bank Profiles の項目に対応", conf=True)
hdr = ["金融機関コード", "金融機関名（数式）", "アクション", "関連仮説／確認質問", "担当", "期日", "ステータス", "結果・次アクション", "更新日"]
header(ws, 4, hdr, [9, 14, 30, 36, 10, 11, 10, 36, 10], conf=True)
for r in range(5, 205):
    body(ws, r, 1, None, FILL_INPUT)
    body(ws, r, 2, f"=IF(A{r}=\"\",\"\",IFERROR(INDEX({UR('C')},MATCH(A{r}&\"\",{UR('A')},0)),\"コード不一致\"))", FILL_CALC)
    for j in range(3, 10):
        body(ws, r, j, None, FILL_INPUT)
d = DataValidation(type="list", formula1='"未着手,準備中,実施済,保留,中止"', allow_blank=True)
ws.add_data_validation(d)
d.add("G5:G204")
finish(ws, 4, 204, len(hdr))

# ================================================================ 12. Data Quality Dashboard
ws = wb.create_sheet("12_Data Quality Dashboard")
title(ws, "Data Quality Dashboard（データ品質・パラメータ）",
      f"母集団基準日 {UNIVERSE_ASOF}｜財務基準 {FIN_ASOF}｜件数・比率はすべて数式")
ws.column_dimensions["A"].width = 34
for c in "BCDEFGHIJKL":
    ws.column_dimensions[c].width = 13
r = 4
body(ws, r, 1, "パラメータ（変更すると全シート再計算）", FILL_SUB)
params = [("Dashboard_LargeTh", "信金・大 の預金残高下限（百万円）", 1_000_000), ("Dashboard_MidTh", "信金・中 の預金残高下限（百万円）", 300_000),
          ("Dashboard_MinFill", "本ランキング掲載の最低データ充足率", 0.6), ("Dashboard_TypeChg", "Type B/C 判定：外国証券2期累計変化の閾値", 0.2)]
weights = [("W_A1", 10), ("W_A2", 10), ("W_A3", 8), ("W_A4", 7), ("W_B1", 10), ("W_B2", 8), ("W_B3", 7), ("W_B4", 5),
           ("W_C1", 8), ("W_C2", 6), ("W_C3", 6), ("W_D1", 7), ("W_D2", 4), ("W_D3", 4)]
from openpyxl.workbook.defined_name import DefinedName  # noqa: E402
for name, lab, v in params:
    r += 1
    body(ws, r, 1, lab)
    body(ws, r, 2, v, FILL_INPUT, "0%" if v < 1 else "#,##0")
    wb.defined_names[name] = DefinedName(name, attr_text=f"'12_Data Quality Dashboard'!$B${r}")
r += 1
body(ws, r, 1, "配点（初期仮説：A35/B30/C20/D15）", FILL_SUB)
for name, v in weights:
    r += 1
    body(ws, r, 1, name.replace("W_", "満点 "))
    body(ws, r, 2, v, FILL_INPUT)
    wb.defined_names[name] = DefinedName(name, attr_text=f"'12_Data Quality Dashboard'!$B${r}")
r += 1
body(ws, r, 1, "合計")
body(ws, r, 2, f"=SUM(B{r-len(weights)}:B{r-1})", FILL_CALC)

r += 2
body(ws, r, 1, "業態別 調査進捗", FILL_SUB)
r += 1
cats = ["地方銀行", "第二地方銀行", "信用金庫", "信金中央金庫"]
heads = ["指標"] + cats + ["合計"]
for j, h in enumerate(heads, 1):
    c = body(ws, r, j, h)
    c.font, c.fill = F_HEAD, FILL_HEAD
U_CAT = UR("D")
SPR = lambda col: spref(col)  # noqa: E731
rows = [
    ("対象数（母集団）", lambda c: f'=COUNTIFS({U_CAT},"{c}")'),
    ("総資産 FY2025 取得済", lambda c: f'=COUNTIFS({U_CAT},"{c}",{SPR(SP["総資産", P0])},">-1E+300")'),
    ("有価証券内訳 FY2025 取得済", lambda c: f'=COUNTIFS({U_CAT},"{c}",{SPR(SP["国債", P0])},">-1E+300")'),
    ("外国証券 FY2025 取得済", lambda c: f'=COUNTIFS({U_CAT},"{c}",{SPR(SP["外国証券", P0])},">-1E+300")'),
    ("外国証券 3期そろい", lambda c: f'=COUNTIFS({U_CAT},"{c}",{SPR(SP["外国証券", P0])},">-1E+300",{SPR(SP["外国証券", P1])},">-1E+300",{SPR(SP["外国証券", P2])},">-1E+300")'),
    ("運用方針の記述あり（証拠件数>0）", lambda c: f'=SUMPRODUCT(({U_CAT}="{c}")*(COUNTIF(\'6_Investment Policy Evidence\'!$B$5:$B$5000,{UR("A")})>0))'),
    ("本ランキング対象（充足率≥閾値）", lambda c: f"=COUNTIFS('7_Target Scoring'!$C$6:$C${TS_LAST},\"{c}\",'7_Target Scoring'!${el}$6:${el}${TS_LAST},1)"),
]
for lab, fn in rows:
    r += 1
    body(ws, r, 1, lab)
    for j, c in enumerate(cats, 2):
        body(ws, r, j, fn(c), FILL_CALC)
    body(ws, r, len(cats) + 2, f"=SUM(B{r}:{L(len(cats)+1)}{r})", FILL_CALC)

r += 2
body(ws, r, 1, "データステータス別件数（3_Financial Data）", FILL_SUB)
r += 1
for j, h in enumerate(heads, 1):
    c = body(ws, r, j, h)
    c.font, c.fill = F_HEAD, FILL_HEAD
for stt in ["VERIFIED", "SECONDARY", "ESTIMATED", "NOT DISCLOSED", "NOT YET CHECKED", "CONFLICT"]:
    r += 1
    body(ws, r, 1, stt)
    for j, c in enumerate(cats, 2):
        body(ws, r, j, f"=COUNTIFS('3_Financial Data'!$J$2:$J${FD_MAX},\"{stt}\",'3_Financial Data'!$O$2:$O${FD_MAX},\"{c}\")", FILL_CALC)
    body(ws, r, len(cats) + 2, f"=SUM(B{r}:{L(len(cats)+1)}{r})", FILL_CALC)

r += 2
body(ws, r, 1, "分析セグメント別 件数（信金規模区分の分布確認用）", FILL_SUB)
segs = ["地銀・第二地銀", "信金・大", "信金・中", "信金・小", "信金（規模未判定）", "信金中金"]
for s in segs:
    r += 1
    body(ws, r, 1, s)
    body(ws, r, 2, f'=COUNTIFS({UR("E")},"{s}")', FILL_CALC)

r += 2
body(ws, r, 1, "地域別：母集団／選定段階（全国偏りの確認）", FILL_SUB)
r += 1
for j, h in enumerate(["地域", "母集団", "優先30", "詳細15", "提案10"], 1):
    c = body(ws, r, j, h)
    c.font, c.fill = F_HEAD, FILL_HEAD
PR_STAGE = f"'8_Priority Ranking'!$P$6:$P${5+N}"
PR_REG = f"'8_Priority Ranking'!$E$6:$E${5+N}"
for reg in ["北海道", "東北", "関東", "甲信越", "北陸", "東海", "近畿", "中国", "四国", "九州・沖縄"]:
    r += 1
    body(ws, r, 1, reg)
    body(ws, r, 2, f'=COUNTIFS({UR("H")},"{reg}")', FILL_CALC)
    for j, st in enumerate(["30", "15", "10"], 3):
        # 段階は包含関係（10⊂15⊂30）。入力値は最終到達段階
        cond = {"30": '{"30","15","10"}', "15": '{"15","10"}', "10": '{"10"}'}[st]
        body(ws, r, j, f"=SUMPRODUCT(({PR_REG}=A{r})*ISNUMBER(MATCH({PR_STAGE}&\"\",{cond},0)))", FILL_CALC)

r += 2
body(ws, r, 1, "バッチ実行記録", FILL_SUB)
r += 1
bh = ["バッチ", "実行日", "範囲", "調査対象数", "調査完了数", "未調査数", "データ取得成功数", "データ不足数", "次回調査対象", "備考"]
for j, h in enumerate(bh, 1):
    c = body(ws, r, j, h)
    c.font, c.fill = F_HEAD, FILL_HEAD
for b in batch:
    r += 1
    for j, k in enumerate(["batch", "run_date", "scope", "target", "completed", "not_started", "data_ok", "data_short", "next_target", "note"], 1):
        body(ws, r, j, num(b.get(k)) if k in ("target", "completed", "not_started", "data_ok", "data_short") else b.get(k, "")).alignment = WRAP

# 並び順：Dashboard を先頭に
wb.move_sheet("12_Data Quality Dashboard", offset=-(len(wb.sheetnames) - 1))
wb.calculation.fullCalcOnLoad = True
OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
print("saved", OUT, "sheets:", wb.sheetnames)
