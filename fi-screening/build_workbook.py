#!/usr/bin/env python3
"""全国金融機関ターゲット分析 — Excel ブック生成。

正本データは data/*.csv（公開情報のみ）。社内情報（営業カバレッジ・M1 フロー）は
リポジトリに置かず、生成したブックの入力シートに直接記入する。

    python3 build_workbook.py            # data/*.csv -> dist/FI_Target_Screening.xlsx
    python3 build_workbook.py --template # データを読まず空テンプレートを出力

設計の要点
- 数値は縦持ち（1 行 = 1 数値）で、出典・ページ・ステータスを必ず伴う。
- 分析シートは SUMIFS / COUNTIFS で縦持ちから抽出する。値が無い場合は
  「未調査」「非開示」「不一致」「重複」を文字で返し、0 と混同しない。
- 使用関数は Excel 2007 以降の標準関数のみ（動的配列関数は使わない）。
"""
import csv
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).parent
TEMPLATE = "--template" in sys.argv
OUT = ROOT / "dist" / ("FI_Target_Screening_template.xlsx" if TEMPLATE else "FI_Target_Screening.xlsx")

# ------------------------------------------------------------------ capacity
N_INST = 400          # Universe の行数（母集団 約350 ＋ 追加指定先の余裕）
FACT_BUFFER = 2000    # 縦持ち表に数式だけ先に用意しておく空行数
EVID_ROWS = 3000
PROFILE_ROWS = 2000
SRC_ROWS = 3000
R0 = 5                # データ開始行（1:タイトル 2:基準日 3:注記 4:見出し）
U_LAST = R0 + N_INST - 1

# ------------------------------------------------------------------ styling
JP = "Yu Gothic"
INK = "1F2933"
F_TITLE = Font(name=JP, size=14, bold=True, color=INK)
F_HEAD = Font(name=JP, size=9, bold=True, color="FFFFFF")
F_BODY = Font(name=JP, size=9, color=INK)
F_SMALL = Font(name=JP, size=9, color="52606D")
F_CONF = Font(name=JP, size=9, bold=True, color="9B1C1C")
FILL_HEAD = PatternFill("solid", fgColor="1C3F6E")
FILL_CALC = PatternFill("solid", fgColor="EEF2F7")      # 数式列（編集しない）
FILL_INPUT = PatternFill("solid", fgColor="FFFFFF")
FILL_INTERNAL = PatternFill("solid", fgColor="7A1F1F")  # 社内情報シートの見出し
FILL_WARN = PatternFill("solid", fgColor="F7EBD8")
FILL_BAD = PatternFill("solid", fgColor="F6DADA")
FILL_GOOD = PatternFill("solid", fgColor="DDEFE3")
_thin = Side(style="thin", color="D5DAE1")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
WRAP = Alignment(vertical="top", wrap_text=True)
HEAD_AL = Alignment(horizontal="left", vertical="center", wrap_text=True)

# ------------------------------------------------------------------ masters
GYOTAI = ["地方銀行", "第二地方銀行", "信用金庫", "信金中央金庫", "別途指定先", "信用組合（対象外）", "信用組合連合会（対象外）"]
STATUS = ["VERIFIED", "SECONDARY", "NOT DISCLOSED", "NOT YET CHECKED", "CONFLICT", "ESTIMATED"]
USABLE = ["VERIFIED", "SECONDARY", "ESTIMATED"]
BASIS = ["単体", "連結", "連結（持株会社）"]
UNITS = [("百万円", 1), ("千円", 0.001), ("億円", 100), ("円", 0.000001), ("%", None), ("年", None), ("件", None)]
TIER = ["1", "2", "3"]
DOC_TYPES = ["有価証券報告書", "決算短信", "決算説明資料", "決算補足資料", "ディスクロージャー誌", "統合報告書",
             "中期経営計画", "半期・四半期決算資料", "公式IR資料", "有価証券運用方針", "金融庁資料", "日本銀行資料",
             "業界団体資料", "法人番号公表サイト", "金融機関コード一覧", "金融専門メディア", "格付レポート", "その他"]
THEMES = ["外債投資の拡大・縮小方針", "米国金利見通し", "為替ヘッジ方針", "ポートフォリオ再構築", "収益性向上",
          "債券運用のリスク管理", "投資信託・ファンド活用", "外部運用会社の活用", "複雑商品への投資", "その他"]
DIRECTION = ["拡大", "維持", "縮小", "入替", "言及なし", "—"]
JUDGE = ["明確", "示唆", "言及なし", "評価保留"]
STMT_KIND = ["事実", "分析上の推論", "営業仮説"]
CONFIDENCE = ["高", "中", "低", "—"]
PROFILE_SECTIONS = ["1 なぜターゲット候補か", "2 公開情報から確認できる事実", "3 推定される運用課題",
                    "4 推定の根拠と確信度", "5 商品の提案仮説", "6 顧客に確認すべき質問",
                    "7 営業アプローチの切り口", "8 提案に適さない可能性がある理由"]
DECLINE = ["売却", "償還", "為替換算", "時価変動", "複合", "不明（断定不可）"]
CANDIDATE = ["初期提案候補（10）", "詳細分析対象（15）", "優先調査対象（30）", "保留（データ不足）", "対象外（適合性懸念）", "未選定"]
UNIV_STATUS = ["未着手", "母集団確認済", "定量調査中", "定量完了", "定性完了", "詳細分析完了"]
EXISTENCE = ["現存", "消滅（合併）", "要確認"]
COVERAGE = ["カバレッジ中", "カバレッジ外", "過去カバレッジ", "不明"]
YESNO3 = ["可", "不可", "条件付き", "不明"]
RATING = ["A", "B", "C", "未評価"]
CONF_CLASS = ["社外秘", "極秘", "公開情報"]
ACTION_STATE = ["未着手", "準備中", "実施済", "保留", "中止"]
IDENTIFIABLE = ["可", "不可", "不明"]
PREF_BLOCK = [
    ("北海道・東北", "北海道 青森県 岩手県 宮城県 秋田県 山形県 福島県"),
    ("関東", "茨城県 栃木県 群馬県 埼玉県 千葉県 東京都 神奈川県"),
    ("甲信越・北陸", "新潟県 富山県 石川県 福井県 山梨県 長野県"),
    ("東海", "岐阜県 静岡県 愛知県 三重県"),
    ("近畿", "滋賀県 京都府 大阪府 兵庫県 奈良県 和歌山県"),
    ("中国", "鳥取県 島根県 岡山県 広島県 山口県"),
    ("四国", "徳島県 香川県 愛媛県 高知県"),
    ("九州・沖縄", "福岡県 佐賀県 長崎県 熊本県 大分県 宮崎県 鹿児島県 沖縄県"),
]
PREFS = [(p, b) for b, ps in PREF_BLOCK for p in ps.split()]

# 指標マスタ：code, 名称, 格納シート(FD=Financial Data / SP=Securities Portfolio), 既定単位, 定義・注意
METRICS = [
    ("TA", "総資産", "FD", "百万円", "期末残高"),
    ("DEP", "預金残高", "FD", "百万円", "譲渡性預金を含まない。含む場合は NCD を別記し備考に明記"),
    ("NCD", "譲渡性預金", "FD", "百万円", ""),
    ("LOAN", "貸出金残高", "FD", "百万円", ""),
    ("SEC", "有価証券残高", "FD", "百万円", "B/S 計上額"),
    ("CAR", "自己資本比率", "FD", "%", "国内基準。国際統一基準行は CET1 を別記し備考に明記"),
    ("CET1", "普通株式等Tier1比率", "FD", "%", "国際統一基準行のみ"),
    ("JGB", "国債", "SP", "百万円", "B/S 区分"),
    ("MUNI", "地方債", "SP", "百万円", ""),
    ("CORP", "社債", "SP", "百万円", "短期社債を含むかは備考に明記"),
    ("EQ", "株式", "SP", "百万円", ""),
    ("OTHSEC", "その他の証券", "SP", "百万円", "B/S 区分。外国証券・投資信託等を含む"),
    ("FXSEC", "外国証券", "SP", "百万円", "外国債券・外国株式・外国投信等を含みうる。米国債と同一視しない"),
    ("FXBOND", "外国債券", "SP", "百万円", "外国証券の内訳として開示がある場合"),
    ("FXEQ", "外国株式", "SP", "百万円", ""),
    ("INVTR", "投資信託", "SP", "百万円", "国内・外国の別を備考に明記"),
    ("FXCCYBOND", "外貨建債券", "SP", "百万円", "説明資料の『外貨建て有価証券』等。円建て外国債券を含まない定義か確認"),
    ("UST", "米国債", "SP", "百万円", "明示的な開示がある場合のみ"),
    ("FXHEDGED", "ヘッジ付き外債残高", "SP", "百万円", ""),
    ("FXUNHEDGED", "ヘッジなし外債残高", "SP", "百万円", ""),
    ("HEDGE_RATIO", "為替ヘッジ比率", "SP", "%", "開示値のみ。推計時は ESTIMATED と計算式"),
    ("UNREAL_ALL", "その他有価証券 評価差額（合計）", "SP", "百万円", "評価差額＝時価−取得原価"),
    ("UNREAL_BOND", "その他有価証券 評価差額（債券）", "SP", "百万円", "円債"),
    ("UNREAL_FX", "その他有価証券 評価差額（外国証券・外債）", "SP", "百万円", "どの範囲かを備考に明記"),
    ("DUR_JPY", "円債デュレーション", "SP", "年", "定義（修正D等）を備考に明記"),
    ("DUR_FX", "外債デュレーション", "SP", "年", "同上"),
    ("FXB_M01", "外国債券 残存1年以内", "SP", "百万円", "ディスクロ誌 残存期間別残高"),
    ("FXB_M13", "外国債券 残存1-3年", "SP", "百万円", ""),
    ("FXB_M35", "外国債券 残存3-5年", "SP", "百万円", ""),
    ("FXB_M57", "外国債券 残存5-7年", "SP", "百万円", ""),
    ("FXB_M710", "外国債券 残存7-10年", "SP", "百万円", ""),
    ("FXB_M10P", "外国債券 残存10年超", "SP", "百万円", ""),
]
MET = {m[0]: m for m in METRICS}

# スコア項目：code, 観点, 名称, 既定配点, 種別(quant/qual/bucket/combo)
SCORE_ITEMS = [
    ("A1", "A", "外国証券残高", 12, "quant"),
    ("A2", "A", "外国証券／有価証券比率", 10, "quant"),
    ("A3", "A", "外債投資方針の記述", 8, "qual"),
    ("A4", "A", "為替ヘッジの利用状況", 5, "qual"),
    ("B1", "B", "外債運用見直しの開示", 10, "qual"),
    ("B2", "B", "再投資ニーズを示唆する情報", 8, "qual"),
    ("B3", "B", "ヘッジコストへの問題意識", 7, "qual"),
    ("B4", "B", "金利・為替リスク管理上の課題", 5, "qual"),
    ("C1", "C", "有価証券運用規模", 8, "quant"),
    ("C2", "C", "自己資本比率（規制水準に対する余裕）", 6, "bucket"),
    ("C3", "C", "運用余力指標（預貸率・評価損益率）", 6, "combo"),
    ("D1", "D", "投資信託・ファンドの利用", 7, "qual"),
    ("D2", "D", "外部運用会社の活用", 4, "qual"),
    ("D3", "D", "複雑な運用商品への投資経験", 4, "qual"),
]


# ------------------------------------------------------------------ helpers
def read_csv(name):
    p = ROOT / "data" / name
    if TEMPLATE or not p.exists():
        return []
    with p.open(encoding="utf-8-sig", newline="") as f:
        return [r for r in csv.DictReader(f) if any(v.strip() for v in r.values() if v)]


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def col_range(sheet, col, first, last):
    return f"'{sheet}'!${col}${first}:${col}${last}"


def sheet_frame(ws, title, note, cols, internal=False):
    """cols: list of (header, width, kind) kind in {'in','calc','key'}"""
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = F_TITLE
    ws["A2"] = ('="母集団基準日: "&TEXT(CFG_BASEDATE,"yyyy-mm-dd")&"　分析対象期: "&CFG_FY0'
                '&"（前期 "&CFG_FY1&"／前々期 "&CFG_FY2&"）　比較基準: "&CFG_BASIS')
    ws["A2"].font = F_SMALL
    ws["A3"] = note
    ws["A3"].font = F_CONF if internal else F_SMALL
    for i, (h, w, kind) in enumerate(cols, start=1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = F_HEAD
        c.fill = FILL_INTERNAL if internal else FILL_HEAD
        c.alignment = HEAD_AL
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[4].height = 36
    ws.freeze_panes = ws.cell(row=R0, column=3)
    return {h: get_column_letter(i) for i, (h, _, _) in enumerate(cols, start=1)}


def fill_rows(ws, cols, first, last, formulas=None, values=None, text_cols=(), fmt=None):
    """formulas: {header: lambda r: '=...'}; values: list of dict rows (header -> value)"""
    formulas = formulas or {}
    fmt = fmt or {}
    letters = {h: i for i, (h, _, _) in enumerate(cols, start=1)}
    for r in range(first, last + 1):
        row_vals = values[r - first] if values and r - first < len(values) else {}
        for h, (_, _, kind) in zip(letters, cols):
            ci = letters[h]
            c = ws.cell(row=r, column=ci)
            if h in formulas:
                c.value = formulas[h](r)
                c.fill = FILL_CALC
            elif h in row_vals and row_vals[h] not in (None, ""):
                v = row_vals[h]
                if h not in text_cols:
                    try:
                        v = float(v) if "." in v or "e" in v.lower() else int(v)
                    except (ValueError, AttributeError):
                        pass
                c.value = v
            c.font = F_BODY
            c.border = BORDER
            if h in text_cols:
                c.number_format = "@"
            if h in fmt:
                c.number_format = fmt[h]
    ws.auto_filter.ref = f"A4:{get_column_letter(len(cols))}{last}"


def dv_list(ws, named_list, ref):
    dv = DataValidation(type="list", formula1=f"={named_list}", allow_blank=True)
    dv.error = "リストから選択してください"
    dv.errorStyle = "warning"
    ws.add_data_validation(dv)
    dv.add(ref)


def lookup(key_cell, ret_range, key_range, blank=""):
    return f'=IF({key_cell}="","",IFERROR(INDEX({ret_range},MATCH({key_cell},{key_range},0)),"{blank}"))'


# ================================================================== workbook
wb = Workbook()
wb.remove(wb.active)
order = ["README", "1 Institution Universe", "2 Source Register", "3 Financial Data", "4 Securities Portfolio",
         "5 Foreign Bond Analysis", "6 Investment Policy Evidence", "7 Target Scoring", "8 Priority Ranking",
         "9 Bank Profiles", "10 Sales Coverage Input", "10b M1 Flow Input", "11 Sales Action Plan",
         "12 Data Quality Dashboard", "Metric Master", "Config", "Lists"]
S = {k: wb.create_sheet(k) for k in order}
UNI, SRC, FD, SP, FBA, EVI, SCO, RNK, PRO, COV, M1, ACT, DQ, MM, CFG, LST = (
    "1 Institution Universe", "2 Source Register", "3 Financial Data", "4 Securities Portfolio",
    "5 Foreign Bond Analysis", "6 Investment Policy Evidence", "7 Target Scoring", "8 Priority Ranking",
    "9 Bank Profiles", "10 Sales Coverage Input", "10b M1 Flow Input", "11 Sales Action Plan",
    "12 Data Quality Dashboard", "Metric Master", "Config", "Lists")

# ------------------------------------------------------------------ Lists
ws = S[LST]
lists = [("L_GYOTAI", "業態", GYOTAI), ("L_STATUS", "データステータス", STATUS), ("L_BASIS", "単体／連結", BASIS),
         ("L_UNIT", "単位", [u for u, _ in UNITS]), ("L_TIER", "Tier", TIER), ("L_DOCTYPE", "資料種別", DOC_TYPES),
         ("L_THEME", "テーマ", THEMES), ("L_DIR", "方向性", DIRECTION), ("L_JUDGE", "判定", JUDGE),
         ("L_KIND", "記述区分", STMT_KIND), ("L_CONF", "確信度", CONFIDENCE), ("L_SECTION", "プロファイル項目", PROFILE_SECTIONS),
         ("L_DECLINE", "減少要因", DECLINE), ("L_CAND", "候補区分", CANDIDATE), ("L_USTAT", "調査ステータス", UNIV_STATUS),
         ("L_EXIST", "存否", EXISTENCE), ("L_COVER", "カバレッジ", COVERAGE), ("L_YN3", "可否", YESNO3),
         ("L_RATING", "担当者評価", RATING), ("L_CCLASS", "機密区分", CONF_CLASS), ("L_ASTATE", "アクション状態", ACTION_STATE),
         ("L_IDENT", "個社特定可否", IDENTIFIABLE), ("L_SCOREITEM", "スコア項目", [s[0] for s in SCORE_ITEMS]),
         ("L_METRIC", "指標コード", [m[0] for m in METRICS])]
for i, (nm, label, vals) in enumerate(lists, start=1):
    L = get_column_letter(i)
    ws.cell(row=1, column=i, value=label).font = F_HEAD
    ws.cell(row=1, column=i).fill = FILL_HEAD
    ws.column_dimensions[L].width = 18
    for j, v in enumerate(vals, start=2):
        ws.cell(row=j, column=i, value=v).font = F_BODY
    name(wb, nm, f"'{LST}'!${L}$2:${L}${len(vals) + 1}")
# 単位換算表・都道府県→地域ブロック
c0 = len(lists) + 2
for k, h in enumerate(["単位", "百万円換算係数"]):
    ws.cell(row=1, column=c0 + k, value=h).font = F_HEAD
    ws.cell(row=1, column=c0 + k).fill = FILL_HEAD
for j, (u, f) in enumerate(UNITS, start=2):
    ws.cell(row=j, column=c0, value=u)
    ws.cell(row=j, column=c0 + 1, value=f if f is not None else "—")
U1, U2 = get_column_letter(c0), get_column_letter(c0 + 1)
name(wb, "T_UNIT", f"'{LST}'!${U1}$2:${U1}${len(UNITS) + 1}")
name(wb, "T_UNITF", f"'{LST}'!${U2}$2:${U2}${len(UNITS) + 1}")
c1 = c0 + 3
for k, h in enumerate(["都道府県", "地域ブロック"]):
    ws.cell(row=1, column=c1 + k, value=h).font = F_HEAD
    ws.cell(row=1, column=c1 + k).fill = FILL_HEAD
for j, (p, b) in enumerate(PREFS, start=2):
    ws.cell(row=j, column=c1, value=p)
    ws.cell(row=j, column=c1 + 1, value=b)
P1, P2 = get_column_letter(c1), get_column_letter(c1 + 1)
name(wb, "L_PREF", f"'{LST}'!${P1}$2:${P1}${len(PREFS) + 1}")
name(wb, "T_PREFBLOCK", f"'{LST}'!${P2}$2:${P2}${len(PREFS) + 1}")
blocks = [b for b, _ in PREF_BLOCK]
ws.cell(row=1, column=c1 + 3, value="地域ブロック").font = F_HEAD
ws.cell(row=1, column=c1 + 3).fill = FILL_HEAD
for j, b in enumerate(blocks, start=2):
    ws.cell(row=j, column=c1 + 3, value=b)
B3 = get_column_letter(c1 + 3)
name(wb, "L_BLOCK", f"'{LST}'!${B3}$2:${B3}${len(blocks) + 1}")

# ------------------------------------------------------------------ Config
ws = S[CFG]
ws.sheet_view.showGridLines = False
ws["A1"] = "Config — 基準日・閾値・スコア配点（ここを変えると全シートが再計算される）"
ws["A1"].font = F_TITLE
ws.column_dimensions["A"].width = 34
ws.column_dimensions["B"].width = 18
ws.column_dimensions["C"].width = 70
import datetime as _dt  # noqa: E402

cfg = [
    ("CFG_BASEDATE", "母集団基準日", _dt.date(2026, 9, 30), "合併・商号変更はこの日時点で判定。以降の変更は Universe の変更イベント列に記録"),
    ("CFG_FY0", "分析対象期（当期）", "2026/03", "縦持ち表の『期』列と完全一致させる（yyyy/mm）"),
    ("CFG_FY1", "前期", "2025/03", ""),
    ("CFG_FY2", "前々期", "2024/03", ""),
    ("CFG_BASIS", "比較基準（単体／連結）", "単体", "分析シートはこの基準の値だけを抽出する"),
    ("CFG_SK_LARGE", "信金 大規模の下限（預金・百万円）", 1000000, "1兆円。分布確認後に見直す"),
    ("CFG_SK_MID", "信金 中規模の下限（預金・百万円）", 300000, "3,000億円"),
    ("CFG_MIN_PEER", "パーセンタイル計算の最少ピア件数", 3, "同一ピアで数値が揃った先がこれ未満なら評価保留"),
    ("CFG_PARTIAL", "定性『示唆』の配点比率", 0.6, "明確=100%／示唆=この比率／言及なし=0／評価保留=算入しない"),
    ("CFG_MIN_COVER", "順位付けに必要なデータ充足率", 0.4, "未満は『データ不足』として順位から分離。定量項目（A1/A2/C1/C2/C3）の配点合計が42のため、0.4＝定量がほぼ揃えば順位対象"),
    ("CFG_C2_T1", "C2 満点の自己資本比率（%）", 12, "国内基準。国際統一基準行は Target Scoring の C2 上書き列を使用"),
    ("CFG_C2_T2", "C2 4点の下限（%）", 10, ""),
    ("CFG_C2_T3", "C2 2点の下限（%）", 8, ""),
    ("CFG_TYPE_TOP", "Type A/D 上位判定パーセンタイル", 0.75, ""),
    ("CFG_TYPE_LOW", "Type D 外債比率 低位判定パーセンタイル", 0.5, ""),
    ("CFG_TYPE_CHG", "Type B/C 前期比の閾値（絶対値）", 0.2, "±20%"),
]
r = 3
for k, h in enumerate(["項目", "値", "説明"]):
    ws.cell(row=r, column=1 + k, value=h).font = F_HEAD
    ws.cell(row=r, column=1 + k).fill = FILL_HEAD
for nm, label, val, note in cfg:
    r += 1
    ws.cell(row=r, column=1, value=label).font = F_BODY
    c = ws.cell(row=r, column=2, value=val)
    c.font = F_BODY
    c.border = BORDER
    if isinstance(val, _dt.date):
        c.number_format = "yyyy-mm-dd"
    elif isinstance(val, str):
        c.number_format = "@"
    ws.cell(row=r, column=3, value=note).font = F_SMALL
    name(wb, nm, f"'{CFG}'!$B${r}")
r += 2
ws.cell(row=r, column=1, value="スコア配点（初期仮説。業態別分布の確認後に見直す）").font = F_TITLE
r += 1
for k, h in enumerate(["項目", "配点", "名称"]):
    ws.cell(row=r, column=1 + k, value=h).font = F_HEAD
    ws.cell(row=r, column=1 + k).fill = FILL_HEAD
w_first = r + 1
for code, axis, label, w, kind in SCORE_ITEMS:
    r += 1
    ws.cell(row=r, column=1, value=f"{code}（{axis}）").font = F_BODY
    ws.cell(row=r, column=2, value=w).border = BORDER
    ws.cell(row=r, column=3, value=f"{label}［{kind}］").font = F_SMALL
    name(wb, f"W_{code}", f"'{CFG}'!$B${r}")
r += 1
ws.cell(row=r, column=1, value="配点合計（100 であること）").font = F_BODY
ws.cell(row=r, column=2, value=f"=SUM(B{w_first}:B{r - 1})")
ws.cell(row=r, column=3, value=f'=IF(B{r}=100,"OK","配点合計が100ではありません")').font = F_SMALL
r += 2
ws.cell(row=r, column=1, value="改訂履歴").font = F_TITLE
r += 1
for k, h in enumerate(["日付", "変更者", "変更内容・理由"]):
    ws.cell(row=r, column=1 + k, value=h).font = F_HEAD
    ws.cell(row=r, column=1 + k).fill = FILL_HEAD
r += 1
ws.cell(row=r, column=1, value="2026-11-05")
ws.cell(row=r, column=3, value="初版（調査設計書 v0.1 の初期仮説）")

# ------------------------------------------------------------------ Metric Master
ws = S[MM]
cols = [("指標コード", 14, "key"), ("指標名", 30, "in"), ("格納シート", 12, "in"), ("既定単位", 10, "in"), ("定義・注意", 70, "in")]
L = sheet_frame(ws, "Metric Master — 指標定義", "指標コードは縦持ち表の入力規則に使われる。追加時は build_workbook.py の METRICS に追記", cols)
for r, m in enumerate(METRICS, start=R0):
    for i, v in enumerate([m[0], m[1], {"FD": FD, "SP": SP}[m[2]], m[3], m[4]], start=1):
        c = ws.cell(row=r, column=i, value=v)
        c.font = F_BODY
        c.border = BORDER
MM_LAST = R0 + len(METRICS) - 1
name(wb, "MM_CODE", col_range(MM, "A", R0, MM_LAST))
name(wb, "MM_NAME", col_range(MM, "B", R0, MM_LAST))
name(wb, "MM_SHEET", col_range(MM, "C", R0, MM_LAST))

# ------------------------------------------------------------------ 1 Universe
universe = read_csv("universe.csv")
ws = S[UNI]
ucols = [("金融機関コード", 10, "key"), ("法人番号", 15, "in"), ("正式名称", 26, "in"), ("略称", 14, "in"),
         ("業態", 12, "in"), ("信金規模区分", 10, "calc"), ("ピアグループ", 16, "calc"), ("本店都道府県", 10, "in"),
         ("地域ブロック", 12, "calc"), ("持株会社名", 20, "in"), ("持株会社法人番号", 15, "in"), ("上場区分", 10, "in"),
         ("決算期", 8, "in"), ("変更イベント（合併・商号変更・業態変更）", 30, "in"), ("存否", 10, "in"),
         ("母集団SourceID", 12, "in"), ("確認日", 11, "in"), ("調査ステータス", 12, "in"), ("バッチ番号", 8, "in"),
         ("預金残高（当期・百万円）", 14, "calc"), ("重複チェック", 10, "calc"), ("備考", 30, "in")]
UL = sheet_frame(ws, "1 Institution Universe — 調査母集団（1 行 = 1 投資主体）",
                 "持株会社は行を作らず属性として記録。信金中央金庫は個別信金と別主体。灰色の列は数式（編集しない）", ucols)
uv = [{h: row.get(h, "") for h, _, _ in ucols} for row in universe]
fill_rows(ws, ucols, R0, U_LAST, values=uv,
          text_cols=("金融機関コード", "法人番号", "持株会社法人番号", "母集団SourceID", "決算期"),
          formulas={
              "信金規模区分": lambda r: (f'=IF(E{r}<>"信用金庫","",IF(NOT(ISNUMBER(T{r})),"未確定",'
                                    f'IF(T{r}>=CFG_SK_LARGE,"大規模",IF(T{r}>=CFG_SK_MID,"中規模","小規模"))))'),
              "ピアグループ": lambda r: f'=IF(E{r}="","",IF(E{r}="信用金庫","信金・"&F{r},E{r}))',
              "地域ブロック": lambda r: f'=IF(H{r}="","",IFERROR(INDEX(T_PREFBLOCK,MATCH(H{r},L_PREF,0)),"要確認"))',
              "預金残高（当期・百万円）": lambda r: None,  # 後で FD 名前定義後に設定
              "重複チェック": lambda r: f'=IF(A{r}="","",IF(COUNTIF($A${R0}:$A${U_LAST},A{r})>1,"重複","OK"))',
          })
dv_list(ws, "L_GYOTAI", f"E{R0}:E{U_LAST}")
dv_list(ws, "L_PREF", f"H{R0}:H{U_LAST}")
dv_list(ws, "L_EXIST", f"O{R0}:O{U_LAST}")
dv_list(ws, "L_USTAT", f"R{R0}:R{U_LAST}")
ws.conditional_formatting.add(f"U{R0}:U{U_LAST}", CellIsRule(operator="equal", formula=['"重複"'], fill=FILL_BAD))
name(wb, "U_CODE", col_range(UNI, "A", R0, U_LAST))
name(wb, "U_NAME", col_range(UNI, "C", R0, U_LAST))
name(wb, "U_GYOTAI", col_range(UNI, "E", R0, U_LAST))
name(wb, "U_PEER", col_range(UNI, "G", R0, U_LAST))
name(wb, "U_BLOCK", col_range(UNI, "I", R0, U_LAST))
name(wb, "U_USTAT", col_range(UNI, "R", R0, U_LAST))

# ------------------------------------------------------------------ 2 Source Register
sources = read_csv("sources.csv")
ws = S[SRC]
scols = [("SourceID", 12, "key"), ("金融機関コード", 10, "in"), ("金融機関名", 20, "calc"), ("資料名", 36, "in"),
         ("資料種別", 16, "in"), ("Tier", 6, "in"), ("対象期間", 10, "in"), ("単体／連結", 10, "in"),
         ("公表日", 11, "in"), ("URL", 48, "in"), ("確認日", 11, "in"), ("取得状態", 10, "in"),
         ("表の読取状態", 12, "in"), ("備考", 36, "in")]
S_LAST = R0 + max(SRC_ROWS, len(sources) + 500) - 1
fill_rows(ws, scols, R0, S_LAST, values=[{h: s.get(h, "") for h, _, _ in scols} for s in sources],
          text_cols=("SourceID", "金融機関コード", "対象期間"),
          formulas={"金融機関名": lambda r: lookup(f"B{r}", "U_NAME", "U_CODE", "母集団外")},
          fmt={"公表日": "yyyy-mm-dd", "確認日": "yyyy-mm-dd"})
sheet_frame(ws, "2 Source Register — 出典台帳（1 行 = 1 資料）",
            "金融機関コード空欄＝業界・公的資料。取得状態：取得済／取得不可／未取得。表の読取状態：読取可／一部不可／不可（不可の表の数値は推測しない）", scols)
dv_list(ws, "L_DOCTYPE", f"E{R0}:E{S_LAST}")
dv_list(ws, "L_TIER", f"F{R0}:F{S_LAST}")
dv_list(ws, "L_BASIS", f"H{R0}:H{S_LAST}")
name(wb, "S_ID", col_range(SRC, "A", R0, S_LAST))
name(wb, "S_DOC", col_range(SRC, "D", R0, S_LAST))
name(wb, "S_TIER", col_range(SRC, "F", R0, S_LAST))
name(wb, "S_PUB", col_range(SRC, "I", R0, S_LAST))
name(wb, "S_URL", col_range(SRC, "J", R0, S_LAST))

# ------------------------------------------------------------------ 3/4 fact tables
points = read_csv("datapoints.csv")
fcols = [("RecordID", 12, "key"), ("金融機関コード", 10, "in"), ("金融機関名", 20, "calc"), ("指標コード", 12, "in"),
         ("指標名", 22, "calc"), ("期", 9, "in"), ("期末日", 11, "in"), ("単体／連結", 10, "in"), ("数値", 14, "in"),
         ("単位", 8, "in"), ("換算値（百万円／%／年）", 14, "calc"), ("SourceID", 12, "in"), ("ページ", 8, "in"),
         ("ステータス", 15, "in"), ("確認日", 11, "in"), ("計算式・推計根拠（ESTIMATED は必須）", 30, "in"),
         ("出典URL", 30, "calc"), ("資料公表日", 11, "calc"), ("業態", 12, "calc"), ("有効", 6, "calc"),
         ("整合チェック", 18, "calc"), ("備考", 30, "in")]
FACT = {}
for key, sheet, title in (("FD", FD, "3 Financial Data — 基本財務（1 行 = 1 数値）"),
                          ("SP", SP, "4 Securities Portfolio — 有価証券・外債（1 行 = 1 数値）")):
    rows = [p for p in points if MET.get(p.get("指標コード", ""), ("", "", ""))[2] == key]
    ws = S[sheet]
    last = R0 + len(rows) + FACT_BUFFER - 1
    usable = "+".join(f'(N{{r}}="{s}")' for s in USABLE)

    def ok(r, usable=usable):
        return f'=IF(B{r}="","",IF({usable.format(r=r)},1,0))'

    def check(r, last=last):
        key_ = (f'$B${R0}:$B${last},B{r},$D${R0}:$D${last},D{r},$F${R0}:$F${last},F{r},'
                f'$H${R0}:$H${last},H{r},$T${R0}:$T${last},1')
        return (f'=IF(B{r}="","",IF(COUNTIF(U_CODE,B{r})=0,"母集団外コード",'
                f'IF(COUNTIF(MM_CODE,D{r})=0,"指標コード不明",'
                f'IF(AND(N{r}="ESTIMATED",P{r}=""),"計算式未記入",'
                f'IF(AND(OR(N{r}="VERIFIED",N{r}="NOT DISCLOSED"),L{r}=""),"SourceID未記入",'
                f'IF(AND(T{r}=1,COUNTIFS({key_})>1),"有効値が重複","OK"))))))')

    fill_rows(ws, fcols, R0, last, values=[{h: p.get(h, "") for h, _, _ in fcols} for p in rows],
              text_cols=("RecordID", "金融機関コード", "期", "SourceID", "ページ"),
              fmt={"期末日": "yyyy-mm-dd", "確認日": "yyyy-mm-dd", "資料公表日": "yyyy-mm-dd",
                   "数値": "#,##0.##", "換算値（百万円／%／年）": "#,##0.##"},
              formulas={
                  "金融機関名": lambda r: lookup(f"B{r}", "U_NAME", "U_CODE", "母集団外"),
                  "指標名": lambda r: lookup(f"D{r}", "MM_NAME", "MM_CODE", "指標コード不明"),
                  "換算値（百万円／%／年）": lambda r: (f'=IF(OR(B{r}="",NOT(ISNUMBER(I{r}))),"",'
                                              f'IFERROR(IF(ISNUMBER(INDEX(T_UNITF,MATCH(J{r},T_UNIT,0))),'
                                              f'I{r}*INDEX(T_UNITF,MATCH(J{r},T_UNIT,0)),I{r}),"単位不明"))'),
                  "出典URL": lambda r: lookup(f"L{r}", "S_URL", "S_ID", "台帳未登録"),
                  "資料公表日": lambda r: lookup(f"L{r}", "S_PUB", "S_ID", "台帳未登録"),
                  "業態": lambda r: lookup(f"B{r}", "U_GYOTAI", "U_CODE", "母集団外"),
                  "有効": ok,
                  "整合チェック": check,
              })
    sheet_frame(ws, title, "有効＝VERIFIED/SECONDARY/ESTIMATED。分析シートは同一コード×指標×期×基準で有効値が 1 件のときだけ値を使う。"
                           "NOT DISCLOSED は確認した資料の SourceID 必須。読めない表の数値は入れない", fcols)
    dv_list(ws, "L_METRIC", f"D{R0}:D{last}")
    dv_list(ws, "L_BASIS", f"H{R0}:H{last}")
    dv_list(ws, "L_UNIT", f"J{R0}:J{last}")
    dv_list(ws, "L_STATUS", f"N{R0}:N{last}")
    ws.conditional_formatting.add(f"U{R0}:U{last}", FormulaRule(formula=[f'AND(U{R0}<>"",U{R0}<>"OK")'], fill=FILL_BAD))
    ws.conditional_formatting.add(f"N{R0}:N{last}", CellIsRule(operator="equal", formula=['"NOT YET CHECKED"'], fill=FILL_WARN))
    ws.conditional_formatting.add(f"N{R0}:N{last}", CellIsRule(operator="equal", formula=['"CONFLICT"'], fill=FILL_BAD))
    for nm, col in (("CODE", "B"), ("MET", "D"), ("PER", "F"), ("BASIS", "H"), ("VAL", "K"), ("STAT", "N"),
                    ("OK", "T"), ("GYO", "S")):
        name(wb, f"{key}_{nm}", col_range(sheet, col, R0, last))
    FACT[key] = last


def pull(code_cell, metric, period):
    """縦持ち表から 1 値を抽出。無ければ理由を文字で返す。"""
    k = MET[metric][2]
    base = f'{k}_CODE,{code_cell},{k}_MET,"{metric}",{k}_PER,{period},{k}_BASIS,CFG_BASIS'
    return (f'IF(COUNTIFS({base},{k}_OK,1)=1,SUMIFS({k}_VAL,{base},{k}_OK,1),'
            f'IF(COUNTIFS({base},{k}_OK,1)>1,"重複",'
            f'IF(COUNTIFS({base},{k}_STAT,"CONFLICT")>0,"不一致",'
            f'IF(COUNTIFS({base},{k}_STAT,"NOT DISCLOSED")>0,"非開示","未調査"))))')


# Universe の預金残高（規模区分用）
ws = S[UNI]
for r in range(R0, U_LAST + 1):
    ws[f"T{r}"] = f'=IF(A{r}="","",{pull(f"A{r}", "DEP", "CFG_FY0")})'

# ------------------------------------------------------------------ 5 Foreign Bond Analysis
ws = S[FBA]
aspec = [  # header, width, formula(r) or None for input, number format
    ("金融機関コード", 10, lambda r: f"=IF('{UNI}'!A{r}=\"\",\"\",'{UNI}'!A{r})", "@"),
    ("金融機関名", 20, lambda r: f"=IF(A{r}=\"\",\"\",'{UNI}'!C{r})", None),
    ("業態", 11, lambda r: f"=IF(A{r}=\"\",\"\",'{UNI}'!E{r})", None),
    ("ピアグループ", 14, lambda r: f"=IF(A{r}=\"\",\"\",'{UNI}'!G{r})", None),
    ("地域ブロック", 11, lambda r: f"=IF(A{r}=\"\",\"\",'{UNI}'!I{r})", None),
    ("総資産", 12, ("TA", "CFG_FY0"), "#,##0"),
    ("預金", 12, ("DEP", "CFG_FY0"), "#,##0"),
    ("貸出金", 12, ("LOAN", "CFG_FY0"), "#,##0"),
    ("有価証券", 12, ("SEC", "CFG_FY0"), "#,##0"),
    ("有価証券（前期）", 12, ("SEC", "CFG_FY1"), "#,##0"),
    ("有価証券／総資産", 9, ("ratio", "I", "F"), "0.0%"),
    ("預貸率", 9, ("ratio", "H", "G"), "0.0%"),
    ("自己資本比率(%)", 9, ("CAR", "CFG_FY0"), "0.00"),
    ("外国証券（当期）", 12, ("FXSEC", "CFG_FY0"), "#,##0"),
    ("外国証券（前期）", 12, ("FXSEC", "CFG_FY1"), "#,##0"),
    ("外国証券（前々期）", 12, ("FXSEC", "CFG_FY2"), "#,##0"),
    ("外国証券／有価証券", 9, ("ratio", "N", "I"), "0.0%"),
    ("外国証券 前期比（額）", 12, ("diff", "N", "O"), "#,##0;[Red]-#,##0"),
    ("外国証券 前期比（%）", 9, ("chg", "N", "O"), "0.0%;[Red]-0.0%"),
    ("外国証券 前々期比（%）", 9, ("chg", "O", "P"), "0.0%;[Red]-0.0%"),
    ("2期連続減少", 8, ("down2",), None),
    ("外国債券", 12, ("FXBOND", "CFG_FY0"), "#,##0"),
    ("外貨建債券", 12, ("FXCCYBOND", "CFG_FY0"), "#,##0"),
    ("米国債", 12, ("UST", "CFG_FY0"), "#,##0"),
    ("ヘッジ比率(%)", 8, ("HEDGE_RATIO", "CFG_FY0"), "0.0"),
    ("外債評価差額", 12, ("UNREAL_FX", "CFG_FY0"), "#,##0;[Red]-#,##0"),
    ("有価証券評価差額（合計）", 12, ("UNREAL_ALL", "CFG_FY0"), "#,##0;[Red]-#,##0"),
    ("評価差額／有価証券", 9, ("ratio", "AA", "I"), "0.0%;[Red]-0.0%"),
    ("外債デュレーション(年)", 8, ("DUR_FX", "CFG_FY0"), "0.0"),
    ("外国債券 残存1年以内", 12, ("FXB_M01", "CFG_FY0"), "#,##0"),
    ("外国証券 減少要因（入力・不明なら断定しない）", 16, None, None),
    ("減少要因の根拠ID", 12, None, None),
    ("備考", 30, None, None),
]
acols = [(h, w, "calc" if f else "in") for h, w, f, _ in aspec]
AL = sheet_frame(ws, "5 Foreign Bond Analysis — 外債関連の横持ち分析（行は Universe と同じ並び）",
                 "数値は比較基準・同一期の有効値のみ（百万円）。文字（未調査／非開示／不一致／重複）は数値として扱わない。"
                 "外国証券＝外国債券・外国株式・外国投信等の合計であり米国債ではない", acols)


def afml(spec, r):
    if callable(spec):
        return spec(r)
    t = spec[0]
    num = lambda c: f"ISNUMBER({c}{r})"  # noqa: E731
    if t == "ratio":
        a, b = spec[1], spec[2]
        return f'=IF(A{r}="","",IF(AND({num(a)},{num(b)}),IF({b}{r}=0,"—",{a}{r}/{b}{r}),"—"))'
    if t == "diff":
        a, b = spec[1], spec[2]
        return f'=IF(A{r}="","",IF(AND({num(a)},{num(b)}),{a}{r}-{b}{r},"—"))'
    if t == "chg":
        a, b = spec[1], spec[2]
        return f'=IF(A{r}="","",IF(AND({num(a)},{num(b)}),IF({b}{r}=0,"—",{a}{r}/{b}{r}-1),"—"))'
    if t == "down2":
        return f'=IF(A{r}="","",IF(AND(ISNUMBER(N{r}),ISNUMBER(O{r}),ISNUMBER(P{r})),IF(AND(N{r}<O{r},O{r}<P{r}),"●",""),"—"))'
    metric, per = spec
    return f'=IF(A{r}="","",{pull(f"A{r}", metric, per)})'


for r in range(R0, U_LAST + 1):
    for i, (h, w, spec, nf) in enumerate(aspec, start=1):
        c = ws.cell(row=r, column=i)
        if spec is not None:
            c.value = afml(spec, r)
            c.fill = FILL_CALC
        c.font = F_BODY
        c.border = BORDER
        if nf:
            c.number_format = nf
ws.auto_filter.ref = f"A4:{get_column_letter(len(aspec))}{U_LAST}"
dv_list(ws, "L_DECLINE", f"AE{R0}:AE{U_LAST}")
ws.conditional_formatting.add(f"F{R0}:AD{U_LAST}", FormulaRule(
    formula=[f'OR(F{R0}="未調査",F{R0}="非開示")'], fill=FILL_WARN))
ws.conditional_formatting.add(f"F{R0}:AD{U_LAST}", FormulaRule(
    formula=[f'OR(F{R0}="不一致",F{R0}="重複")'], fill=FILL_BAD))
assert get_column_letter([h for h, *_ in aspec].index("有価証券評価差額（合計）") + 1) == "AA"
assert get_column_letter([h for h, *_ in aspec].index("外国証券 減少要因（入力・不明なら断定しない）") + 1) == "AE"
FCOL = {h: get_column_letter(i) for i, (h, *_) in enumerate(aspec, start=1)}


def frng(h):
    return col_range(FBA, FCOL[h], R0, U_LAST)


# ------------------------------------------------------------------ 6 Investment Policy Evidence
evid = read_csv("evidence.csv")
ws = S[EVI]
ecols = [("EvidenceID", 11, "key"), ("金融機関コード", 10, "in"), ("金融機関名", 18, "calc"), ("業態", 10, "calc"),
         ("テーマ", 18, "in"), ("関連スコア項目", 8, "in"), ("方向性", 8, "in"), ("判定案", 8, "in"),
         ("原文の要旨・短い引用（事実）", 44, "in"), ("原文の所在（章・表題）", 18, "in"), ("SourceID", 11, "in"),
         ("資料名", 24, "calc"), ("ページ", 7, "in"), ("資料公表日", 11, "calc"), ("出典URL", 26, "calc"),
         ("確認日", 11, "in"), ("ステータス", 13, "in"), ("分析者の解釈（推論・事実と分ける）", 36, "in")]
E_LAST = R0 + max(EVID_ROWS, len(evid) + 500) - 1
fill_rows(ws, ecols, R0, E_LAST, values=[{h: e.get(h, "") for h, _, _ in ecols} for e in evid],
          text_cols=("EvidenceID", "金融機関コード", "SourceID", "ページ"),
          fmt={"資料公表日": "yyyy-mm-dd", "確認日": "yyyy-mm-dd"},
          formulas={"金融機関名": lambda r: lookup(f"B{r}", "U_NAME", "U_CODE", "母集団外"),
                    "業態": lambda r: lookup(f"B{r}", "U_GYOTAI", "U_CODE", "母集団外"),
                    "資料名": lambda r: lookup(f"K{r}", "S_DOC", "S_ID", "台帳未登録"),
                    "資料公表日": lambda r: lookup(f"K{r}", "S_PUB", "S_ID", "台帳未登録"),
                    "出典URL": lambda r: lookup(f"K{r}", "S_URL", "S_ID", "台帳未登録")})
sheet_frame(ws, "6 Investment Policy Evidence — 運用方針の記述（1 行 = 1 記述）",
            "I 列は資料に書かれていること（事実）のみ。解釈は R 列。Tier 3 のみの記述はステータス SECONDARY とし、スコアの『明確』根拠にしない", ecols)
for col, nm in (("E", "L_THEME"), ("F", "L_SCOREITEM"), ("G", "L_DIR"), ("H", "L_JUDGE"), ("Q", "L_STATUS")):
    dv_list(ws, nm, f"{col}{R0}:{col}{E_LAST}")
name(wb, "E_CODE", col_range(EVI, "B", R0, E_LAST))
name(wb, "E_ITEM", col_range(EVI, "F", R0, E_LAST))

# ------------------------------------------------------------------ 7 Target Scoring
ws = S[SCO]
sc = [("金融機関コード", 10, lambda r: f"=IF('{UNI}'!A{r}=\"\",\"\",'{UNI}'!A{r})", "@"),
      ("金融機関名", 18, lambda r: f"=IF(A{r}=\"\",\"\",'{UNI}'!C{r})", None),
      ("業態", 10, lambda r: f"=IF(A{r}=\"\",\"\",'{UNI}'!E{r})", None),
      ("ピアグループ", 13, lambda r: f"=IF(A{r}=\"\",\"\",'{UNI}'!G{r})", None)]
QUANT_SRC = {"A1": ("外国証券（当期）", "hi"), "A2": ("外国証券／有価証券", "hi"), "C1": ("有価証券", "hi"),
             "C3a": ("預貸率", "lo"), "C3b": ("評価差額／有価証券", "hi")}
SCOL = {}


def add(h, w, f, nf=None):
    sc.append((h, w, f, nf))
    SCOL[h] = get_column_letter(len(sc))


def pct_formula(vcol, peer_col="D", direction="hi"):
    def f(r):
        v = f"{vcol}{r}"
        vr = f"${vcol}${R0}:${vcol}${U_LAST}"
        pr = f"${peer_col}${R0}:${peer_col}${U_LAST}"
        n = f'COUNTIFS({pr},{peer_col}{r},{vr},">-1E+300")'
        cmp = '"<"' if direction == "hi" else '">"'
        return (f'=IF(A{r}="","",IF(NOT(ISNUMBER({v})),"保留",IF({n}<CFG_MIN_PEER,"保留",'
                f'COUNTIFS({pr},{peer_col}{r},{vr},{cmp}&{v})/({n}-1))))')
    return f


def qual_score(item, jcol):
    return lambda r: (f'=IF(A{r}="","",IF({jcol}{r}="明確",W_{item},IF({jcol}{r}="示唆",ROUND(W_{item}*CFG_PARTIAL,1),'
                      f'IF({jcol}{r}="言及なし",0,"保留"))))')


for code, axis, label, w, kind in SCORE_ITEMS:
    if kind == "quant":
        src, d = QUANT_SRC[code]
        add(f"{code} 値", 11, (lambda s: lambda r: f"=IF(A{r}=\"\",\"\",'{FBA}'!{FCOL[s]}{r})")(src),
            "0.0%" if "／" in src else "#,##0")
        add(f"{code} ピア内pct", 8, pct_formula(SCOL[f"{code} 値"], direction=d), "0%")
        pc = SCOL[f"{code} ピア内pct"]
        add(f"{code} 点", 6, (lambda c, pc: lambda r: f'=IF(A{r}="","",IF(ISNUMBER({pc}{r}),ROUND({pc}{r}*W_{c},1),"保留"))')(code, pc), "0.0")
    elif kind == "bucket":
        add("C2 値(%)", 8, lambda r: f"=IF(A{r}=\"\",\"\",'{FBA}'!{FCOL['自己資本比率(%)']}{r})", "0.00")
        add("C2 上書き点（国際統一基準行等）", 10, None)
        v, o = SCOL["C2 値(%)"], SCOL["C2 上書き点（国際統一基準行等）"]
        add("C2 点", 6, lambda r: (f'=IF(A{r}="","",IF(ISNUMBER({o}{r}),{o}{r},IF(NOT(ISNUMBER({v}{r})),"保留",'
                                  f'IF({v}{r}>=CFG_C2_T1,W_C2,IF({v}{r}>=CFG_C2_T2,ROUND(W_C2*4/6,1),'
                                  f'IF({v}{r}>=CFG_C2_T3,ROUND(W_C2*2/6,1),0))))))'), "0.0")
    elif kind == "combo":
        for sub in ("C3a", "C3b"):
            src, d = QUANT_SRC[sub]
            add(f"{sub} 値（{src}）", 9, (lambda s: lambda r: f"=IF(A{r}=\"\",\"\",'{FBA}'!{FCOL[s]}{r})")(src), "0.0%")
            add(f"{sub} ピア内pct", 8, pct_formula(SCOL[f"{sub} 値（{src}）"], direction=d), "0%")
        a, b = SCOL["C3a ピア内pct"], SCOL["C3b ピア内pct"]
        add("C3 点", 6, lambda r: (f'=IF(A{r}="","",IF(AND(NOT(ISNUMBER({a}{r})),NOT(ISNUMBER({b}{r}))),"保留",'
                                  f'ROUND(W_C3*AVERAGE({a}{r},{b}{r}),1)))'), "0.0")
    else:
        add(f"{code} 判定", 8, None)
        if code == "A3":
            add("A3 方向性", 7, None)
        add(f"{code} 根拠EvidenceID", 11, None)
        add(f"{code} 点", 6, qual_score(code, SCOL[f"{code} 判定"]), "0.0")

pts = [SCOL[f"{c} 点"] for c, *_ in SCORE_ITEMS]
add("確認済み得点", 8, lambda r: f'=IF(A{r}="","",SUM({",".join(f"{c}{r}" for c in pts)}))', "0.0")
add("評価済み配点", 8, lambda r: '=IF(A{r}="","",{s})'.format(
    r=r, s="+".join(f"IF(ISNUMBER({c}{r}),W_{i},0)" for c, (i, *_) in zip(pts, SCORE_ITEMS))), "0")
cs, ev = SCOL["確認済み得点"], SCOL["評価済み配点"]
add("データ充足率", 8, lambda r: f'=IF(A{r}="","",{ev}{r}/100)', "0%")
add("参考スコア（評価済み項目で100点換算）", 10, lambda r: f'=IF(A{r}="","",IF({ev}{r}=0,"—",ROUND({cs}{r}/{ev}{r}*100,1)))', "0.0")
add("観点A", 6, lambda r: f'=IF(A{r}="","",SUM({",".join(SCOL[c + " 点"] + str(r) for c in ("A1", "A2", "A3", "A4"))}))', "0.0")
add("観点B", 6, lambda r: f'=IF(A{r}="","",SUM({",".join(SCOL[c + " 点"] + str(r) for c in ("B1", "B2", "B3", "B4"))}))', "0.0")
add("観点C", 6, lambda r: f'=IF(A{r}="","",SUM({",".join(SCOL[c + " 点"] + str(r) for c in ("C1", "C2", "C3"))}))', "0.0")
add("観点D", 6, lambda r: f'=IF(A{r}="","",SUM({",".join(SCOL[c + " 点"] + str(r) for c in ("D1", "D2", "D3"))}))', "0.0")
p1, p2, pc1 = SCOL["A1 ピア内pct"], SCOL["A2 ピア内pct"], SCOL["C1 ピア内pct"]
F_ = lambda h: f"'{FBA}'!{FCOL[h]}"  # noqa: E731
add("Type A 外債規模大", 7, lambda r: f'=IF(A{r}="","",IF(OR(AND(ISNUMBER({p1}{r}),{p1}{r}>=CFG_TYPE_TOP),AND(ISNUMBER({p2}{r}),{p2}{r}>=CFG_TYPE_TOP)),"●",""))')
add("Type B 外債縮小", 7, lambda r: (f'=IF(A{r}="","",IF(OR({F_("2期連続減少")}{r}="●",AND(ISNUMBER({F_("外国証券 前期比（%）")}{r}),'
                                    f'{F_("外国証券 前期比（%）")}{r}<=-CFG_TYPE_CHG)),"●",""))'))
add("Type C 外債積極化", 7, lambda r: (f'=IF(A{r}="","",IF(OR(AND(ISNUMBER({F_("外国証券 前期比（%）")}{r}),'
                                      f'{F_("外国証券 前期比（%）")}{r}>=CFG_TYPE_CHG),AND({SCOL["A3 方向性"]}{r}="拡大",{SCOL["A3 判定"]}{r}="明確")),"●",""))'))
add("Type D 有価証券大・外債比率低", 8, lambda r: (f'=IF(A{r}="","",IF(AND(ISNUMBER({pc1}{r}),ISNUMBER({p2}{r})),'
                                               f'IF(AND({pc1}{r}>=CFG_TYPE_TOP,{p2}{r}<CFG_TYPE_LOW),"●",""),""))'))
add("Type E ファンド活用", 7, lambda r: f'=IF(A{r}="","",IF(OR({SCOL["D1 判定"]}{r}="明確",{SCOL["D1 判定"]}{r}="示唆"),"●",""))')
add("算出メモ（分析者）", 30, None)

scols_ = [(h, w, "calc" if f else "in") for h, w, f, _ in sc]
sheet_frame(ws, "7 Target Scoring — スコア（行は Universe と同じ並び）",
            "定量はピア内パーセンタイル×配点、定性は 明確=配点／示唆=配点×比率／言及なし=0／空欄・評価保留=算入しない。"
            "参考スコア＝確認済み得点÷評価済み配点×100。財務健全性(C)は購入意欲を示さない。配点は Config", scols_)
for r in range(R0, U_LAST + 1):
    for i, (h, w, f, nf) in enumerate(sc, start=1):
        c = ws.cell(row=r, column=i)
        if f:
            c.value = f(r)
            c.fill = FILL_CALC
        c.font = F_BODY
        c.border = BORDER
        if nf:
            c.number_format = nf
ws.auto_filter.ref = f"A4:{get_column_letter(len(sc))}{U_LAST}"
for code, *_ in SCORE_ITEMS:
    if f"{code} 判定" in SCOL:
        L_ = SCOL[f"{code} 判定"]
        dv_list(ws, "L_JUDGE", f"{L_}{R0}:{L_}{U_LAST}")
dv_list(ws, "L_DIR", f"{SCOL['A3 方向性']}{R0}:{SCOL['A3 方向性']}{U_LAST}")
for code, *_ in SCORE_ITEMS:
    L_ = SCOL[f"{code} 点"]
    ws.conditional_formatting.add(f"{L_}{R0}:{L_}{U_LAST}", CellIsRule(operator="equal", formula=['"保留"'], fill=FILL_WARN))
for h in ("確認済み得点", "評価済み配点", "データ充足率", "参考スコア（評価済み項目で100点換算）"):
    name(wb, "SC_" + {"確認済み得点": "PTS", "評価済み配点": "EVAL", "データ充足率": "COVER",
                      "参考スコア（評価済み項目で100点換算）": "SCORE"}[h], col_range(SCO, SCOL[h], R0, U_LAST))
name(wb, "SC_PEER", col_range(SCO, "D", R0, U_LAST))

# ------------------------------------------------------------------ 8 Priority Ranking
ws = S[RNK]
SCR = lambda h: f"'{SCO}'!{SCOL[h]}"  # noqa: E731
rk = [("金融機関コード", 10, lambda r: f"=IF('{UNI}'!A{r}=\"\",\"\",'{UNI}'!A{r})", "@"),
      ("金融機関名", 18, lambda r: f"=IF(A{r}=\"\",\"\",'{UNI}'!C{r})", None),
      ("業態", 10, lambda r: f"=IF(A{r}=\"\",\"\",'{UNI}'!E{r})", None),
      ("ピアグループ", 13, lambda r: f"=IF(A{r}=\"\",\"\",'{UNI}'!G{r})", None),
      ("地域ブロック", 11, lambda r: f"=IF(A{r}=\"\",\"\",'{UNI}'!I{r})", None),
      ("参考スコア", 8, lambda r: f"=IF(A{r}=\"\",\"\",{SCR('参考スコア（評価済み項目で100点換算）')}{r})", "0.0"),
      ("データ充足率", 8, lambda r: f"=IF(A{r}=\"\",\"\",{SCR('データ充足率')}{r})", "0%"),
      ("確認済み得点", 8, lambda r: f"=IF(A{r}=\"\",\"\",{SCR('確認済み得点')}{r})", "0.0"),
      ("順位対象", 9, lambda r: (f'=IF(A{r}="","",IF(AND(ISNUMBER(F{r}),G{r}>=CFG_MIN_COVER),'
                                 f'IF(C{r}="信金中央金庫","別枠","対象"),"データ不足"))'), None),
      ("ピア内順位", 7, lambda r: (f'=IF(I{r}<>"対象","",COUNTIFS($D${R0}:$D${U_LAST},D{r},$I${R0}:$I${U_LAST},"対象",'
                                  f'$F${R0}:$F${U_LAST},">"&F{r})+1)'), "0"),
      ("ピア内件数", 7, lambda r: f'=IF(I{r}<>"対象","",COUNTIFS($D${R0}:$D${U_LAST},D{r},$I${R0}:$I${U_LAST},"対象"))', "0"),
      ("全国順位（正規化スコア）", 8, lambda r: f'=IF(I{r}<>"対象","",COUNTIFS($I${R0}:$I${U_LAST},"対象",$F${R0}:$F${U_LAST},">"&F{r})+1)', "0"),
      ("外国証券（当期・百万円・参考）", 12, lambda r: f"=IF(A{r}=\"\",\"\",'{FBA}'!{FCOL['外国証券（当期）']}{r})", "#,##0"),
      ("外国証券／有価証券（参考）", 9, lambda r: f"=IF(A{r}=\"\",\"\",'{FBA}'!{FCOL['外国証券／有価証券']}{r})", "0.0%"),
      ("該当Type", 10, lambda r: '=IF(A{r}="","",{s})'.format(r=r, s="&".join(
          f'IF({SCR(t)}{r}="●","{t[5]}","")' for t in ("Type A 外債規模大", "Type B 外債縮小", "Type C 外債積極化",
                                                         "Type D 有価証券大・外債比率低", "Type E ファンド活用"))), None),
      ("根拠Evidence件数", 8, lambda r: f'=IF(A{r}="","",COUNTIF(E_CODE,A{r}))', "0"),
      ("候補区分（手入力）", 16, None, None),
      ("選定理由（事実ベース）", 30, None, None),
      ("適合性懸念・除外理由", 30, None, None),
      ("選定者", 9, None, None), ("選定日", 10, None, "yyyy-mm-dd")]
rcols = [(h, w, "calc" if f else "in") for h, w, f, _ in rk]
sheet_frame(ws, "8 Priority Ranking — 優先順位（ピア内／全国を分けて表示）",
            "全国順位は参考スコア（評価済み項目で正規化）で計算し、絶対額は参考列。充足率が Config 未満は順位を付けない。"
            "候補区分は順位を参考に人が決める（件数を無理に満たさない）。並べ替えはオートフィルタで", rcols)
for r in range(R0, U_LAST + 1):
    for i, (h, w, f, nf) in enumerate(rk, start=1):
        c = ws.cell(row=r, column=i)
        if f:
            c.value = f(r)
            c.fill = FILL_CALC
        c.font = F_BODY
        c.border = BORDER
        if nf:
            c.number_format = nf
ws.auto_filter.ref = f"A4:{get_column_letter(len(rk))}{U_LAST}"
dv_list(ws, "L_CAND", f"Q{R0}:Q{U_LAST}")
name(wb, "R_CODE", col_range(RNK, "A", R0, U_LAST))
name(wb, "R_PEERRANK", col_range(RNK, "J", R0, U_LAST))
name(wb, "R_NATRANK", col_range(RNK, "L", R0, U_LAST))
name(wb, "R_CAND", col_range(RNK, "Q", R0, U_LAST))
name(wb, "R_TYPE", col_range(RNK, "O", R0, U_LAST))

# ------------------------------------------------------------------ 9 Bank Profiles
ws = S[PRO]
pcols = [("ProfileRowID", 11, "key"), ("金融機関コード", 10, "in"), ("金融機関名", 18, "calc"), ("候補区分", 14, "calc"),
         ("項目", 24, "in"), ("記述区分", 11, "in"), ("記述", 60, "in"), ("根拠ID（Evidence/Record/Source）", 18, "in"),
         ("確信度", 7, "in"), ("作成者", 9, "in"), ("作成日", 10, "in"), ("チェック", 16, "calc")]
P_LAST = R0 + PROFILE_ROWS - 1
fill_rows(ws, pcols, R0, P_LAST, text_cols=("ProfileRowID", "金融機関コード"), fmt={"作成日": "yyyy-mm-dd"},
          formulas={"金融機関名": lambda r: lookup(f"B{r}", "U_NAME", "U_CODE", "母集団外"),
                    "候補区分": lambda r: lookup(f"B{r}", "R_CAND", "R_CODE"),
                    "チェック": lambda r: (f'=IF(B{r}="","",IF(AND(F{r}="事実",H{r}=""),"事実には根拠ID必須",'
                                         f'IF(AND(F{r}<>"事実",I{r}=""),"確信度未記入","OK")))')})
sheet_frame(ws, "9 Bank Profiles — 金融機関別プロファイル（1 行 = 1 記述）",
            "8 項目それぞれに 事実／分析上の推論／営業仮説 を区別して記入。事実は根拠 ID 必須。公開情報から購入意欲・適合性を断定しない", pcols)
dv_list(ws, "L_SECTION", f"E{R0}:E{P_LAST}")
dv_list(ws, "L_KIND", f"F{R0}:F{P_LAST}")
dv_list(ws, "L_CONF", f"I{R0}:I{P_LAST}")
ws.conditional_formatting.add(f"L{R0}:L{P_LAST}", FormulaRule(formula=[f'AND(L{R0}<>"",L{R0}<>"OK")'], fill=FILL_BAD))

# ------------------------------------------------------------------ 10 Sales Coverage Input (internal)
ws = S[COV]
ccols = [("金融機関コード", 10, "in"), ("金融機関名", 18, "calc"), ("業態", 10, "calc"), ("公開情報 ピア内順位", 8, "calc"),
         ("公開情報 全国順位", 8, "calc"), ("公開情報 該当Type", 9, "calc"), ("営業担当者", 10, "in"),
         ("既存カバレッジ状況", 12, "in"), ("既存ファンド購入実績", 20, "in"), ("過去の商談履歴", 30, "in"),
         ("ファンド投資の可否", 10, "in"), ("現在の外債投資ニーズ", 24, "in"), ("次回面談予定", 11, "in"),
         ("営業担当者の評価", 8, "in"), ("情報源（面談／電話／メール等）", 14, "in"), ("機密区分", 9, "in"),
         ("入力者", 9, "in"), ("更新日", 10, "in")]
C_LAST = R0 + N_INST - 1
fill_rows(ws, ccols, R0, C_LAST, text_cols=("金融機関コード",), fmt={"次回面談予定": "yyyy-mm-dd", "更新日": "yyyy-mm-dd"},
          formulas={"金融機関名": lambda r: lookup(f"A{r}", "U_NAME", "U_CODE", "母集団外"),
                    "業態": lambda r: lookup(f"A{r}", "U_GYOTAI", "U_CODE", "母集団外"),
                    "公開情報 ピア内順位": lambda r: lookup(f"A{r}", "R_PEERRANK", "R_CODE"),
                    "公開情報 全国順位": lambda r: lookup(f"A{r}", "R_NATRANK", "R_CODE"),
                    "公開情報 該当Type": lambda r: lookup(f"A{r}", "R_TYPE", "R_CODE")})
sheet_frame(ws, "10 Sales Coverage Input — 【社内情報・機密】営業カバレッジ",
            "【社外秘】社内情報はこのシートと 10b・11 のみに記入する。公開情報シート（1〜9）からは参照しない。リポジトリにはコミットしない",
            ccols, internal=True)
for col, nm in (("H", "L_COVER"), ("K", "L_YN3"), ("N", "L_RATING"), ("P", "L_CCLASS")):
    dv_list(ws, nm, f"{col}{R0}:{col}{C_LAST}")

# ------------------------------------------------------------------ 10b M1 Flow Input (internal)
ws = S[M1]
mcols = [("レコードID", 10, "in"), ("データ名", 20, "in"), ("提供元", 8, "in"), ("対象期間（月）", 10, "in"),
         ("粒度（月次／業態別／地域別／個社別 等）", 18, "in"), ("対象範囲の説明", 30, "in"), ("個社特定可否", 9, "in"),
         ("金融機関コード（個社特定可の場合のみ）", 12, "in"), ("値", 12, "in"), ("単位", 8, "in"),
         ("個社分析への利用可否", 22, "calc"), ("機密区分", 9, "in"), ("入力者", 9, "in"), ("更新日", 10, "in"), ("備考", 30, "in")]
M_LAST = R0 + 1000 - 1
fill_rows(ws, mcols, R0, M_LAST, text_cols=("金融機関コード（個社特定可の場合のみ）", "対象期間（月）"), fmt={"更新日": "yyyy-mm-dd"},
          formulas={"個社分析への利用可否": lambda r: (f'=IF(B{r}="","",IF(G{r}<>"可","個社の投資行動の根拠に使用不可",'
                                                   f'IF(H{r}="","コード未記入","個社分析可（範囲を要確認）")))')})
sheet_frame(ws, "10b M1 Flow Input — 【社内情報・機密】M1 月次ファンドフロー",
            "【社外秘】まずデータの対象範囲と粒度を記録。個別金融機関の売買が特定できないデータを特定行の投資行動の証拠にしない", mcols, internal=True)
dv_list(ws, "L_IDENT", f"G{R0}:G{M_LAST}")
dv_list(ws, "L_CCLASS", f"L{R0}:L{M_LAST}")

# ------------------------------------------------------------------ 11 Sales Action Plan (internal)
ws = S[ACT]
acols2 = [("ActionID", 9, "in"), ("金融機関コード", 10, "in"), ("金融機関名", 18, "calc"), ("候補区分", 14, "calc"),
          ("提案仮説（Profile の ProfileRowID）", 14, "in"), ("提案仮説の要旨", 36, "in"), ("顧客に確認すべき質問", 36, "in"),
          ("次アクション", 24, "in"), ("担当", 9, "in"), ("期限", 10, "in"), ("状態", 9, "in"), ("結果メモ", 30, "in"),
          ("機密区分", 9, "in"), ("更新日", 10, "in")]
A_LAST = R0 + 1000 - 1
fill_rows(ws, acols2, R0, A_LAST, text_cols=("ActionID", "金融機関コード"), fmt={"期限": "yyyy-mm-dd", "更新日": "yyyy-mm-dd"},
          formulas={"金融機関名": lambda r: lookup(f"B{r}", "U_NAME", "U_CODE", "母集団外"),
                    "候補区分": lambda r: lookup(f"B{r}", "R_CAND", "R_CODE")})
sheet_frame(ws, "11 Sales Action Plan — 【社内情報・機密】営業アクション", "【社外秘】", acols2, internal=True)
dv_list(ws, "L_ASTATE", f"K{R0}:K{A_LAST}")
dv_list(ws, "L_CCLASS", f"M{R0}:M{A_LAST}")

# ------------------------------------------------------------------ 12 Data Quality Dashboard
ws = S[DQ]
ws.sheet_view.showGridLines = False
ws["A1"] = "12 Data Quality Dashboard — データ品質とバッチ進捗"
ws["A1"].font = F_TITLE
ws["A2"] = S[UNI]["A2"].value
ws["A2"].font = F_SMALL
hdr = ["業態", "金融機関数", "母集団確認済以降", "定量完了以降", "定性完了以降"] + \
      [f"数値行 {s}" for s in STATUS] + ["平均データ充足率", "順位対象数", "データ不足数"]
for i, h in enumerate(hdr, start=1):
    c = ws.cell(row=4, column=i, value=h)
    c.font = F_HEAD
    c.fill = FILL_HEAD
    c.alignment = HEAD_AL
    ws.column_dimensions[get_column_letter(i)].width = 12
ws.column_dimensions["A"].width = 20
ws.row_dimensions[4].height = 36
groups = GYOTAI[:5]
for j, g in enumerate(groups + ["合計"], start=R0):
    ws.cell(row=j, column=1, value=g).font = F_BODY
    crit = f'"{g}"' if g != "合計" else '"<>"'
    stat_after = lambda xs: "+".join(f'COUNTIFS(U_GYOTAI,{crit},U_USTAT,"{x}")' for x in xs)  # noqa: E731
    vals = [f"=COUNTIFS(U_GYOTAI,{crit},U_CODE,\"<>\")",
            "=" + stat_after(UNIV_STATUS[1:]), "=" + stat_after(UNIV_STATUS[3:]), "=" + stat_after(UNIV_STATUS[4:])]
    for s in STATUS:
        vals.append(f'=COUNTIFS(FD_GYO,{crit},FD_STAT,"{s}")+COUNTIFS(SP_GYO,{crit},SP_STAT,"{s}")')
    vals += [f'=IFERROR(AVERAGEIFS(SC_COVER,\'{SCO}\'!$C${R0}:$C${U_LAST},{crit}),"—")',
             f"=COUNTIFS('{RNK}'!$C${R0}:$C${U_LAST},{crit},'{RNK}'!$I${R0}:$I${U_LAST},\"対象\")",
             f"=COUNTIFS('{RNK}'!$C${R0}:$C${U_LAST},{crit},'{RNK}'!$I${R0}:$I${U_LAST},\"データ不足\")"]
    for i, v in enumerate(vals, start=2):
        c = ws.cell(row=j, column=i, value=v)
        c.font = F_BODY
        c.border = BORDER
        c.number_format = "0%" if i == len(hdr) - 2 else "#,##0"
r = R0 + len(groups) + 3
ws.cell(row=r, column=1, value="整合チェック（0 であるべき件数）").font = F_TITLE
checks = [("Universe の重複コード", f"=COUNTIF('{UNI}'!$U${R0}:$U${U_LAST},\"重複\")"),
          ("Financial Data の整合エラー", f"=COUNTIFS('{FD}'!$U${R0}:$U${FACT['FD']},\"<>OK\",'{FD}'!$B${R0}:$B${FACT['FD']},\"<>\")"),
          ("Securities Portfolio の整合エラー", f"=COUNTIFS('{SP}'!$U${R0}:$U${FACT['SP']},\"<>OK\",'{SP}'!$B${R0}:$B${FACT['SP']},\"<>\")"),
          ("Profile の根拠・確信度未記入", f"=COUNTIFS('{PRO}'!$L${R0}:$L${P_LAST},\"<>OK\",'{PRO}'!$B${R0}:$B${P_LAST},\"<>\")"),
          ("配点合計が100でない", "=IF(SUM(W_A1,W_A2,W_A3,W_A4,W_B1,W_B2,W_B3,W_B4,W_C1,W_C2,W_C3,W_D1,W_D2,W_D3)=100,0,1)")]
for k, (label, fml) in enumerate(checks, start=r + 1):
    ws.cell(row=k, column=1, value=label).font = F_BODY
    c = ws.cell(row=k, column=2, value=fml)
    c.border = BORDER
    ws.conditional_formatting.add(f"B{k}", CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_BAD))
    ws.conditional_formatting.add(f"B{k}", CellIsRule(operator="equal", formula=["0"], fill=FILL_GOOD))
r = r + len(checks) + 3
ws.cell(row=r, column=1, value="バッチ記録（各バッチ完了時に記入。資料を開いていない先は完了に数えない）").font = F_TITLE
bh = ["バッチ番号", "対象（業態・地域）", "調査対象数", "調査完了数", "未調査数", "データ取得成功数", "データ不足数",
      "次回調査対象", "実施日", "制約・メモ"]
for i, h in enumerate(bh, start=1):
    c = ws.cell(row=r + 1, column=i, value=h)
    c.font = F_HEAD
    c.fill = FILL_HEAD
    c.alignment = HEAD_AL
batches = read_csv("batch_log.csv")
for k in range(60):
    rr = r + 2 + k
    b = batches[k] if k < len(batches) else {}
    for i, h in enumerate(bh, start=1):
        c = ws.cell(row=rr, column=i, value=b.get(h) or None)
        c.border = BORDER
        c.font = F_BODY
    ws.cell(row=rr, column=5, value=f'=IF(C{rr}="","",C{rr}-D{rr})').fill = FILL_CALC

# ------------------------------------------------------------------ README
ws = S["README"]
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 26
ws.column_dimensions["B"].width = 110
lines = [
    ("全国金融機関ターゲット分析 — 米国債ラダー × Dynamic FX Hedging", None),
    ("目的", "公開情報に基づき営業仮説を検証するためのデータベース。購入意欲・商品適合性を判定するものではない。"),
    ("状態", "テンプレート（データ未投入）" if TEMPLATE else "data/*.csv から生成"),
    ("基準日", S[UNI]["A2"].value),
    ("正本データ", "公開情報は data/*.csv（universe / sources / datapoints / evidence / batch_log）。build_workbook.py で再生成する。"),
    ("入力の流れ", "① Universe に母集団 → ② Source Register に資料 → ③④ 縦持ちに数値（出典・ページ・ステータス必須）→ "
                "⑥ 方針記述 → ⑦ 定性項目の判定を入力 → ⑧ 候補区分を人が決定 → ⑨ プロファイル"),
    ("灰色セル", "数式。編集しない。白セルが入力欄。"),
    ("ステータス", "VERIFIED 一次資料で確認 / SECONDARY 二次資料のみ / NOT DISCLOSED 調査したが非開示（SourceID 必須）/ "
                "NOT YET CHECKED 未調査 / CONFLICT 資料間で不一致 / ESTIMATED 明示した計算式による推計"),
    ("抽出ルール", "分析シートは Config の比較基準（単体等）・期に一致し、有効（VERIFIED/SECONDARY/ESTIMATED）な値が"
                "ちょうど 1 件のときだけ数値を使う。0 件なら 未調査／非開示／不一致 を文字で表示。"),
    ("外国証券の注意", "外国証券には外国株式・外国投信等が含まれうる。米国債残高とはみなさない。"
                   "残高減少の要因（売却・償還・換算・時価）は資料で区別できない限り断定しない。"),
    ("スコア", "A 外債関連性35／B 投資機会30／C 財務余力20／D 商品親和性15（Config で変更可）。"
             "評価保留は 0 点にせず、参考スコア・データ充足率・確認済み得点を併記。"),
    ("財務健全性", "C の得点は運用上の制約が少ない可能性を示すだけで、投資余力・購入意欲とはみなさない。"),
    ("社内情報", "10・10b・11 は社外秘。公開情報シートとは分離し、公開情報シートから参照しない。"
              "M1 データは個社特定可否を必ず記録し、特定不可のデータを特定行の投資行動の根拠にしない。"),
    ("関数", "Excel 2007 以降の標準関数のみ（COUNTIFS / SUMIFS / AVERAGEIFS / INDEX / MATCH / IFERROR）。"),
]
for k, (a, b) in enumerate(lines, start=1):
    ca = ws.cell(row=k, column=1, value=a)
    ca.font = F_TITLE if b is None else Font(name=JP, size=10, bold=True, color=INK)
    if b is not None:
        cb = ws.cell(row=k, column=2, value=b)
        cb.font = F_BODY
        cb.alignment = WRAP

for s in (CFG, LST, MM):
    S[s].sheet_properties.tabColor = "9AA5B1"
for s in (COV, M1, ACT):
    S[s].sheet_properties.tabColor = "7A1F1F"

wb.calculation.fullCalcOnLoad = True
OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
print(f"wrote {OUT}  (universe={len(universe)}, sources={len(sources)}, datapoints={len(points)}, evidence={len(evid)})")
