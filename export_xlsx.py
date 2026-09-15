#!/usr/bin/env python3
"""data/products.json から Excel ブック（dist/*.xlsx）を生成する。

集計シート（会社×類型マトリクス・凡例の件数）はすべて数式で書く。値をハードコード
しないので、商品一覧の行を足し引きすれば Excel 側で再計算される。
"""
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation  # noqa: F401  (future use)

ROOT = Path(__file__).parent
DB = json.loads((ROOT / "data" / "products.json").read_text(encoding="utf-8"))
OUT = ROOT / "dist" / "日本の生保_貯蓄性商品ラインナップ横断表.xlsx"

CATS = DB["taxonomies"]["categories"]
CHANS = {c["id"]: c["label"] for c in DB["taxonomies"]["channels"]}
PURPS = {p["id"]: p["label"] for p in DB["taxonomies"]["purposes"]}
CAT = {c["id"]: c for c in CATS}
P = DB["products"]

INSURERS = list(dict.fromkeys(p["insurer"] for p in P))

# ---------------------------------------------------------------- styling
JP = "Yu Gothic"          # Office の日本語既定フォント（Windows / macOS 双方にある）
MONO = "Consolas"

INK = "1F2933"
HEAD_BG = "1C3F6E"
SUB_BG = "E2EAF4"
FLAG_BG = "F6E4EC"
WARN_BG = "F7EBD8"
ZEBRA = "F5F7F9"

F_TITLE = Font(name=JP, size=16, bold=True, color=INK)
F_H2 = Font(name=JP, size=11, bold=True, color=INK)
F_HEAD = Font(name=JP, size=10, bold=True, color="FFFFFF")
F_BODY = Font(name=JP, size=10, color=INK)
F_SMALL = Font(name=JP, size=9, color="52606D")
F_MONO = Font(name=MONO, size=9, color="52606D")
F_LINK = Font(name=JP, size=9, color="1C3F6E", underline="single")

FILL_HEAD = PatternFill("solid", fgColor=HEAD_BG)
FILL_SUB = PatternFill("solid", fgColor=SUB_BG)
FILL_FLAG = PatternFill("solid", fgColor=FLAG_BG)
FILL_WARN = PatternFill("solid", fgColor=WARN_BG)
FILL_ZEBRA = PatternFill("solid", fgColor=ZEBRA)

_thin = Side(style="thin", color="D5DAE1")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

TOP_WRAP = Alignment(vertical="top", wrap_text=True)
TOP = Alignment(vertical="top")
CENTER = Alignment(horizontal="center", vertical="center")
HEAD_AL = Alignment(horizontal="left", vertical="center", wrap_text=True)


def put(ws, row, col, value, font=F_BODY, align=TOP_WRAP, fill=None, border=True):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font
    c.alignment = align
    if fill:
        c.fill = fill
    if border:
        c.border = BORDER
    return c


def header_row(ws, row, labels, widths=None):
    for i, label in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=label)
        c.font = F_HEAD
        c.fill = FILL_HEAD
        c.alignment = HEAD_AL
        c.border = BORDER
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 30


def join(xs, sep=" / "):
    return sep.join(x for x in xs if x)


wb = Workbook()

# ============================================================ 1. 凡例・注意
ws = wb.active
ws.title = "凡例・注意"
ws.sheet_view.showGridLines = False
for col, w in zip("ABCD", (22, 46, 46, 30)):
    ws.column_dimensions[col].width = w

r = 1
put(ws, r, 1, DB["meta"]["title"], F_TITLE, TOP, border=False)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws.row_dimensions[r].height = 26
r += 1
put(ws, r, 1, DB["meta"]["scope"], F_SMALL, TOP_WRAP, border=False)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws.row_dimensions[r].height = 30
r += 2

put(ws, r, 1, "編纂日", F_H2, TOP, border=False)
put(ws, r, 2, DB["meta"]["compiledOn"], F_BODY, TOP, border=False)
r += 1

# 件数は数式。商品一覧の行を増減すれば追随する。
LAST = len(P) + 1  # 商品一覧のデータ最終行（1 行目はヘッダ）
counts = [
    ("収録商品数", f'=COUNTA(商品一覧!$A$2:$A${LAST})'),
    ("引受会社数（重複除く）",
     f'=SUMPRODUCT((商品一覧!$B$2:$B${LAST}<>"")/COUNTIF(商品一覧!$B$2:$B${LAST},商品一覧!$B$2:$B${LAST}&""))'),
    ("裏側の証券会社を特定した商品数", f'=COUNTA(商品一覧!$Q$2:$Q${LAST})'),
    ("指数連動商品数", f'=COUNTIF(商品一覧!$K$2:$K${LAST},"●")'),
    ("確度「確認済」の商品数", f'=COUNTIF(商品一覧!$X$2:$X${LAST},"確認済")'),
    ("確度「要確認」の商品数", f'=COUNTIF(商品一覧!$X$2:$X${LAST},"要確認")'),
]
put(ws, r, 1, "収録サマリー（数式・商品一覧シートを参照）", F_H2, TOP, border=False)
r += 1
for label, formula in counts:
    put(ws, r, 1, label, F_BODY, TOP)
    c = put(ws, r, 2, formula, F_BODY, Alignment(horizontal="left", vertical="top"))
    c.number_format = "#,##0"
    r += 1
r += 1

put(ws, r, 1, "ご指定の切り口と列の対応", F_H2, TOP, border=False)
r += 1
mapping = [
    ("販売チャネル・取り扱い会社", "商品一覧 L 列（販売チャネル）／M 列（取扱会社：具体的な銀行・証券会社名）"),
    ("商品類型", "商品一覧 G 列。複数該当は「/」区切り。正規化した 1 行 1 組の表は『商品×類型』シート"),
    ("商品概要（2-3 行）", "商品一覧 N 列"),
    ("各社でのペットネーム", "商品一覧 D 列（主名称）／E 列（別名・世代違い・販売会社別の販売名称）"),
    ("対象年齢", "商品一覧 O 列（目的・ニーズ）／P 列（契約可能年齢）"),
    ("裏側の証券会社", "商品一覧 Q〜T 列。特定できた商品のみを抜き出した表は『裏側の証券会社』シート"),
    ("最新の料率", "商品一覧 U〜W 列（水準・改定サイクル・公式確認先 URL）"),
]
for k, v in mapping:
    put(ws, r, 1, k, F_BODY, TOP_WRAP, FILL_SUB)
    c = put(ws, r, 2, v, F_BODY, TOP_WRAP)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws.row_dimensions[r].height = 28
    r += 1
r += 1

put(ws, r, 1, "料率の扱い（重要）", F_H2, TOP, border=False)
r += 1
rate_note = (
    "積立利率・予定利率は各社とも原則として毎月 2 回（1 日・16 日）改定され、通貨・契約年齢・プラン・"
    "高額契約の別で異なります。確定値を表に固定すると即座に陳腐化するため、本ブックは「水準・決まり方・"
    "公式の確認先 URL」を保持する設計です（商品一覧 U〜W 列）。個別に判明した水準のみ U 列に記載しています。\n"
    "参考となる横並び水準（2026 年 7 月 16〜31 日契約分、8 社 53 件の調査）：円建 1.95〜3.90%、米ドル建 3.34〜5.85%。"
)
put(ws, r, 1, rate_note, F_BODY, TOP_WRAP, FILL_WARN)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws.row_dimensions[r].height = 72
r += 2

put(ws, r, 1, "データの限界", F_H2, TOP, border=False)
r += 1
for i, lim in enumerate(DB["meta"]["limitations"], start=1):
    put(ws, r, 1, f"{i}.", F_BODY, TOP)
    put(ws, r, 2, lim, F_BODY, TOP_WRAP)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws.row_dimensions[r].height = 46
    r += 1
r += 1
put(ws, r, 1, "出典の方針", F_H2, TOP, border=False)
r += 1
put(ws, r, 1, DB["meta"]["sourcePolicy"], F_BODY, TOP_WRAP)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws.row_dimensions[r].height = 34

# ============================================================ 2. 商品一覧
ws = wb.create_sheet("商品一覧")
COLS = [
    ("商品ID", 20), ("引受会社", 20), ("企業グループ", 18),
    ("ペットネーム（主）", 26), ("別名・世代違い・販売会社別名称", 32), ("正式名称", 30),
    ("商品類型", 24), ("類型グループ", 12), ("払方", 10), ("通貨", 16), ("指数連動", 8),
    ("販売チャネル", 20), ("取扱会社", 30), ("商品概要", 62), ("目的・ニーズ", 22),
    ("契約年齢", 20), ("裏側の証券会社", 26), ("証券会社の役割", 34),
    ("参照指数・戦略", 38), ("証券会社に関する補足", 34),
    ("最新の料率", 30), ("料率の改定・基準", 20), ("料率確認URL", 34),
    ("確度", 8), ("出典", 46),
]
header_row(ws, 1, [c[0] for c in COLS], [c[1] for c in COLS])

for i, p in enumerate(P):
    row = i + 2
    cp = p["counterparty"]
    groups = list(dict.fromkeys(CAT[c]["group"] for c in p["categories"]))
    vals = [
        p["id"], p["insurer"], p["group"],
        p["petNames"][0],
        join(p["petNames"][1:], "、") or None,
        p["officialName"],
        join([CAT[c]["label"] for c in p["categories"]]),
        join(groups, "、"),
        p["payment"], join(p["currencies"], "、"),
        "●" if p["indexLinked"] else None,
        join([CHANS[c] for c in p["channels"]], "、"),
        join(p["distributors"], " ／ "),
        p["summary"],
        join([PURPS[c] for c in p["purposes"]], "、"),
        p["ageRange"],
        cp["name"] if cp["present"] else None,
        cp["role"] if cp["present"] else None,
        cp["index"] if cp["present"] else None,
        cp["note"],
        p["rate"]["value"], p["rate"]["asOf"], p["rate"]["url"],
        "確認済" if p["confidence"] == "confirmed" else "要確認",
        "\n".join(p["sources"]),
    ]
    for j, v in enumerate(vals, start=1):
        align = TOP_WRAP
        font = F_BODY
        if j in (1, 23, 25):
            font = F_MONO
        if j == 11:
            align = CENTER
        c = put(ws, row, j, v, font, align)
        if i % 2 == 1:
            c.fill = FILL_ZEBRA
    if cp["present"]:
        for j in (17, 18, 19):
            ws.cell(row=row, column=j).fill = FILL_FLAG
    if p["confidence"] != "confirmed":
        ws.cell(row=row, column=24).fill = FILL_WARN
    ws.cell(row=row, column=24).alignment = CENTER
    ws.row_dimensions[row].height = 74

ws.freeze_panes = "E2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{len(P) + 1}"

# ============================================================ 3. 商品×類型
ws = wb.create_sheet("商品×類型")
header_row(ws, 1, ["商品ID", "引受会社", "ペットネーム（主）", "類型グループ", "商品類型"],
           [20, 20, 26, 14, 24])
r = 2
for p in P:
    for cid in p["categories"]:
        put(ws, r, 1, p["id"], F_MONO, TOP)
        put(ws, r, 2, p["insurer"])
        put(ws, r, 3, p["petNames"][0])
        put(ws, r, 4, CAT[cid]["group"])
        put(ws, r, 5, CAT[cid]["label"])
        r += 1
PAIR_LAST = r - 1
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:E{PAIR_LAST}"

# ============================================================ 4. マトリクス
ws = wb.create_sheet("会社×類型マトリクス")
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 30
put(ws, 1, 1, "引受会社", F_HEAD, HEAD_AL, FILL_HEAD)
for j, c in enumerate(CATS, start=2):
    put(ws, 1, j, c["label"], F_HEAD, Alignment(horizontal="center", vertical="bottom", wrap_text=True), FILL_HEAD)
    ws.column_dimensions[get_column_letter(j)].width = 11
last_col = len(CATS) + 2
put(ws, 1, last_col, "計", F_HEAD, CENTER, FILL_HEAD)
ws.column_dimensions[get_column_letter(last_col)].width = 8
ws.row_dimensions[1].height = 58

for i, ins in enumerate(INSURERS):
    row = i + 2
    put(ws, row, 1, ins, F_BODY, TOP, FILL_SUB)
    for j, c in enumerate(CATS, start=2):
        col = get_column_letter(j)
        f = (f"=COUNTIFS('商品×類型'!$B$2:$B${PAIR_LAST},$A{row},"
             f"'商品×類型'!$E$2:$E${PAIR_LAST},{col}$1)")
        cell = put(ws, row, j, f, F_BODY, CENTER)
        cell.number_format = '0;;"・"'
    tot = get_column_letter(len(CATS) + 1)
    put(ws, row, last_col, f"=SUM(B{row}:{tot}{row})", F_BODY, CENTER, FILL_SUB)
    ws.row_dimensions[row].height = 18

trow = len(INSURERS) + 2
put(ws, trow, 1, "計", F_H2, TOP, FILL_SUB)
for j in range(2, last_col + 1):
    col = get_column_letter(j)
    put(ws, trow, j, f"=SUM({col}2:{col}{trow - 1})", F_H2, CENTER, FILL_SUB)
put(ws, trow + 2, 1,
    "※ 1 商品が複数類型に該当する場合は各類型に計上されるため、右端「計」の合計は収録商品数より大きくなります。",
    F_SMALL, TOP_WRAP, border=False)
ws.freeze_panes = "B2"

# ============================================================ 5. 裏側の証券会社
ws = wb.create_sheet("裏側の証券会社")
header_row(ws, 1,
           ["引受会社", "ペットネーム（主）", "商品類型", "証券会社・指数提供者",
            "役割", "参照指数・戦略", "補足", "出典"],
           [22, 26, 22, 30, 38, 44, 40, 44])
r = 2
for p in P:
    cp = p["counterparty"]
    if not cp["present"]:
        continue
    put(ws, r, 1, p["insurer"])
    put(ws, r, 2, p["petNames"][0])
    put(ws, r, 3, join([CAT[c]["label"] for c in p["categories"]]))
    put(ws, r, 4, cp["name"], Font(name=JP, size=10, bold=True, color="7A2F52"), TOP_WRAP, FILL_FLAG)
    put(ws, r, 5, cp["role"])
    put(ws, r, 6, cp["index"])
    put(ws, r, 7, cp["note"], F_SMALL)
    put(ws, r, 8, "\n".join(p["sources"]), F_MONO)
    ws.row_dimensions[r].height = 88
    r += 1
ws.freeze_panes = "A2"

# ============================================================ 6. マスタ
ws = wb.create_sheet("マスタ")
ws.sheet_view.showGridLines = False
for col, w in zip("ABCD", (24, 14, 30, 80)):
    ws.column_dimensions[col].width = w
r = 1
put(ws, r, 1, "商品類型マスタ（ご指定の 14 類型）", F_H2, TOP, border=False)
r += 1
header_row(ws, r, ["ID", "グループ", "名称", "定義・含めた範囲"], [24, 14, 30, 80])
r += 1
for c in CATS:
    put(ws, r, 1, c["id"], F_MONO, TOP)
    put(ws, r, 2, c["group"])
    put(ws, r, 3, c["label"], F_H2)
    put(ws, r, 4, c["note"], F_SMALL)
    ws.row_dimensions[r].height = 32
    r += 1
r += 1
put(ws, r, 1, "販売チャネル", F_H2, TOP, border=False)
r += 1
header_row(ws, r, ["ID", "名称"], [24, 30])
r += 1
for cid, label in CHANS.items():
    put(ws, r, 1, cid, F_MONO, TOP)
    put(ws, r, 2, label)
    r += 1
r += 1
put(ws, r, 1, "目的・ニーズ", F_H2, TOP, border=False)
r += 1
header_row(ws, r, ["ID", "名称"], [24, 30])
r += 1
for pid, label in PURPS.items():
    put(ws, r, 1, pid, F_MONO, TOP)
    put(ws, r, 2, label)
    r += 1
r += 1
put(ws, r, 1,
    "類型の割り当ては、一次資料の正式名称と仕組み記載に基づく本表側の分類です。"
    "各社の公式な商品分類とは一致しない場合があります。", F_SMALL, TOP_WRAP, border=False)

# 数式セルにキャッシュ値を持たせられない（この環境では LibreOffice が xlsx を
# ロードできず再計算を回せない）ため、開いた時点で必ず全再計算させる。
wb.calculation.fullCalcOnLoad = True

OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
print(f"wrote {OUT} ({OUT.stat().st_size / 1024:.0f} KB, {len(wb.sheetnames)} sheets: {', '.join(wb.sheetnames)})")
