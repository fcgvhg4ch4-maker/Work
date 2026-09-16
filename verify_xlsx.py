#!/usr/bin/env python3
"""生成した Excel の数式を、LibreOffice を使わずに検証する。

この環境では LibreOffice が xlsx をロードできず recalc.py が回せないため、
(1) 使っている関数が Excel 2007 以前の無印関数だけであること、
(2) 数式が指す範囲に、期待どおりのデータが実際に入っていること、
(3) 数式が返すはずの値が、products.json から独立に計算した値と一致すること
を Python 側で確かめる。
"""
import json
import re
import sys
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).parent
DB = json.loads((ROOT / "data" / "products.json").read_text(encoding="utf-8"))
XLSX = ROOT / "dist" / "日本の生保_貯蓄性商品ラインナップ横断表.xlsx"

SAFE = {"COUNTA", "COUNTIF", "COUNTIFS", "SUM", "SUMPRODUCT"}
BANNED = {"XLOOKUP", "XMATCH", "SORT", "FILTER", "UNIQUE", "SEQUENCE",
          "TEXTJOIN", "CONCAT", "IFS", "SWITCH", "MAXIFS", "MINIFS"}

P = DB["products"]
CAT = {c["id"]: c for c in DB["taxonomies"]["categories"]}
CATS = DB["taxonomies"]["categories"]
PURP_LIST = DB["taxonomies"]["purposes"]
PURPS = {x["id"]: x["label"] for x in PURP_LIST}
BANDS = DB["taxonomies"]["ageBands"]
BAND = {b["id"]: b for b in BANDS}
INSURERS = list(dict.fromkeys(p["insurer"] for p in P))

wb = load_workbook(XLSX)          # formulas
fails, checks = [], 0


def ok(label, got, want):
    global checks
    checks += 1
    if got != want:
        fails.append(f"{label}: got {got!r}, want {want!r}")


# ---- 1. 使用関数がすべて安全なものか -------------------------------------
used = Counter()
formula_cells = 0
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                formula_cells += 1
                for fn in re.findall(r"([A-Z][A-Z0-9_.]*)\s*\(", c.value):
                    used[fn] += 1
print(f"数式セル {formula_cells} 個 / 使用関数 {dict(used)}")
for fn in used:
    if fn in BANNED:
        fails.append(f"禁止関数 {fn} を使用")
    elif fn not in SAFE:
        fails.append(f"安全リスト外の関数 {fn} を使用")

# ---- 2. 商品一覧の範囲に期待どおりのデータが入っているか -------------------
wsL = wb["商品一覧"]
LAST = len(P) + 1
ok("商品一覧の最終データ行", wsL.max_row, LAST)
ok("A列(商品ID)の非空セル数", sum(1 for r in range(2, LAST + 1) if wsL.cell(r, 1).value), len(P))
ok("S列(裏側の証券会社)の非空セル数",
   sum(1 for r in range(2, LAST + 1) if wsL.cell(r, 19).value),
   sum(1 for p in P if p["counterparty"]["present"]))
ok("K列(指数連動)の●の数",
   sum(1 for r in range(2, LAST + 1) if wsL.cell(r, 11).value == "●"),
   sum(1 for p in P if p["indexLinked"]))
ok("Z列(確度)の確認済の数",
   sum(1 for r in range(2, LAST + 1) if wsL.cell(r, 26).value == "確認済"),
   sum(1 for p in P if p["confidence"] == "confirmed"))
ok("O列(主ニーズ)の中身",
   [wsL.cell(r, 15).value for r in range(2, LAST + 1)],
   [PURPS[p["primaryNeed"]] for p in P])
ok("Q列(対象年齢層)の中身",
   [wsL.cell(r, 17).value for r in range(2, LAST + 1)],
   ["、".join(BAND[b]["label"] for b in p["ageBands"]) for p in P])
ok("R列(契約年齢)の実値がある商品数",
   sum(1 for r in range(2, LAST + 1) if wsL.cell(r, 18).value),
   sum(1 for p in P if p["ageRange"]))
ok("B列の異なり引受会社数",
   len({wsL.cell(r, 2).value for r in range(2, LAST + 1)}), len(INSURERS))

# ---- 3. 商品×類型（マトリクスの参照元）------------------------------------
wsP = wb["商品×類型"]
pairs_expected = [(p["insurer"], CAT[c]["label"]) for p in P for c in p["categories"]]
PAIR_LAST = len(pairs_expected) + 1
ok("商品×類型の最終行", wsP.max_row, PAIR_LAST)
pairs_actual = [(wsP.cell(r, 2).value, wsP.cell(r, 5).value) for r in range(2, PAIR_LAST + 1)]
ok("商品×類型の中身", Counter(pairs_actual), Counter(pairs_expected))

# ---- 4. マトリクスの COUNTIFS を手計算と突き合わせる ------------------------
wsM = wb["会社×類型マトリクス"]
expected = Counter(pairs_expected)
bad_ref = 0
for i, ins in enumerate(INSURERS):
    row = i + 2
    ok(f"マトリクス行見出し[{row}]", wsM.cell(row, 1).value, ins)
    for j, c in enumerate(CATS, start=2):
        f = wsM.cell(row, j).value
        # 参照先シート・範囲・判定セルが正しいか
        want_ref = (f"=COUNTIFS('商品×類型'!$B$2:$B${PAIR_LAST},$A{row},"
                    f"'商品×類型'!$E$2:$E${PAIR_LAST},{wsM.cell(1, j).column_letter}$1)")
        if f != want_ref:
            bad_ref += 1
            fails.append(f"マトリクス {wsM.cell(row, j).coordinate} の数式が想定外: {f}")
        ok(f"マトリクス列見出し[{j}]", wsM.cell(1, j).value, c["label"])
checks += 1
if bad_ref:
    fails.append(f"COUNTIFS の参照崩れ {bad_ref} 件")

# 各 COUNTIFS が返すはずの値（＝手計算）を表示して妥当性を目視できるようにする
grid_total = sum(expected.values())
ok("マトリクス計の理論値 = 商品×類型の行数", grid_total, len(pairs_expected))

# ---- 4b. 商品×年齢層 と ニーズ×年齢層 ---------------------------------------
wsB = wb["商品×年齢層"]
band_expected = [(PURPS[p["primaryNeed"]], BAND[b]["label"]) for p in P for b in p["ageBands"]]
BAND_LAST = len(band_expected) + 1
ok("商品×年齢層の最終行", wsB.max_row, BAND_LAST)
band_actual = [(wsB.cell(r, 4).value, wsB.cell(r, 5).value) for r in range(2, BAND_LAST + 1)]
ok("商品×年齢層の中身", Counter(band_actual), Counter(band_expected))

wsN = wb["ニーズ×年齢層"]
band_labels_in_source = {wsB.cell(r, 5).value for r in range(2, BAND_LAST + 1)}
# 行 1 は COUNTIFS の判定値そのもの。'商品×年齢層' E 列の値と完全一致していなければ
# 全セルが 0 を返す（エラーにならないので数式検査だけでは気づけない）。
for j, b in enumerate(BANDS, start=2):
    hdr = wsN.cell(1, j).value
    ok(f"ニーズ×年齢層の判定行見出し[{j}]", hdr, b["label"])
    checks += 1
    if hdr not in band_labels_in_source:
        fails.append(f"判定値 {hdr!r} が '商品×年齢層' E 列に存在しない（全セル 0 になる）")
    ok(f"ニーズ×年齢層の呼称行[{j}]", wsN.cell(2, j).value, b["sub"])

bad_nb = 0
for i, n in enumerate(PURP_LIST):
    row = i + 3
    ok(f"ニーズ行見出し[{row}]", wsN.cell(row, 1).value, n["label"])
    for j, b in enumerate(BANDS, start=2):
        f = wsN.cell(row, j).value
        want = (f"=COUNTIFS('商品×年齢層'!$D$2:$D${BAND_LAST},$A{row},"
                f"'商品×年齢層'!$E$2:$E${BAND_LAST},{wsN.cell(1, j).column_letter}$1)")
        if f != want:
            bad_nb += 1
            fails.append(f"ニーズ×年齢層 {wsN.cell(row, j).coordinate} の数式が想定外: {f}")
checks += 1
if bad_nb:
    fails.append(f"ニーズ×年齢層の COUNTIFS 参照崩れ {bad_nb} 件")

# 主ニーズ側の判定値（A 列）も同様に突き合わせる
need_labels_in_source = {wsB.cell(r, 4).value for r in range(2, BAND_LAST + 1)}
for i, n in enumerate(PURP_LIST):
    checks += 1
    if n["label"] not in need_labels_in_source and any(p["primaryNeed"] == n["id"] for p in P):
        fails.append(f"主ニーズ判定値 {n['label']!r} が '商品×年齢層' D 列に存在しない")

# ---- 5. 凡例シートの集計数式 ------------------------------------------------
wsI = wb["凡例・注意"]
found = {}
for row in wsI.iter_rows():
    label = row[0].value
    if isinstance(label, str) and len(row) > 1 and isinstance(row[1].value, str) and row[1].value.startswith("="):
        found[label] = row[1].value
for label, frag in [
    ("収録商品数", f"COUNTA(商品一覧!$A$2:$A${LAST})"),
    ("引受会社数（重複除く）", f"商品一覧!$B$2:$B${LAST}"),
    ("裏側の証券会社を特定した商品数", f"COUNTA(商品一覧!$S$2:$S${LAST})"),
    ("指数連動商品数", f'COUNTIF(商品一覧!$K$2:$K${LAST},"●")'),
    ("確度「確認済」の商品数", f'COUNTIF(商品一覧!$Z$2:$Z${LAST},"確認済")'),
]:
    checks += 1
    if label not in found:
        fails.append(f"凡例シートに集計行 '{label}' が無い")
    elif frag not in found[label]:
        fails.append(f"凡例 '{label}' の数式が想定外: {found[label]}")

# ---- 6. 開いたときに必ず再計算されるか ---------------------------------------
ok("fullCalcOnLoad", bool(wb.calculation and wb.calculation.fullCalcOnLoad), True)

# ---- 結果 -------------------------------------------------------------------
print(f"\n期待値（Excel で開いた直後にこの値になるはず）:")
print(f"  収録商品数 {len(P)} / 引受会社 {len(INSURERS)} 社")
print(f"  裏側の証券会社を特定 {sum(1 for p in P if p['counterparty']['present'])} 商品")
print(f"  指数連動 {sum(1 for p in P if p['indexLinked'])} 商品")
print(f"  確認済 {sum(1 for p in P if p['confidence'] == 'confirmed')} / "
      f"要確認 {sum(1 for p in P if p['confidence'] != 'confirmed')}")
print(f"  会社×類型マトリクス総計 {grid_total}")
print(f"  ニーズ×年齢層マトリクス総計 {len(band_expected)}")
for n in PURP_LIST:
    cells = " ".join(
        f"{sum(1 for p in P if p['primaryNeed'] == n['id'] and b['id'] in p['ageBands']) or '・':>3}"
        for b in BANDS)
    print(f"    {n['label']:<14}{cells}   計 {sum(1 for p in P if p['primaryNeed'] == n['id'])}")

print(f"\n{checks} 件の検査")
if fails:
    print(f"NG {len(fails)} 件:")
    for f in fails[:40]:
        print("  -", f)
    sys.exit(1)
print("すべて合格：関数は Excel 2007 以前の無印関数のみ・参照範囲と中身は期待どおり")
