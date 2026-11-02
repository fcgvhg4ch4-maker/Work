"""openpyxl の薄いラッパ：表（列定義＋数式テンプレート）、名前定義、書式。

数式テンプレートのトークン
  {Key}     同じ表・同じ行のセル（例 {Notional} -> K12）
  {Key@1}   同じ表の先頭データ行の絶対参照（例 $K$6）。累積範囲 {Key@1}:{Key} に使う
  {r}       行番号
  {n}       表内の連番（1 始まり）
それ以外の参照は名前定義（PH_Date など）で書く。
"""
from __future__ import annotations

import datetime as dt
import re

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula

# ---------------------------------------------------------------- style
JP = "Meiryo UI"
NAVY = "1F3864"
NAVY2 = "2F5496"
GREY = "D9D9D9"
GREY2 = "F2F2F2"
GREY3 = "808080"
WHITE = "FFFFFF"
INPUT_BG = "FFF8E1"     # 入力セル（淡いクリーム）
INPUT_FONT = "1F4E9A"   # 入力値は青字
RED = "C00000"          # リスクアラート専用
RED_LIGHT = "F4CCCC"
AMBER_LIGHT = "FCE4D6"

F_TITLE = Font(name=JP, size=15, bold=True, color=WHITE)
F_SUB = Font(name=JP, size=9, color="404040")
F_H2 = Font(name=JP, size=11, bold=True, color=NAVY)
F_HEAD = Font(name=JP, size=9, bold=True, color=WHITE)
F_BODY = Font(name=JP, size=9, color="000000")
F_INPUT = Font(name=JP, size=9, color=INPUT_FONT)
F_SMALL = Font(name=JP, size=8, color="595959")
F_BIG = Font(name=JP, size=16, bold=True, color=NAVY)

FILL_TITLE = PatternFill("solid", fgColor=NAVY)
FILL_HEAD = PatternFill("solid", fgColor=NAVY2)
FILL_HEAD_IN = PatternFill("solid", fgColor="44546A")
FILL_INPUT = PatternFill("solid", fgColor=INPUT_BG)
FILL_GREY = PatternFill("solid", fgColor=GREY2)
FILL_GREY2 = PatternFill("solid", fgColor=GREY)
FILL_RED = PatternFill("solid", fgColor=RED)
FILL_RED_L = PatternFill("solid", fgColor=RED_LIGHT)
FILL_AMBER_L = PatternFill("solid", fgColor=AMBER_LIGHT)

_thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

FMT_DATE = "yyyy/mm/dd;;;@"
FMT_PCT = "0.00%"
FMT_PCT1 = "0.0%"
FMT_NUM = "#,##0.00"
FMT_INT = "#,##0"
FMT_4 = "0.0000"
FMT_6 = "0.000000"

XLFN = ["MAXIFS", "MINIFS", "TEXTJOIN", "IFS", "IFNA", "CONCAT", "SWITCH"]
_xlfn_re = re.compile(r"(?<![\w.])(" + "|".join(XLFN) + r")\(")


def fx(s: str) -> str:
    """Excel 2016 以降の関数に _xlfn. を付ける（付けないと Excel で #NAME? になる）。"""
    return _xlfn_re.sub(r"_xlfn.\1(", s)


def q(sheet: str) -> str:
    return "'" + sheet + "'"


def add_name(wb, name, sheet, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=f"{q(sheet)}!{ref}")


def set_formula(ws, cell, formula, array=False):
    f = fx(formula)
    if array:
        ws[cell] = ArrayFormula(cell, f)
    else:
        ws[cell] = f


def title(ws, text, sub, width_cols=14):
    ws.sheet_view.showGridLines = False
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    for c in range(1, width_cols + 1):
        ws.cell(1, c).fill = FILL_TITLE
    ws.row_dimensions[1].height = 26
    ws["A2"] = sub
    ws["A2"].font = F_SUB
    ws["A3"] = "凡例：■ 薄黄色＋青字＝入力セル　■ 白＝計算セル（数式）　■ 灰色＝補助計算　■ 赤＝リスクアラートのみ"
    ws["A3"].font = F_SMALL


def section(ws, row, text, col=1):
    c = ws.cell(row, col, text)
    c.font = F_H2


def style_value(cell, kind, fmt=None):
    cell.font = F_INPUT if kind == "in" else F_BODY
    if kind == "in":
        cell.fill = FILL_INPUT
    elif kind == "aux":
        cell.fill = FILL_GREY
    cell.border = BORDER
    if fmt:
        cell.number_format = fmt


def _tok(tpl, cols, r, n, first):
    def rep(m):
        k, a = m.group(1), m.group(2)
        if k == "r":
            return str(r)
        if k == "n":
            return str(n)
        if k not in cols:
            raise KeyError(f"unknown token {{{k}}} in: {tpl}")
        if a:
            return f"${cols[k]}${first}"
        return f"{cols[k]}{r}"
    return re.sub(r"\{(\w+)(@1)?\}", rep, tpl)


class Table:
    """1 行 1 レコードの表。cols は dict のリスト:
       key, head, kind('in'|'calc'|'aux'), w, fmt, f(テンプレート), arr(bool), note
    """

    def __init__(self, wb, ws, prefix, hdr_row, cols, cap, first_col=1, name_cols=True):
        self.wb, self.ws, self.prefix = wb, ws, prefix
        self.hdr, self.first, self.last = hdr_row, hdr_row + 1, hdr_row + cap
        self.cap = cap
        self.cols = {}
        self.spec = cols
        for i, c in enumerate(cols):
            self.cols[c["key"]] = get_column_letter(first_col + i)
        # header
        for c in cols:
            L = self.cols[c["key"]]
            cell = ws[f"{L}{hdr_row}"]
            cell.value = c["head"]
            cell.font = F_HEAD
            cell.fill = FILL_HEAD_IN if c.get("kind") == "in" else FILL_HEAD
            cell.alignment = CENTER
            cell.border = BORDER
            if c.get("w"):
                ws.column_dimensions[L].width = c["w"]
        ws.row_dimensions[hdr_row].height = 30
        if name_cols:
            for c in cols:
                L = self.cols[c["key"]]
                add_name(wb, f"{prefix}_{c['key']}", ws.title, f"${L}${self.first}:${L}${self.last}")

    def col(self, key):
        return self.cols[key]

    def rng(self, key):
        L = self.cols[key]
        return f"{q(self.ws.title)}!${L}${self.first}:${L}${self.last}"

    def cell(self, key, n):
        return f"{self.cols[key]}{self.first + n - 1}"

    def write(self, rows=None, formulas=True):
        """rows: 入力列の値（dict のリスト）。数式列は容量いっぱいまで書く。"""
        rows = rows or []
        for n in range(1, self.cap + 1):
            r = self.first + n - 1
            data = rows[n - 1] if n - 1 < len(rows) else {}
            for c in self.spec:
                L = self.cols[c["key"]]
                cell = self.ws[f"{L}{r}"]
                kind = c.get("kind", "calc")
                if c.get("f") and formulas:
                    f = _tok(c["f"], self.cols, r, n, self.first)
                    set_formula(self.ws, f"{L}{r}", f, array=c.get("arr", False))
                elif kind == "in" or c["key"] in data:
                    v = data.get(c["key"])
                    if isinstance(v, dt.datetime):
                        v = v.date()
                    cell.value = v
                style_value(cell, kind, c.get("fmt"))
                if c.get("wrap"):
                    cell.alignment = WRAP


def dv_list(ws, rng, formula, allow_blank=True):
    dv = DataValidation(type="list", formula1=formula, allow_blank=allow_blank)
    dv.error = "リストから選択してください"
    ws.add_data_validation(dv)
    dv.add(rng)


def severity_cf(ws, rng):
    """重要度セルの条件付き書式（赤はアラートのみに使用）"""
    first = rng.split(":")[0].replace("$", "")
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="High"'], fill=FILL_RED, font=Font(name=JP, size=9, bold=True, color=WHITE)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="Medium"'], fill=FILL_RED_L, font=Font(name=JP, size=9, bold=True, color=RED)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="Low"'], fill=FILL_GREY2))


def text_cf(ws, rng, texts, strong=True):
    first = rng.split(":")[0].replace("$", "")
    for t in texts:
        if strong:
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="{t}"'], fill=FILL_RED, font=Font(name=JP, size=9, bold=True, color=WHITE)))
        else:
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="{t}"'], fill=FILL_RED_L, font=Font(name=JP, size=9, color=RED)))
