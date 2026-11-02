"""Workbook-writing helpers: table builder, formula expansion, styling."""
import re
from dataclasses import dataclass, field

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableFormula, TableStyleInfo

JP = "Yu Gothic"
INK = "1F2933"

FILL = {
    "raw": PatternFill("solid", fgColor="EDEDED"),     # 元データ：淡いグレー
    "input": PatternFill("solid", fgColor="FFF2CC"),   # 入力：淡い黄色
    "calc": PatternFill("solid", fgColor="DDEBF7"),    # 計算：淡い青
    "head": PatternFill("solid", fgColor="1C3F6E"),
    "sect": PatternFill("solid", fgColor="E2EAF4"),
}
SEV_FILL = {
    "High": "FF7C80", "Medium": "F4B183", "Low": "FFE699",
    "READY": "A9D08E", "REVIEW REQUIRED": "F4B183", "DATA REQUIRED": "FF7C80",
    "RED": "FF7C80", "AMBER": "F4B183", "YELLOW": "FFE699", "GREEN": "A9D08E",
}

F_TITLE = Font(name=JP, size=14, bold=True, color=INK)
F_NOTE = Font(name=JP, size=9, color="52606D")
F_HEAD = Font(name=JP, size=9, bold=True, color="FFFFFF")
F_BODY = Font(name=JP, size=9, color=INK)
F_BOLD = Font(name=JP, size=10, bold=True, color=INK)
F_SECT = Font(name=JP, size=11, bold=True, color="1C3F6E")
F_HELP = Font(name=JP, size=8, color="A0A0A0")
_thin = Side(style="thin", color="C9D1DB")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
WRAP = Alignment(vertical="top", wrap_text=True)
TOP = Alignment(vertical="top")

FMT = {
    "date": "yyyy-mm-dd", "dt": "yyyy-mm-dd hh:mm", "amt": "#,##0",
    "pct": "0.00%", "pct1": "0.0%", "nav": "#,##0.0000", "num": "#,##0.00",
    "num4": "0.0000", "int": "0", "bps": '#,##0" bps"', "text": "@",
}

# Functions introduced after ECMA-376 1st ed. that must carry the _xlfn. prefix.
XLFN = {"MAXIFS", "MINIFS", "TEXTJOIN", "IFS", "SWITCH", "CONCAT", "IFNA"}


def xlfn(formula: str) -> str:
    def rep(m):
        name = m.group(1)
        return ("_xlfn." + name + "(") if name in XLFN else m.group(0)
    return re.sub(r"(?<![\w.])([A-Z][A-Z0-9]*)\(", rep, formula)


def expand(formula: str, table: str | None) -> str:
    """'§[col]' -> own-table column; '[@col]' -> this-row reference."""
    if table:
        formula = formula.replace("§", table)
        formula = re.sub(r"\[@([^\]]+)\]", lambda m: f"{table}[[#This Row],[{m.group(1)}]]", formula)
    return xlfn(formula)


@dataclass
class Col:
    name: str
    kind: str = "raw"          # raw | input | calc
    f: str | None = None       # formula template (starts with '=')
    fmt: str | None = None
    w: float = 12
    arr: bool = False          # write as single-cell array (CSE) formula
    dv: str | None = None      # inline data-validation list "A,B,C"


@dataclass
class TableSpec:
    name: str
    cols: list
    top: int = 7
    left: int = 1
    rows: list = field(default_factory=list)   # list of dicts for raw/input cols
    nrows: int | None = None                   # capacity for mirror tables


def write_table(ws, spec: TableSpec, dv_registry=None):
    from openpyxl.worksheet.datavalidation import DataValidation
    n = spec.nrows if spec.nrows else max(1, len(spec.rows))
    top, left = spec.top, spec.left
    for j, c in enumerate(spec.cols):
        cell = ws.cell(top, left + j, c.name)
        cell.font, cell.fill, cell.border = F_HEAD, FILL["head"], BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(left + j)].width = c.w
    for i in range(n):
        r = top + 1 + i
        data = spec.rows[i] if i < len(spec.rows) else {}
        for j, c in enumerate(spec.cols):
            cell = ws.cell(r, left + j)
            if c.f:
                text = expand(c.f, spec.name)
                if c.arr:
                    ref = f"{get_column_letter(left + j)}{r}"
                    ws[ref] = ArrayFormula(ref, text)
                else:
                    cell.value = text
            else:
                v = data.get(c.name)
                if v is not None:
                    cell.value = v
            cell.fill = FILL[c.kind]
            cell.font = F_BODY
            cell.border = BORDER
            if c.fmt:
                cell.number_format = FMT.get(c.fmt, c.fmt)
    ref = f"{get_column_letter(left)}{top}:{get_column_letter(left + len(spec.cols) - 1)}{top + n}"
    t = Table(displayName=spec.name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    t._initialise_columns()
    for j, c in enumerate(spec.cols):
        if c.f:
            t.tableColumns[j].calculatedColumnFormula = TableFormula(
                array=bool(c.arr), attr_text=expand(c.f, spec.name)[1:])
    ws.add_table(t)
    for j, c in enumerate(spec.cols):
        if c.dv:
            col = get_column_letter(left + j)
            dv = DataValidation(type="list", formula1=f'"{c.dv}"', allow_blank=True)
            dv.add(f"{col}{top + 1}:{col}{top + n + 500}")
            ws.add_data_validation(dv)
    return {c.name: get_column_letter(left + j) for j, c in enumerate(spec.cols)}, top + 1, top + n


def intro(ws, title, purpose, source, update, notes):
    ws["A1"] = title
    ws["A1"].font = F_TITLE
    for i, (lab, txt) in enumerate([("目的", purpose), ("データソース", source),
                                     ("更新方法", update), ("注意事項", notes)], start=2):
        c = ws.cell(i, 1, f"{lab}：{txt}")
        c.font = F_NOTE
        c.alignment = Alignment(vertical="top", wrap_text=False)
    ws.sheet_view.showGridLines = False


def legend(ws, row, col):
    items = [("入力セル", "input"), ("計算セル", "calc"), ("元データ", "raw")]
    for k, (lab, kind) in enumerate(items):
        c = ws.cell(row, col + k, lab)
        c.fill, c.font, c.border = FILL[kind], F_NOTE, BORDER


def add_status_cf(ws, rng):
    from openpyxl.formatting.rule import CellIsRule
    for val, color in SEV_FILL.items():
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=[f'"{val}"'],
                            fill=PatternFill("solid", fgColor=color, bgColor=color)))
