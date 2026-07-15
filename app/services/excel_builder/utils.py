"""
openpyxl 共享工具：样式常量、原子写入函数。
"""
from datetime import date as _date

from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# ── 字体 ─────────────────────────────────────────────────────────────────────
F_TITLE   = Font(name="黑体",  size=16, bold=True)
F_HDR     = Font(name="宋体",  size=11, bold=True)
F_BODY    = Font(name="宋体",  size=11)
F_WR_HDR  = Font(name="宋体",  size=10, bold=True)
F_WR_BODY = Font(name="宋体",  size=10)
F_PT_HDR      = Font(name="Arial", size=9, bold=True)
F_PT_BODY     = Font(name="Arial", size=9)
F_PT_BODY_RED = Font(name="Arial", size=9, color="FF0000")

# ── 对齐 ─────────────────────────────────────────────────────────────────────
A_CTR     = Alignment(horizontal="center", vertical="center")
A_CTR_W   = Alignment(horizontal="center", vertical="center", wrap_text=True)
A_LEFT    = Alignment(horizontal="left",   vertical="center")
A_LEFT_W  = Alignment(horizontal="left",   vertical="center", wrap_text=True)
A_RIGHT   = Alignment(horizontal="right",  vertical="center")

# ── 边框 ─────────────────────────────────────────────────────────────────────
_THIN = Side(style="thin")
_MED  = Side(style="medium")
B_THIN = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
B_MED  = Border(left=_MED,  right=_MED,  top=_MED,  bottom=_MED)


# ── 原子操作 ──────────────────────────────────────────────────────────────────

def _w(ws, r: int, c: int, v):
    ws.cell(row=r, column=c).value = v


def _s(ws, r: int, c: int, **kw):
    cell = ws.cell(row=r, column=c)
    for attr, val in kw.items():
        setattr(cell, attr, val)


def _ws(ws, r: int, c: int, v, **kw):
    _w(ws, r, c, v)
    _s(ws, r, c, **kw)


def _mc(ws, r1: int, c1: int, r2: int, c2: int):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def _wmc(ws, r1: int, c1: int, r2: int, c2: int, v, **kw):
    _mc(ws, r1, c1, r2, c2)
    _w(ws, r1, c1, v)
    _s(ws, r1, c1, **kw)


def border_range(ws, r1: int, c1: int, r2: int, c2: int, border=None):
    b = border or B_THIN
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = b


def set_col_width(ws, col: int, width: float):
    ws.column_dimensions[get_column_letter(col)].width = width


def set_row_height(ws, row: int, height: float):
    ws.row_dimensions[row].height = height


# ── 日期工具 ─────────────────────────────────────────────────────────────────

def excel_serial(d: _date) -> int:
    """Python date → Excel 序列号（1900 日期系统）。"""
    from datetime import date
    return (d - date(1899, 12, 31)).days


def fmt_date(d: _date) -> str:
    return d.strftime("%Y.%-m.%-d") if d else ""


def fmt_date_range(d1: _date, d2: _date) -> str:
    return f"{fmt_date(d1)}-{fmt_date(d2)}"
