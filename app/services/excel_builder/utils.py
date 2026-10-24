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

# ── 填充色 ───────────────────────────────────────────────────────────────────
# 与权威样例及阶段1进展报告保持一致：分区、标签/表头、正文三层视觉层级。
FILL_SECTION = PatternFill("solid", fgColor="D9E1F2")
FILL_LABEL = PatternFill("solid", fgColor="E7E6E6")
FILL_GRID_HEADER = PatternFill("solid", fgColor="DFE3E8")
FILL_BODY = PatternFill("solid", fgColor="FFFFFF")


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


def fill_range(ws, r1: int, c1: int, r2: int, c2: int, fill):
    """填充完整区域，包括合并区域内非左上角的单元格。"""
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).fill = fill


def set_outer_border(ws, r1: int, c1: int, r2: int, c2: int):
    """设置中等线宽外框，同时保留区域内部原有边框。"""
    for r in range(r1, r2 + 1):
        left = ws.cell(row=r, column=c1)
        left.border = Border(
            left=_MED, right=left.border.right,
            top=left.border.top, bottom=left.border.bottom,
        )
        right = ws.cell(row=r, column=c2)
        right.border = Border(
            left=right.border.left, right=_MED,
            top=right.border.top, bottom=right.border.bottom,
        )
    for c in range(c1, c2 + 1):
        top = ws.cell(row=r1, column=c)
        top.border = Border(
            left=top.border.left, right=top.border.right,
            top=_MED, bottom=top.border.bottom,
        )
        bottom = ws.cell(row=r2, column=c)
        bottom.border = Border(
            left=bottom.border.left, right=bottom.border.right,
            top=bottom.border.top, bottom=_MED,
        )


def adaptive_row_height(
    text,
    *,
    min_height: float = 25,
    chars_per_line: int = 45,
    line_height: float = 15,
    max_height: float = 150,
) -> float:
    """按换行和文本长度估算显式行高，兼容不自动伸展行高的 Excel 版本。"""
    if text is None:
        return min_height
    lines = str(text).splitlines() or [""]
    visual_lines = sum(max(1, (len(line) + chars_per_line - 1) // chars_per_line) for line in lines)
    return min(max_height, max(min_height, visual_lines * line_height))


def style_bold_labels(ws, r1: int, c1: int, r2: int, c2: int):
    """将生成器已用粗体明确标识的标签/表头统一填充为浅灰色。"""
    fill_range(ws, r1, c1, r2, c2, FILL_BODY)
    for merged in ws.merged_cells.ranges:
        if (
            merged.min_row >= r1 and merged.max_row <= r2
            and merged.min_col >= c1 and merged.max_col <= c2
        ):
            if ws.cell(merged.min_row, merged.min_col).font.bold:
                fill_range(
                    ws, merged.min_row, merged.min_col,
                    merged.max_row, merged.max_col, FILL_LABEL,
                )
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            if cell.font.bold:
                cell.fill = FILL_LABEL


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
