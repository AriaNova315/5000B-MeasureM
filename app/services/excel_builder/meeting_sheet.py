"""
例会单 Sheet 构建器。
11 行 × 8 列结构：标题 | 编号 | 时间/地点 | 主持人/记录员 | 参加人员 | 议题×N
"""
from .utils import (
    _w, _mc, _wmc, _s, border_range,
    F_TITLE, F_HDR, F_BODY, A_CTR, A_LEFT, A_LEFT_W, B_THIN,
    set_col_width, set_row_height,
)


def build_meeting_sheet(
    wb,
    sheet_name: str,
    project: dict,
    phase: dict,
    meeting_number: int,
    report_number: str,
    get_nar,          # callable(section) -> str
):
    """
    在工作簿 wb 中追加一张例会单 Sheet。
    get_nar: 给定 section key ('time'|'location'|'host'|'attendees'|'agenda') 返回填充内容。
    """
    ws = wb.create_sheet(sheet_name)

    # ── 列宽 ─────────────────────────────────────────────────────────────────
    widths = [14, 20, 14, 20, 14, 20, 14, 20]
    for i, w in enumerate(widths, 1):
        set_col_width(ws, i, w)

    # ── Row 1: 标题 ───────────────────────────────────────────────────────────
    _wmc(ws, 1, 1, 1, 8, "项目例会单",
         font=F_TITLE, alignment=A_CTR)

    # ── Row 2: 编号 ───────────────────────────────────────────────────────────
    _wmc(ws, 2, 1, 2, 8, f"编号：{report_number}",
         font=F_BODY, alignment=A_LEFT)

    # ── Row 3: 时间 / 地点 ────────────────────────────────────────────────────
    _wmc(ws, 3, 1, 3, 2, "时间", font=F_HDR, alignment=A_CTR)
    _wmc(ws, 3, 3, 3, 4, get_nar("time"), font=F_BODY, alignment=A_LEFT)
    _wmc(ws, 3, 5, 3, 6, "地点", font=F_HDR, alignment=A_CTR)
    _wmc(ws, 3, 7, 3, 8, get_nar("location"), font=F_BODY, alignment=A_LEFT)

    # ── Row 4: 主持人 / 记录员 ────────────────────────────────────────────────
    _wmc(ws, 4, 1, 4, 2, "主持人", font=F_HDR, alignment=A_CTR)
    _wmc(ws, 4, 3, 4, 4, get_nar("host"), font=F_BODY, alignment=A_LEFT)
    _wmc(ws, 4, 5, 4, 6, "记录员", font=F_HDR, alignment=A_CTR)
    _wmc(ws, 4, 7, 4, 8, project.get("cm_person", ""), font=F_BODY, alignment=A_LEFT)

    # ── Row 5: 参加人员 ───────────────────────────────────────────────────────
    _wmc(ws, 5, 1, 5, 2, "参加人员", font=F_HDR, alignment=A_CTR)
    _wmc(ws, 5, 3, 5, 8, get_nar("attendees"), font=F_BODY, alignment=A_LEFT_W)
    set_row_height(ws, 5, 30)

    # ── Row 6+: 主要内容 ──────────────────────────────────────────────────────
    agenda_raw = get_nar("agenda")
    agenda_lines = [l.strip() for l in agenda_raw.strip().split("\n") if l.strip()] if agenda_raw else []
    if not agenda_lines:
        agenda_lines = ["（待填写）"]

    # 第一行有 "主要内容" 标签
    _wmc(ws, 6, 1, 6 + len(agenda_lines) - 1, 2, "主要内容",
         font=F_HDR, alignment=A_CTR)

    for i, line in enumerate(agenda_lines):
        r = 6 + i
        _wmc(ws, r, 3, r, 8, line, font=F_BODY, alignment=A_LEFT_W)
        set_row_height(ws, r, 30)

    last_row = 6 + len(agenda_lines) - 1

    # ── 边框 ─────────────────────────────────────────────────────────────────
    border_range(ws, 3, 1, last_row, 8)
