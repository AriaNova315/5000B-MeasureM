"""
ProgressTracking Sheet 构建器。
固定 8 列（A-H） + 每个跟踪节点若干列：
  - 双周节点：4 列（计划应完成工作量 | 完成百分比 | 实际已完成工作量 | 完成工作量偏差）
  - 阶段节点：6 列（计划应完成工作量 | 实际完成百分比 | 实际已完成工作量 | 实际工时 | 进度偏差 | 工作量偏差）
包含 WBS 父子层级，父节点数据由叶节点聚合计算。
"""
from datetime import date

from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from .utils import (
    _w, _mc, _wmc, _s, border_range,
    F_PT_HDR, F_PT_BODY, F_PT_BODY_RED, A_CTR, A_CTR_W, A_LEFT, A_LEFT_W, B_THIN,
    set_col_width, set_row_height, excel_serial,
)

_N_FIXED = 8  # A-H 固定列数

_BW_HEADERS = ["计划应完成工作量", "完成百分比", "实际已完成工作量", "完成工作量偏差"]
_PH_HEADERS = ["计划应完成工作量", "实际完成百分比", "实际已完成工作量",
               "实际工时", "进度偏差（阶段统计）", "工作量偏差（阶段统计）"]


def _wr_formula(wr_name: str, rows: list, col: str, agg: str) -> str | None:
    """Build MIN/MAX formula referencing WorkReport cells. Returns None if no rows."""
    if not wr_name or not rows:
        return None
    refs = ",".join(f"'{wr_name}'!{col}{r}" for r in rows)
    return f'=IF(COUNT({refs})=0,"",{agg}({refs}))'


def _bw_dev(plan_should: float, actual: float) -> float:
    """完成工作量偏差 = (actual - plan_should) / plan_should。"""
    return round((actual - plan_should) / plan_should, 4) if plan_should else 0.0


def build_progress_tracking(
    wb,
    sheet_name: str,
    tasks: list,       # [{id, wbs, name, start_date, end_date, plan_hours, resources}]
    phase_nodes: list, # [{id, date, type}]
    tnd: dict,         # {(task_id, node_id): {plan_should, pct, actual}}
    phase: dict = None,          # {wbs, name, start_date, end_date} — 用于写入阶段汇总行
    wr_sheet_name: str = None,   # 最后一个 WorkReport 的 sheet 名
    task_wr_rows: dict = None,   # {task_id: [行号列表]}，来自 build_workreport 返回值
    workreport_refs: list[dict] = None,  # 每个节点的 {sheet, task_wr_rows}
):
    ws = wb.create_sheet(sheet_name)

    # 每个节点的列数（双周4列，阶段6列）及起始列
    node_is_bw     = [n["type"] == "双周" for n in phase_nodes]
    node_col_counts = [4 if bw else 6 for bw in node_is_bw]
    node_col_starts: list[int] = []
    offset = _N_FIXED + 1
    for cnt in node_col_counts:
        node_col_starts.append(offset)
        offset += cnt
    total_cols = offset - 1

    # ── 列宽 ─────────────────────────────────────────────────────────────────
    fixed_widths = [8, 28, 12, 12, 20, 12, 12, 10]
    for i, w in enumerate(fixed_widths, 1):
        set_col_width(ws, i, w)
    for i in range(_N_FIXED + 1, total_cols + 1):
        set_col_width(ws, i, 11)

    # ── Row 1: 固定表头 + 节点分组标题 ────────────────────────────────────────
    r = 1
    headers_fixed = ["WBS", "任务名称", "计划开始时间", "计划完成时间",
                     "资源名称", "实际开始时间", "实际完成时间", "计划总工作量"]
    for i, h in enumerate(headers_fixed, 1):
        _w(ws, r, i, h)
        _s(ws, r, i, font=F_PT_HDR, alignment=A_CTR_W)

    for k, node in enumerate(phase_nodes):
        c_start = node_col_starts[k]
        cnt     = node_col_counts[k]
        d = node["date"]
        track_type = "双周跟踪" if node_is_bw[k] else "阶段跟踪"
        label = f"{track_type}（{d.year}.{d.month}.{d.day}）"
        _wmc(ws, r, c_start, r, c_start + cnt - 1, label, font=F_PT_HDR, alignment=A_CTR)
    set_row_height(ws, r, 20)

    # ── Row 2: 节点子列标题 ───────────────────────────────────────────────────
    r = 2
    for i in range(1, _N_FIXED + 1):
        _s(ws, r, i, font=F_PT_HDR, alignment=A_CTR_W)

    for k in range(len(phase_nodes)):
        c_start  = node_col_starts[k]
        sub_hdrs = _BW_HEADERS if node_is_bw[k] else _PH_HEADERS
        for j, h in enumerate(sub_hdrs):
            _w(ws, r, c_start + j, h)
            _s(ws, r, c_start + j, font=F_PT_HDR, alignment=A_CTR_W)
    set_row_height(ws, r, 30)

    # ── 构建 WBS 映射与叶节点识别 ──────────────────────────────────────────────
    task_by_id  = {t["id"]: t for t in tasks}
    task_by_wbs = {t["wbs"]: t for t in tasks}

    def is_leaf(task):
        return bool(task.get("resources"))

    def get_leaf_descendants(parent_wbs: str) -> list:
        prefix = parent_wbs + "."
        return [t for t in tasks if t["wbs"].startswith(prefix) and is_leaf(t)]

    def get_direct_children(parent_wbs: str) -> list:
        """返回 parent_wbs 的直接子节点（比父节点多一个 WBS 层级）。"""
        prefix = parent_wbs + "."
        depth  = len(parent_wbs.split(".")) + 1
        return [t for t in tasks if t["wbs"].startswith(prefix) and len(t["wbs"].split(".")) == depth]

    # ── 预计算每个 task 在 ProgressTracking 中的行号 ──────────────────────────
    # 行 3 = 阶段汇总行（当 phase 存在时），任务数据行从 4 开始
    task_pt_row: dict[int, int] = {}
    _r = 4 if phase else 3
    for task in tasks:
        task_pt_row[task["id"]] = _r
        _r += 1

    # ── 阶段节点公式列写入（实际工时 / 进度偏差 / 工作量偏差）────────────────
    def _write_phase_formula_cols(row: int, k: int, wbs: str, leaf: bool):
        """为阶段跟踪节点写入 col+3/+4/+5 的 Excel 公式。"""
        if node_is_bw[k]:
            return
        c_start      = node_col_starts[k]
        col_act_hrs  = get_column_letter(c_start + 3)  # 实际工时列
        col_actual   = get_column_letter(c_start + 2)  # 实际已完成工作量列

        # 实际工时
        if leaf:
            h_formula = (
                f"=SUMIF('{wr_sheet_name}'!B:B,A{row},'{wr_sheet_name}'!K:K)"
                if wr_sheet_name else "=0"
            )
        else:
            children   = get_direct_children(wbs)
            child_rows = [task_pt_row[t["id"]] for t in children if t["id"] in task_pt_row]
            h_formula  = ("=" + "+".join(f"{col_act_hrs}{cr}" for cr in child_rows)
                          if child_rows else "=0")
        ws.cell(row, c_start + 3).value = h_formula
        _s(ws, row, c_start + 3, font=F_PT_BODY, alignment=A_CTR, number_format="0.00")

        # 进度偏差 = (实际完成时间 - 计划完成时间) / (计划完成时间 - 计划开始时间 + 1)
        ws.cell(row, c_start + 4).value = (
            f'=IF(G{row}="","",(G{row}-D{row})/(D{row}-C{row}+1))'
        )
        _s(ws, row, c_start + 4, font=F_PT_BODY, alignment=A_CTR, number_format="0.00%")

        # 工作量偏差 = (实际工时 - 实际已完成工作量) / 实际已完成工作量
        ws.cell(row, c_start + 5).value = (
            f'=IF({col_actual}{row}=0,"",({col_act_hrs}{row}-{col_actual}{row})/{col_actual}{row})'
        )
        _s(ws, row, c_start + 5, font=F_PT_BODY, alignment=A_CTR, number_format="0.00%")

    # ── 写入节点数据（双周前 3 列静态 + 第 4 列公式 / 阶段前 3 列静态）──────────
    def _write_node_data(row: int, k: int, plan_should: float, pct: float,
                         actual: float, prev_actual: float, plan_h: float,
                         pct_formula: str = None, actual_formula: str = None):
        c_start = node_col_starts[k]
        if node_is_bw[k]:
            # 前 3 列：计划应完成工作量 | 完成百分比 | 实际已完成工作量（静态值）
            vals = [plan_should, pct_formula or pct, actual_formula or actual]
            fmts = ["0.00", "0.00%", "0.00"]
            for j, (v, fmt) in enumerate(zip(vals, fmts)):
                _w(ws, row, c_start + j, v)
                _s(ws, row, c_start + j, font=F_PT_BODY, alignment=A_CTR, number_format=fmt)
            # 第 4 列：完成工作量偏差 = (实际已完成 - 计划应完成) / 计划应完成（公式）
            col_ps = get_column_letter(c_start)
            col_ac = get_column_letter(c_start + 2)
            ws.cell(row, c_start + 3).value = (
                f'=IF({col_ps}{row}=0,"",({col_ac}{row}-{col_ps}{row})/{col_ps}{row})'
            )
            _s(ws, row, c_start + 3, font=F_PT_BODY, alignment=A_CTR, number_format="0.00%")
        else:
            # 仅写前 3 列静态值；后 3 列由 _write_phase_formula_cols 写公式
            vals = [plan_should, pct_formula or pct, actual_formula or actual]
            fmts = ["0.00", "0.00%", "0.00"]
            for j, (v, fmt) in enumerate(zip(vals, fmts)):
                _w(ws, row, c_start + j, v)
                _s(ws, row, c_start + j, font=F_PT_BODY, alignment=A_CTR, number_format=fmt)

    # ── 数据行 ────────────────────────────────────────────────────────────────
    r = 3
    leaf_tasks = [t for t in tasks if is_leaf(t)]

    # 阶段汇总行
    if phase:
        _w(ws, r, 1, phase["wbs"])
        _w(ws, r, 2, phase["name"])
        _w(ws, r, 3, phase.get("start_date"))
        _w(ws, r, 4, phase.get("end_date"))
        for col in range(1, _N_FIXED + 1):
            _s(ws, r, col, font=F_PT_BODY, alignment=A_LEFT)
        for dc in (3, 4, 6, 7):
            _s(ws, r, dc, number_format="YYYY/MM/DD")
        summary_wr_rows = []
        for task in leaf_tasks:
            summary_wr_rows.extend((task_wr_rows or {}).get(task["id"], []))
        summary_start_formula = _wr_formula(wr_sheet_name, summary_wr_rows, "G", "MIN")
        summary_end_formula = _wr_formula(wr_sheet_name, summary_wr_rows, "H", "MAX")
        ws.cell(r, 6).value = summary_start_formula or ""
        ws.cell(r, 7).value = summary_end_formula or ""
        plan_h_total = sum(t["plan_hours"] for t in leaf_tasks)
        _w(ws, r, 8, round(plan_h_total, 2))

        prev_actual_phase = 0.0
        for k, node in enumerate(phase_nodes):
            nid = node["id"]
            ps  = round(sum(tnd.get((t["id"], nid), {}).get("plan_should", 0) for t in leaf_tasks), 2)
            ac  = round(sum(tnd.get((t["id"], nid), {}).get("actual",      0) for t in leaf_tasks), 2)
            pct = round(ac / plan_h_total, 4) if plan_h_total > 0 else 0.0
            actual_col = get_column_letter(node_col_starts[k] + 2)
            actual_formula = "=" + "+".join(
                f"{actual_col}{task_pt_row[t['id']]}" for t in leaf_tasks
            ) if leaf_tasks else "=0"
            pct_formula = f'=IF(H{r}=0,0,{actual_col}{r}/H{r})'
            _write_node_data(
                r, k, ps, pct, ac, prev_actual_phase, plan_h_total,
                pct_formula=pct_formula, actual_formula=actual_formula,
            )
            _write_phase_formula_cols(r, k, phase["wbs"], leaf=False)
            prev_actual_phase = ac
        set_row_height(ws, r, 18)
        r += 1

    for task in tasks:
        leaf        = is_leaf(task)
        descendants = [] if leaf else get_leaf_descendants(task["wbs"])
        resources_str = ",".join(res["name"] for res in task.get("resources", [])) if leaf else None

        # 收集该任务在最后 WorkReport 中的行号（用于公式引用）
        if leaf:
            wr_rows = (task_wr_rows or {}).get(task["id"], [])
        else:
            wr_rows = []
            for d in descendants:
                wr_rows.extend((task_wr_rows or {}).get(d["id"], []))

        # 固定列
        _w(ws, r, 1, task["wbs"])
        _w(ws, r, 2, task["name"])
        _w(ws, r, 3, task["start_date"])
        _w(ws, r, 4, task["end_date"])
        _w(ws, r, 5, resources_str)

        # 实际开始/完成：公式引用最后 WorkReport，回退到计划日期
        f6 = _wr_formula(wr_sheet_name, wr_rows, "G", "MIN")
        f7 = _wr_formula(wr_sheet_name, wr_rows, "H", "MAX")
        ws.cell(r, 6).value = f6 if f6 is not None else task.get("start_date")
        ws.cell(r, 7).value = f7 if f7 is not None else task.get("end_date")

        if leaf:
            _w(ws, r, 8, task["plan_hours"])
        else:
            descendant_rows = [task_pt_row[d["id"]] for d in descendants]
            ws.cell(r, 8).value = "=" + "+".join(
                f"H{child_row}" for child_row in descendant_rows
            ) if descendant_rows else "=0"

        for col in range(1, _N_FIXED + 1):
            _s(ws, r, col, font=F_PT_BODY, alignment=A_LEFT)
        for dc in (3, 4, 6, 7):
            _s(ws, r, dc, number_format="YYYY/MM/DD")

        # 动态列
        prev_actual = 0.0
        for k, node in enumerate(phase_nodes):
            nid = node["id"]
            if leaf:
                td          = tnd.get((task["id"], nid), {})
                plan_should = td.get("plan_should", 0.0)
                pct         = td.get("pct",         0.0)
                actual      = td.get("actual",       0.0)
            else:
                plan_should = round(sum(tnd.get((d["id"], nid), {}).get("plan_should", 0) for d in descendants), 2)
                actual      = round(sum(tnd.get((d["id"], nid), {}).get("actual",      0) for d in descendants), 2)
                total_plan  = sum(d["plan_hours"] for d in descendants) or task["plan_hours"] or 1
                pct         = round(actual / total_plan, 4) if total_plan > 0 else 0.0

            c_start = node_col_starts[k]
            actual_col = get_column_letter(c_start + 2)
            pct_formula = None
            actual_formula = None
            if leaf:
                ref = (workreport_refs or [])[k] if k < len(workreport_refs or []) else None
                wr_rows = (ref or {}).get("task_wr_rows", {}).get(task["id"], [])
                if ref and wr_rows:
                    wr_sheet = ref["sheet"].replace("'", "''")
                    first_wr_row = wr_rows[0]
                    pct_formula = f"='{wr_sheet}'!I{first_wr_row}"
                    actual_formula = f"='{wr_sheet}'!J{first_wr_row}"
            else:
                descendant_rows = [task_pt_row[d["id"]] for d in descendants]
                actual_formula = "=" + "+".join(
                    f"{actual_col}{child_row}" for child_row in descendant_rows
                ) if descendant_rows else "=0"
                pct_formula = f'=IF(H{r}=0,0,{actual_col}{r}/H{r})'

            _write_node_data(
                r, k, plan_should, pct, actual, prev_actual, task["plan_hours"] or 0,
                pct_formula=pct_formula, actual_formula=actual_formula,
            )
            _write_phase_formula_cols(r, k, task["wbs"], leaf=leaf)
            prev_actual = actual

        set_row_height(ws, r, 18)
        r += 1

    # ── 偏差列：非零单元格红色字体（条件格式）────────────────────────────────
    # 双周节点：完成工作量偏差（col+3）；阶段节点：工作量偏差（col+5）
    data_start = 3  # 阶段汇总行或第一条数据行
    data_end   = r - 1
    for k in range(len(phase_nodes)):
        dev_col = node_col_starts[k] + (3 if node_is_bw[k] else 5)
        col_letter = get_column_letter(dev_col)
        cell_range = f"{col_letter}{data_start}:{col_letter}{data_end}"
        ws.conditional_formatting.add(
            cell_range,
            FormulaRule(
                formula=[f'AND({col_letter}{data_start}<>0,{col_letter}{data_start}<>"")'],
                font=Font(name="Arial", size=9, color="FF0000"),
            )
        )

    # ── 边框 ─────────────────────────────────────────────────────────────────
    border_range(ws, 1, 1, r - 1, total_cols)
