# MeasureM 开发计划

> 基于架构评审（2026-07-06）定稿  
> 技术栈：Python + FastAPI + SQLite(aiosqlite) + Alpine.js + openpyxl  
> 端口：8002 | 启动：`python3 run.py`

---

## 阶段一：项目脚手架

**目标**：跑起来一个空的 FastAPI 服务，访问 http://localhost:8002 能看到页面。

- [ ] 创建目录结构
  ```
  MeasureM/
  ├── run.py
  ├── requirements.txt
  ├── generated/
  ├── tests/
  └── app/
      ├── __init__.py
      ├── main.py
      ├── database.py
      ├── models.py
      ├── routers/
      │   ├── __init__.py
      │   ├── projects.py
      │   └── export.py
      ├── services/
      │   ├── __init__.py
      │   ├── cn_holidays.py
      │   ├── workhour_engine.py
      │   ├── importer.py
      │   ├── report_templates.py
      │   └── excel_builder/
      │       ├── __init__.py
      │       ├── _primitives.py
      │       ├── workreport.py
      │       ├── progress_tracking.py
      │       ├── measure_report.py
      │       ├── progress_report.py
      │       └── meeting_sheet.py
      └── static/
          └── index.html
  ```
- [ ] `requirements.txt`：fastapi uvicorn aiosqlite openpyxl sse-starlette
- [ ] `run.py`：uvicorn 启动，端口 8002
- [ ] `app/main.py`：FastAPI 实例 + 路由注册 + 静态文件挂载
- [ ] `app/static/index.html`：占位页面（空的 Alpine.js 骨架，4 个 Tab）

**完成标志**：`python3 run.py` 启动不报错，浏览器打开 localhost:8002 能看到页面。

---

## 阶段二：数据库层

**目标**：建好所有表，启动时自动初始化。

- [ ] `app/database.py`：aiosqlite 连接 + 建表 DDL

  ```sql
  -- 建表顺序（按依赖关系）
  CREATE TABLE IF NOT EXISTS projects (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT, code TEXT, manager TEXT,
      team_members TEXT,  -- JSON数组
      qa_person TEXT, cm_person TEXT,
      created_at TEXT
  );

  CREATE TABLE IF NOT EXISTS phases (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      project_id INTEGER REFERENCES projects(id),
      wbs TEXT,        -- '1', '2', '3'...
      name TEXT,       -- '系统分析与设计', '需求分析阶段'...
      start_date TEXT, -- YYYY-MM-DD
      end_date TEXT
  );

  CREATE TABLE IF NOT EXISTS tracking_nodes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      phase_id INTEGER REFERENCES phases(id),
      date TEXT,       -- YYYY-MM-DD，跟踪节点日期
      node_type TEXT   -- '双周' | '阶段'
  );

  CREATE TABLE IF NOT EXISTS wbs_tasks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      phase_id INTEGER REFERENCES phases(id),
      wbs TEXT,        -- '2.1.1', '2.1.2'...
      name TEXT,
      start_date TEXT,
      end_date TEXT,
      plan_hours REAL,
      resources TEXT   -- JSON数组，人员名称
  );

  CREATE TABLE IF NOT EXISTS task_node_data (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      task_id INTEGER REFERENCES wbs_tasks(id),
      node_id INTEGER REFERENCES tracking_nodes(id),
      plan_should_complete REAL,  -- 计划应完成工作量
      pct_complete REAL,          -- 完成百分比 0~1
      actual_complete REAL        -- 实际已完成工作量
  );

  CREATE TABLE IF NOT EXISTS narratives (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      tracking_node_id INTEGER REFERENCES tracking_nodes(id),
      report_type TEXT,  -- 见下方枚举
      section TEXT,      -- 段落标识
      content TEXT
      -- report_type枚举：
      -- 'measure_biweekly' | 'measure_phase' | 'measure_milestone' | 'measure_phase1'
      -- 'progress_biweekly' | 'progress_phase' | 'progress_milestone' | 'progress_phase1'
      -- 'meeting'
  );
  ```

- [ ] `app/main.py` 启动时调用 `await init_db()`

**完成标志**：启动后 `measurem.db` 自动创建，用 DB Browser 或命令行能看到所有表。

---

## 阶段三：节假日表 + 工时算法

**目标**：能正确分配日工时，确定性随机种子保证一致性。

- [ ] `app/services/cn_holidays.py`
  - 内置 2025-2026 年法定节假日集合（set of 'YYYY-MM-DD'）
  - `is_workday(date: date) -> bool`
  - `get_workdays(start: date, end: date) -> list[date]`
  - 超出 2025-2026 范围：退化为仅排除周末，并 `logging.warning(...)`

- [ ] `app/services/workhour_engine.py`
  - `generate_daily_hours(task_wbs, person_name, start_date, end_date, node_date, target_hours, exact=False) -> dict[date, float]`
  - 随机种子：`random.seed(hash(task_wbs + person_name))`
  - 正常模式：总和 ∈ [target×0.8, target×1.2]
  - Phase1精确模式（exact=True）：总和精确等于 target，末日调整夹紧[1,8]不够则向前均摊
  - 防御断言：`assert len(workdays) > 0` / `assert mean <= 8`

- [ ] `tests/test_workhour_engine.py`
  - 测试1：同参数调用两次，重叠日期工时相同（确定性种子验证）
  - 测试2：Phase1 exact=True，sum == target（精确模式）
  - 测试3：普通模式，sum ∈ [target×0.8, target×1.2]
  - 测试4：单日任务（start==end，有工时）

**完成标志**：`python3 -m pytest tests/test_workhour_engine.py -v` 全部通过。

---

## 阶段四：导入服务

**目标**：上传两个 Excel 文件，解析后写入数据库，并自动生成默认叙述文字。

- [ ] `app/services/importer.py`
  - `parse_project_info(path) -> dict`：读取"项目信息" Sheet
  - `parse_schedule(path) -> list[dict]`：读取"项目进度计划" Sheet，识别阶段（顶层WBS）和叶节点任务
  - `parse_tracking_nodes(path) -> dict`：读取跟踪节点数据.xlsx，Sheet名格式 `{阶段WBS}-{YYYYMMDD}`
  - 节点类型识别：任务名含"双周"→双周，含"阶段"→阶段，无匹配→阶段
  - `import_all(info_path, tracking_path, db) -> int`：主入口，写入所有表，返回 project_id
  - 最后调用 `report_templates.generate_all_narratives(project_id, db)`

- [ ] `tests/test_importer.py`
  - 测试1：Sheet名 `2-20260112` 解析为阶段WBS=2，日期=2026-01-12
  - 测试2：含"双周"的任务名 → node_type='双周'
  - 测试3：含"阶段"的任务名 → node_type='阶段'，且触发里程碑

**完成标志**：导入样例文件不报错，DB 里能查到 projects/phases/tracking_nodes/wbs_tasks 数据。

---

## 阶段五：报告模板

**目标**：为每份报告的每个叙述段落提供默认文字（模板 + 动态变量填充）。

- [ ] `app/services/report_templates.py`
  - 每类报告（measure_phase / progress_phase / progress_milestone / meeting / ...）的文字模板
  - `generate_all_narratives(project_id, db)`：遍历所有节点，为每份报告写入 narratives 表
  - `get_cumulative_workload(phase_id, db) -> float`：里程碑报告用，查询所有前序阶段的累计实际工作量
    - 查所有 end_date ≤ 当前阶段 end_date 的前序阶段
    - 取每个前序阶段最后一个节点的 task_node_data.actual_complete 求和

**完成标志**：导入后 narratives 表有数据，每个节点+报告类型都有对应记录。

---

## 阶段六：Excel 生成器

**目标**：给定 phase_id，生成对应的 .xlsx 文件并返回 bytes。

这是最复杂的部分，严格对照 `02.需求分析阶段.xlsx` 样例文件。

- [ ] `app/services/excel_builder/_primitives.py`（直接参考 EstimateM 的写法）
  ```python
  def _w(ws, r, c, v): ws.cell(row=r, column=c).value = v
  def _mc(ws, r1, c1, r2, c2): ws.merge_cells(...)
  def _wmc(ws, r1, c1, r2, c2, v): _mc(...); _w(...)
  ```

- [ ] `app/services/excel_builder/workreport.py`（write_only 模式）
  - 动态列：从阶段起始日到跟踪节点日期，每天一列
  - 行结构：Row1-3表头，Row4+按人员分组的任务行
  - G/H/K列：Python预计算（实际开始/完成/总工时），写值不写公式

- [ ] `app/services/excel_builder/progress_tracking.py`（write_only 模式）
  - 固定列A-H + 每节点6列
  - 偏差>20%标红

- [ ] `app/services/excel_builder/measure_report.py`（普通模式）
  - `build_measure_report(ws, data, variant)`  variant: '双周'/'阶段'/'里程碑'
  - 阶段1特殊格式（5区块）单独处理

- [ ] `app/services/excel_builder/progress_report.py`（普通模式）
  - `build_progress_report(ws, data, variant)`  variant: '双周'/'阶段'/'里程碑'
  - 阶段版：含利益相关方/资源/培训节；里程碑版：含QA/CM节（两套不同结构）
  - 阶段1特殊格式单独处理

- [ ] `app/services/excel_builder/meeting_sheet.py`（普通模式）
  - 常规例会单
  - 阶段1专属：LH001启动例会、LH002任务书研讨例会

- [ ] `app/services/excel_builder/__init__.py`
  - `async def build_phase_excel(phase_id, db) -> bytes`：主入口，按节点顺序创建所有Sheet

**完成标志**：调用 `build_phase_excel` 能生成 .xlsx，用 Excel 打开格式基本正确。

---

## 阶段七：API 路由

**目标**：前端能调用接口导入文件、编辑叙述、触发生成。

- [ ] `app/routers/projects.py`
  - `POST /api/import`：接收两个文件，调用 importer，返回 project_id 和导入摘要
  - `GET /api/projects/{id}/phases`：返回阶段列表及各阶段跟踪节点
  - `GET /api/narratives/{node_id}`：返回某节点所有报告的叙述文字
  - `PUT /api/narratives/{node_id}/{report_type}/{section}`：更新某段叙述
  - `POST /api/narratives/{node_id}/reset`：重置为模板默认值

- [ ] `app/routers/export.py`（SSE）
  - `GET /api/export/{project_id}`：SSE 端点
    - `asyncio.run_in_executor` 跑每个阶段的 build_phase_excel
    - 每完成一个阶段推送事件：`data: {"phase": "需求分析", "file": "/download/xxx.xlsx"}\n\n`
  - `GET /api/download/{filename}`：下载单个 xlsx
  - `GET /api/download-all/{project_id}`：打包 zip 下载

**完成标志**：用 curl 或 Postman 能调通导入和 SSE 生成接口。

---

## 阶段八：前端 UI

**目标**：Alpine.js 4-Tab 界面，能完成完整的导入→编辑→生成→下载流程。

- [ ] `app/static/index.html`（单文件，Alpine.js CDN）
  - Tab 1（导入）：两个文件上传按钮 + 导入状态表格
  - Tab 2（阶段配置）：展示阶段/节点树，配置例会地点/记录员
  - Tab 3（叙述编辑）：按阶段→节点→报告类型展开，textarea 编辑，重置按钮
  - Tab 4（生成下载）：生成按钮，EventSource 接收 SSE，逐个出现下载链接，打包下载

**完成标志**：浏览器完整跑通：上传文件→看到导入摘要→编辑叙述→生成→下载 xlsx 能用 Excel 正常打开。

---

## 阶段九：测试与收尾

- [ ] `tests/test_report_numbering.py`：验证多阶段多节点的三类报告编号按日期顺序正确递增
- [ ] 对比生成结果与 `02.需求分析阶段.xlsx` 样例：关键 Sheet 的单元格值和格式
- [ ] 更新 `CLAUDE.md` 中的「当前阶段」为编码完成

---

## 快速参考

| 文件 | 依赖 |
|------|------|
| database.py | 无 |
| cn_holidays.py | 无 |
| workhour_engine.py | cn_holidays.py |
| importer.py | database.py, report_templates.py |
| report_templates.py | database.py |
| excel_builder/ | workhour_engine.py, report_templates.py |
| routers/ | 所有 services |
| index.html | routers（API接口） |

**样例文件权威参照**：`/Users/gaoyiman/Desktop/Code/5000B/MeasureM/02.需求分析阶段.xlsx`  
**EstimateM 参考**：`/Users/gaoyiman/Desktop/Code/5000B/EstimateM/app/services/excel_builder.py`
