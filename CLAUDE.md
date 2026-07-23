## 项目上下文（每次新会话必读）

### 这是什么项目
MeasureM：项目测量与监控记录生成系统。项目结束后，对照实际数据生成一套看起来足够真实、逻辑自洽的测量与监控阶段记录文件（按阶段生成多份 Excel）。
技术栈：Python + FastAPI + SQLite (aiosqlite) + Alpine.js + openpyxl

### 关键文档路径
- **样例输出文件（格式权威参照）**：`/Users/gaoyiman/Desktop/Code/5000B/MeasureM/02.需求分析阶段.xlsx`
- **样例输入文件（字段来源）**：`/Users/gaoyiman/Desktop/Code/5000B/EstimateM/项目相关信息表.xlsx`（共享同一份 Excel）
- **同族系统参考**：ConfigM（../ConfigM/），EstimateM（../EstimateM/）

### 当前阶段（更新于 2026-07-22）
**Phase 1-13 编码完成 —— 核心功能完整可用**

- ✅ 样例文件结构分析完毕（见本文件"样例文件结构详解"节）
- ✅ 需求讨论（/office-hours 完成）
- ✅ 技术方案设计（见设计文档）
- ✅ Phase 1：项目脚手架（run.py / main.py / database.py / requirements.txt）
- ✅ Phase 2：假日/工时引擎（cn_holidays.py / workhour_engine.py）
- ✅ Phase 3：数据导入服务（importer.py）
- ✅ Phase 4：报告模板工具（report_templates.py）
- ✅ Phase 5：Excel 构建器（excel_builder/）
- ✅ Phase 6：API 路由（projects.py / export.py）
- ✅ Phase 7：Alpine.js 前端 UI（index.html，4 功能区）
- ✅ Phase 8：端到端对比测试（与样例文件对比，已修正格式差异）
- ✅ Phase 9：WorkReport 多节点 + 算法修正 + ProgressTracking 双模式（2026-07-09）
- ✅ Phase 10：测量项分析自动读取 ProgressTracking 节点偏差数据（2026-07-15）
- ✅ Phase 11：例会摘要范围与全局流水号修正（2026-07-22）
- ✅ Phase 12：导入文件核对信息与已导入项目卡片完善（2026-07-22）
- ✅ Phase 13：例会自动内容与叙述编辑统一（2026-07-22）

**设计文档路径**：`~/.gstack/projects/MeasureM/gaoyiman-MeasureM-design-20260703-134845.md`

**当前状态：核心功能完整，可正常使用**

**Phase 8 修正内容（2026-07-07）**:
- Sheet 名称加入日期和类型标签，格式 `{seq}测量与分析报告{date}（{label}）`
- 里程碑进展报告去掉偏差统计/综合情况，改为"计划完成情况"段落
- 里程碑报告文档列表扩展至 7 份（与阶段版一致）
- ProgressTracking 增加阶段汇总行（WBS=阶段序号，最顶层）
- ProgressTracking 节点标题日期格式改为无前导零（YYYY.M.D）
- HTTP 响应中文件名正确 URL 编码（RFC 5987）
- Sheet 名称超 31 字符时自动截断

**Phase 9 修正内容（2026-07-09）**:

*算法层（workhour_engine.py / importer.py）*
- 实际已完成工作量采用**三态逻辑**（非工时反推）：已完成→plan_hours/1.0；未开始→0/0.0；进行中→最近5%倍数，|actual-plan_should|≤plan_should×20%
- 工时分布（WorkReport 日工时）与实际已完成工作量是**两个独立概念**，分别生成
- 非项目组成员过滤：从`项目信息`Sheet 解析`team_members`字段，资源列表中不在团队的人员不写入 WorkReport
- 内部成员工时按均分：`plan_hours / 内部成员数`

*WorkReport（excel_builder/workreport.py）*
- **每个跟踪节点生成一张 WorkReport**（不再是一张/阶段）
- 日期列显示全阶段范围，但工时仅填写到本节点的 `node_date`
- `hour_dists` 在 builder.py 中**预计算一次**（`hash(wbs+person_name)` 确定性种子），各节点 WorkReport 复用，保证同一日期工时跨 WorkReport 一致
- 列 G（实际开始时间）：ArrayFormula，找第一个有工时的日期
- 列 H（实际完成时间）：普通公式，仅 I 列=100% 时填最后有工时的日期
- 列 K（个人实际总工时）：`=SUM(L{r}:{last_col}{r})`
- 列 I（任务完成百分比）/ 列 J（实际已完成工作量）：来自**当前节点**的 tnd 数据（不再误用最后节点）
- 返回 `task_wr_rows: dict[int, list[int]]`（task_id → 行号列表），供 ProgressTracking 写公式

*ProgressTracking（excel_builder/progress_tracking.py）*
- 列 F（实际开始时间）/ 列 G（实际完成时间）：Excel 公式引用最后节点 WorkReport，格式 `=IF(COUNT(...)=0,"",MIN/MAX(...))`；无 WorkReport 数据时回退到计划日期
- **双周节点**：4 列（计划应完成工作量 | 完成百分比 | 实际已完成工作量 | 完成工作量偏差）
  - 完成工作量偏差 = (实际已完成-计划应完成)/计划应完成，百分比格式，非零时红色字体
- **阶段节点**：6 列（原有结构不变）

*builder.py*
- WorkReport 生成移至节点循环内，ProgressTracking 移至循环后（引用最后一个 WorkReport）

**Phase 10 修正内容（2026-07-15）**:
- 阶段版/里程碑版从阶段组“工作量偏差（阶段统计）”列收集工作量偏差任务及超过 20% 的任务
- 双周版从双周组“完成工作量偏差”列收集进度偏差任务；阶段版/里程碑版从阶段组“进度偏差（阶段统计）”列收集
- 规模区文档和代码模块按绝对偏差率超过 20% 自动列出
- 阶段一保留原有特殊测量报告逻辑；多项名称换行显示并动态调整行高

**Phase 11 修正内容（2026-07-22）**:
- 例会动态工作摘要只统计“软件工程活动”和“评审活动”下的叶节点任务
- 没有正在进行的任务时，不再输出“正在进行的工作：无”
- 例会流水号按前序阶段实际例会数量计算，阶段 2 项目策划会计入后续阶段编号
- 全项目 7 个阶段验证为 12 张例会，编号连续为 LH001-LH012

**Phase 12 修正内容（2026-07-22）**:
- 选择导入文件后显示文件名、大小和最后修改时间；明确浏览器不能读取本地完整路径
- 项目表新增来源文件名并兼容旧数据库自动迁移；重复导入时同步更新来源和导入时间
- 已导入项目显示项目名称、代号、来源、阶段/进度跟踪节点/任务规模和导入时间
- 项目卡片增加“叙述编辑”“生成导出”“删除”快捷操作
- 修正项目、阶段和节点列表的 Alpine.js `x-for` 模板结构，恢复已有数据渲染

**Phase 13 修正内容（2026-07-22）**:
- 例会自动内容提取为共享解析层，叙述编辑 API 和 Excel 生成器使用同一套默认值
- `narratives.content` 明确作为人工覆盖层：空值使用自动内容，非空值优先写入 Excel
- 阶段1例会拆分为 LH001 项目启动会、LH002 软件研制任务书研讨会、LH003 阶段例会
- LH003 按普通阶段例会生成，但不包含“项目组成员报告个人工作情况”议题
- 阶段2项目策划会纳入叙述编辑，后续双周/阶段例会按实际全局流水号显示
- 编辑字段标识“自动生成/人工修改”，支持恢复自动生成

**注意**：端口使用 **8002**（8000=ConfigM，8001=EstimateM）

### 项目结构（目标，待确认后更新）

```
MeasureM/
├── run.py                        # 启动入口：python3 run.py
├── measurem.db                   # SQLite 数据库
├── generated/                    # 生成的 Excel 文件
│   └── {project_id}/
├── app/
│   ├── main.py                   # FastAPI app + 路由注册
│   ├── database.py               # aiosqlite 连接 + DDL 初始化
│   ├── models.py                 # Pydantic 请求/响应模型
│   ├── routers/
│   │   ├── projects.py           # CRUD 项目信息
│   │   ├── phases.py             # 阶段管理（时间段、测量配置）
│   │   └── export.py             # 生成并下载 Excel
│   ├── services/
│   │   ├── importer.py           # 解析项目相关信息表.xlsx → DB
│   │   └── excel_builder.py      # openpyxl 生成多-Sheet Excel
│   └── static/
│       └── index.html            # Alpine.js 单页应用
└── requirements.txt
```

---

## 样例文件结构详解

样例文件：`02.需求分析阶段.xlsx`（需求分析阶段的完整测量记录）

### Sheet 1：WorkReport20250804（工时记录）
- 尺寸：133 行 × 83 列（A:CE）
- **行结构**：
  - Row 1：列名（A-K 各列标题，与 Row 2-3 三行合并）
  - Row 2：日期序列号（从 L 列起，每列一天，动态扩展）
  - Row 3：日期值（datetime，对应 Row 2 的日期）
  - Row 4+：数据行，按人员分组（A 列仅在每人第一行填写姓名，其余行 A 列为空）
- **固定列（A-K）**：姓名 | WBS | 任务名称 | 计划开始时间 | 计划完成时间 | 资源名称 | 实际开始时间 | 实际完成时间 | 任务完成百分比 | 实际已完成工作量 | 个人实际总工时
- **动态列（L 起）**：每天一列，列数 = 阶段日历天数（样例从 2025.7.9 到 2025.8.4，共 27 天 → 83-11=72 列偏多，含更大日期范围）
- **实际样例日期范围**：2025.7.9（L3）至阶段末，83-11=72 天列
- **人员分组**（样例含 11 人）：穆笑岩(row4)、张晓丰(13)、董薇(18)、李国荣(19)、李静静(20)、孙贵磊(21)、刁鹏(25)、方小娟(26)、李德馨(32)、高蕾(33)、王令敏(52)
- **字体**：宋体 10pt Bold（表头），宋体 10pt（数据）

### Sheet 2：ProgressTracking-ok（任务进度跟踪）
- 尺寸：46 行 × 动态列数（按节点类型累加）
- **行结构**：Row 1-2 = 双行表头（合并），Row 3 = 阶段汇总行，Row 4+ = 数据（WBS 层级树）
- **固定列（A-H）**：WBS | 任务名称 | 计划开始时间 | 计划完成时间 | 资源名称 | 实际开始时间 | 实际完成时间 | 计划总工作量
  - 列 F/G（实际开始/完成时间）：Excel 公式引用最后节点 WorkReport 的 G/H 列
- **动态列组（按节点类型）**：
  - 阶段节点（6 列）：计划应完成工作量 | 实际完成百分比 | 实际已完成工作量 | 实际工时 | 进度偏差（阶段统计）| 工作量偏差（阶段统计）
  - 双周节点（4 列）：计划应完成工作量 | 完成百分比 | 实际已完成工作量 | 完成工作量偏差（=(实际-计划应完成)/计划应完成，百分比格式，非零红色）
  - 节点列头（Row 1）：合并跨对应列数，标注"阶段跟踪/双周跟踪（YYYY.M.D）"
- **WBS 层级**：最顶层=阶段汇总行，再下为父节点汇总行，叶节点任务行
- **字体**：Arial 9pt Bold（表头），Arial 9pt（数据），Arial 9pt Red（双周偏差非零）

### Sheet 3/4：002测量与分析报告（阶段）/ 里程碑版
- 尺寸：~31 行 × 10 列（J 列）
- **报告编号格式**：`ZDPX-{项目代号}/0.0-MA-测分报告-{全项目顺序号}`
- **表头区（Row 1-5）**：标题"测量与分析报告" + 编号、项目名称/编号、项目周期/项目经理、所处阶段/报告时间段
- **测量项统计区（Row 6-23）**：
  - 规模：文档列表（序号 | 文档名称 | 计划规模 | 实际规模 | 偏差率）+ 代码模块列表
  - 工作量：计划 vs 实际 + 偏差率
  - 进度：计划完成工作量 vs 实际完成工作量 + 偏差
  - 需求：需求总数 + 变更数 + 变更率
- **测量项分析区（Row 24-31）**：各类偏差任务描述 + 异常原因分析 + 拟制人员
- **阶段报告 vs 里程碑报告的区别**：
  - 阶段：报告时间段 = 本次测量周期（如 2025.7.9-2025.8.4），工作量 = 当期统计
  - 里程碑：报告时间段 = 从阶段起始至里程碑（如 2025.6.30-2025.8.4），工作量含所有先前测量，文档列表多一行"软件研制任务书"（计划规模为"——"）
- **字体**：标题黑体 16pt Bold，正文宋体 11pt，表头宋体 11pt Bold

### Sheet 5/6：进展报告（阶段）/ 里程碑版
- 尺寸：~36 行 × 9 列（I 列，从 B 列起）
- **报告编号格式**：`ZDPX-{项目代号}/0.0-PMC-进展报告-{全项目顺序号}`
- **阶段版结构（7 节）**：
  1. 项目基本信息（名称/编号/周期/负责人/阶段/时间段）
  2. 项目进展情况（变更情况 + 进展叙述 + 偏差统计 + 综合情况）
  3. 风险问题说明
  4. 利益相关方参与计划跟踪情况
  5. 资源计划跟踪情况
  6. 人员培训计划跟踪情况
  7. 项目问题跟踪情况
  8. 本期工作总结 + 下期工作安排
  9. 拟制人员 + 领导意见
  10. 查看相关文档列表（文件名 + 日期）
- **里程碑版额外包含**：质量保证工作、配置管理工作（比阶段版多 2 节）
- **字体**：标题黑体 16pt，正文宋体 11pt

### Sheet 7/8：例会单（LH{N}）
- 尺寸：~11-13 行 × 8 列（A:H）
- **Sheet 名称**：`LH{序号}项目例会单`（序号为本阶段内的会议顺序，从 Excel 文件内递增）
- **报告编号格式**：`ZDPX-{项目代号}/0.0-PMC-工作例会-{全项目顺序号}`
- **固定结构**：标题"项目例会单" | 编号行 | 时间/地点 | 主持人/记录员 | 参加人员 | 主要内容（多行议题）
- **主要内容**：多条 "N. 议题文字" 格式，每条占一行

### 关键编号规律（待 office-hours 确认）
- **报告序号**是全项目累计序号，跨阶段递增
- 样例阶段内：测分报告002(阶段)、003(里程碑)；进展报告002(阶段)、005(里程碑)；例会单004、006
- 用户需要在每个阶段指定各类报告的起始编号

---

## 已知规则和约定

**启动方式**
```bash
python3 run.py   # 系统 Python，不使用 venv，端口 8002
```

**独立性原则**
- 不读取 ConfigM/EstimateM 数据库
- 从同一份 `项目相关信息表.xlsx` 自行解析所需 Sheet
- MeasureM 复用的 Sheet：`项目信息`（基本字段）、`项目进度计划`（WBS/任务/资源/工作量）

**数字格式**
- 偏差率保留小数，不转百分比（与样例一致：`-0.085714...`）
- 日期在 Excel 中以序列号存储，openpyxl 写入时使用 `datetime.date` 对象

**阶段1（系统分析与设计）特殊规则**
- WorkReport 工时总和必须精确等于实际已完成工作量，**禁止任何偏差**（其他阶段允许 ±20%）
- 额外生成两张专属例会单：LH001 项目启动例会（全员）、LH002 项目任务书研讨例会（核心组），编号从 001 起；常规阶段例会为 LH003
- 测量分析报告仅 1 份（无阶段/里程碑区分），结构仅 5 区块（无进度/需求/偏差分析）
- 进展报告仅 1 份（无阶段/里程碑区分），无利益相关方/资源/培训/QA/CM 追踪节
- 详细格式见设计文档第九节

---

## 里程碑更新规则
**每完成一个功能模块后，立即更新本文件：**

1. 更新「当前阶段」描述和日期
2. 打 ✅ 标记完成项
3. 更新「下一步」
4. 补充编码过程中发现的新规则

---

## Skill routing

When the user's request matches an available skill, invoke it via the Skill tool. When in doubt, invoke the skill.

Key routing rules:
- Product ideas/brainstorming → invoke /office-hours
- Strategy/scope → invoke /plan-ceo-review
- Architecture → invoke /plan-eng-review
- Design system/plan review → invoke /design-consultation or /plan-design-review
- Full review pipeline → invoke /autoplan
- Bugs/errors → invoke /investigate
- QA/testing site behavior → invoke /qa or /qa-only
- Code review/diff check → invoke /review
- Visual polish → invoke /design-review
- Ship/deploy/PR → invoke /ship or /land-and-deploy
- Save progress → invoke /context-save
- Resume context → invoke /context-restore
