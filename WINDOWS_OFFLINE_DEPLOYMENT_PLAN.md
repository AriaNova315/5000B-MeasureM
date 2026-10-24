# MeasureM Windows 单机离线部署交付方案

日期：2026-10-24
状态：检查完成，方案待用户批准；尚未实施。
基线：当前工作区，HEAD `606abf21b24fc166f565f6f19f01ff8f93ce86b2`，包含用户原有未提交修改。

## 1. 简明结论与范围

推荐采用“源码 ZIP + 固定版本 Python 完整安装程序 + Windows wheelhouse + 部署说明”的方式。每台电脑使用独立虚拟环境、独立 SQLite 文件，只监听 `127.0.0.1:8002`，浏览器本机访问。数据库首次启动自动创建，不携带已有数据库、项目数据或历史导出。

已有私有仓库 `AriaNova315/5000B-MeasureM` 可作为交付仓库沿用，无需初始化 Git 或创建仓库；待用户确认归属并批准实施后，检查、提交并推送获准纳入的代码，始终保持私有。

当前不能直接宣称已具备离线交付条件：主页 Alpine.js 使用 CDN，API 文档默认引用外网资源；启动器使用全网卡监听和热重载；依赖未锁定；路径依赖启动目录；Windows 文件名、文件占用等行为未验证；关键测试依赖被忽略的样例文件。当前 macOS 检查不能代替 Windows 验收。

最小方案复用 FastAPI、Uvicorn、aiosqlite、openpyxl 现有流程，不新增服务层、部署平台或数据库迁移框架。数据流如下：

```text
联网 Windows（部署人员）
  Python 完整安装程序 + 已审核源码/锁定清单 -> 下载 Windows wheels
  -> 干净虚拟环境安装验证 -> 转运交付介质
离线 Windows（每台独立）
  安装相同 Python -> 创建 .venv -> 仅从 wheelhouse 安装
  -> 本机启动 -> 新建空 SQLite -> 浏览器上传本机 Excel
  -> 本机导入/叙述编辑 -> 生成 Excel -> 本机下载/Excel 重算
```

本阶段仅新增本方案文档。未修改业务代码、依赖、启动器或 Git 配置；未安装依赖、启动服务、初始化业务数据库、运行导入导出、提交、推送、创建仓库或操作 Actions。技能中涉及自动写状态、同步、提交或逐项交互的流程，按用户本次明确边界跳过；采用其架构、代码质量、测试、性能检查框架组织方案，不标记为已获实施批准。

## 2. 已检查资料和当前状态

已阅读 `AGENTS.md`、`CLAUDE.md`、`DEV_PLAN.md`、既有架构设计文档、全部运行模块的依赖/路径/平台相关逻辑及相关测试。`DEV_PLAN.md` 的未勾选任务属于历史计划，不能据此判断功能未完成。旧设计中“全部写值不写公式”“SSE 导出”等描述与实际代码不同，本方案以当前实现和较新的明确约定为准，不据旧文档改业务。

| 检查项 | 结果 | 状态与影响 |
|---|---|---|
| 项目边界 | Python/FastAPI/SQLite/Alpine.js/openpyxl；端口 8002 | 已满足基本架构；保持 Excel 结构和业务规则 |
| 独立运行 | `app/` 内未发现读取 ConfigM/EstimateM 数据库或导入同级源码 | 已满足；文档里的同族项目绝对路径仅为参考 |
| 启动入口 | `run.py:4`：`host="0.0.0.0", port=8002, reload=True` | 必须收敛到回环地址，关闭热重载，单进程 |
| Python 声明 | 没有 pyproject/版本约束；代码包含 `date \| None` 等语法 | 从语法看至少 Python 3.10；不是依赖兼容承诺 |
| 当前检查环境 | macOS arm64，Python 3.14.3 | 仅用于静态检查，不是 Windows 交付环境 |
| 依赖清单 | requirements.txt 只有 6 项，均未指定版本 | 必须冻结经过 Windows 验证的直接及传递依赖 |
| 数据库 | main.py lifespan 调用 init_db；DDL 建 8 张业务表；兼容 ALTER 内嵌 database.py | 无需数据库种子；不得交付真实 measurem.db |
| 数据库连接 | 请求依赖 finally 关闭连接，初始化使用 async with | 未发现平台专用驱动；写锁/占用须实机验证 |
| 写入路径 | measurem.db、generated/ 相对当前工作目录 | 必须保证启动目录固定且可写；SQLite 同目录也须允许写日志文件 |
| 前端离线 | index.html:7 引用 alpinejs@3.x.x CDN | 必须本地化且固定版本 |
| API 文档 | FastAPI 默认 /docs、/redoc 使用 CDN；ReDoc 引用 Google Fonts | 推荐关闭两种文档 UI，保留本地 openapi.json |
| 外部业务服务 | app/ 未发现在线 API、自动下载、在线授权、远程数据库 | 已满足离线业务前提；仍需断网实测 |
| 文件资源 | builder 创建新 Workbook，模板文字/样式在 Python 中 | 不需要运行时加载根目录样例 Excel |
| 输入公式 | importer.py:116 使用 data_only=True；输入样例含 129 个公式 | 上传前需在表格软件重算保存，避免读到缺失/旧缓存 |
| 输出公式 | builder.py 设置自动重算；WorkReport 使用 ArrayFormula；AnalysisHelper 辅助列实现分析 | 生成不需要 Office；查看完整公式结果需要表格软件重算 |
| 忽略规则 | DB、generated/、temp/ 已忽略；*.xlsx 一律忽略 | 业务数据主要目录已排除，但也漏交付测试样例 |
| Windows 文件语义 | export.py:51 把阶段名直接拼为本地文件名并以 wb 打开 | 无非法字符/占用处理，存在跨平台边界风险 |
| 本机状态 | 20 个已跟踪文件修改，7 个未跟踪文件 | 必须保留并审核；不能直接把 HEAD 当成当前可用代码 |
| 验证现状 | 30 个 Python 文件通过 AST 语法检查；两份样例 ZIP/XML 可读 | 未做功能测试或 Windows 运行认证 |

当前环境中的依赖元数据：FastAPI 0.115.5、Uvicorn 0.32.1、aiosqlite 0.20.0、openpyxl 3.1.5、python-multipart 0.0.12、Pydantic 2.12.5、Starlette 0.41.3；未发现 sse-starlette、pytest 的已安装元数据。这些是本机事实，不应直接复制为最终锁定版本，也不构成兼容性或安全性结论。运行代码直接导入 Pydantic，冻结时需明确包含它。

## 3. 推荐环境及依据

| 项目 | 推荐 | 依据/限制 |
|---|---|---|
| 系统 | Windows 11 x64，使用普通本地账户 | 给出一个明确交付基线；Windows 10 等环境须另做对应版本验收 |
| Python | 标准 CPython 3.14 x64；候选 3.14.9 | 与本机检查环境同一 minor；官方有完整 64 位安装程序；最终仍须验证 wheels 和功能 |
| Python 安装包 | python-3.14.9-amd64.exe 完整安装程序，包含 pip、venv | 不使用联网安装管理器、Microsoft Store 安装或 embeddable ZIP |
| 架构 | x64，普通 GIL 构建 | 不把 ARM64、x86 或自由线程构建混入同一 wheelhouse |
| 浏览器 | Windows 自带且可用的 Edge，或已安装的 Chrome | 本地访问 http://127.0.0.1:8002；无需浏览器在线服务 |
| 表格软件 | 程序运行不要求；建议使用已授权且可离线运行的桌面 Excel | 导入文件缓存、输出公式重算和版式验收需要它；Excel 2016 为代码兼容目标，实际可用版本须验收 |
| Word | 不需要 | 未发现 Word 自动化或 docx 处理 |
| 数据库软件 | 无需单独安装 SQLite 服务 | Python sqlite3 + aiosqlite |
| 其他 | 不要求 Git、Node.js、Docker、编译工具、IIS、Nginx 或 Windows 服务管理组件 | 目标机使用 ZIP 源码和完整 wheels，手动启动 |
| 安装目录 | 本地磁盘用户可写目录，例如 %LOCALAPPDATA%\MeasureM | 避免 Program Files、只读介质、UNC/网络共享和云同步目录 |

部署人员必须在与目标机相同 Python minor/架构的联网 Windows 上准备 wheels，使用完全相同的补丁版做最终验收并记录。若有 wheel 缺失，不在离线目标机临时编译，也不静默换版本；回到准备机调整候选依赖并重新验证。3.13 等备用基线仅在 3.14 无法通过验证时重新形成明确版本组合。

截至检查日期，Python 官方列出了 3.14.9 完整 Windows x64 安装程序；选择依据是交付可获得性与本项目环境连续性，不代表“已在 Windows 测过”。参考：[Python 3.14.9](https://www.python.org/downloads/release/python-3149/)、[Windows 下载列表](https://www.python.org/downloads/windows/)。

## 4. 离线依赖和业务边界

### 4.1 可直接本地化的资源

- Alpine.js：采用批准后确定的固定 3.x 版本，下载对应 dist/cdn.min.js 至 app/static/vendor/，修改页面为同源 /static/vendor/...；一并记录版本、来源、SHA-256 和许可证。目标机不下载脚本。
- API 文档 UI：推荐 `FastAPI(..., docs_url=None, redoc_url=None)`，保留 /openapi.json。若确需浏览器交互 API 文档，再选择本地化 Swagger 静态资源，这属于可选工作。
- 页面 CSS 内嵌，字体为 Microsoft YaHei/sans-serif 系统字体，没有在线字体资源；Excel 使用宋体/Arial 等字体名称，不需要交付字体文件，但查看机字体会影响显示/打印。
- requirements 的 wheels 由部署人员准备；必须包含 Pydantic Core 等平台相关传递依赖。uvicorn[standard] 在 Windows 的依赖由环境标记决定，不复制 macOS 的 uvloop 或 .venv。
- 现有 sse-starlette 未发现被运行代码使用。为控制变动，首轮可保留并一并锁定，不因旧设计提到 SSE 就新增 SSE 实现。

FastAPI 官方说明了 CDN 和自托管文档资源方式：[文档 UI 本地资源](https://fastapi.tiangolo.com/how-to/custom-docs-ui-assets/)。

### 4.2 需要明确的业务限制

1. **工时确定性**：workhour_engine.py:55 使用 `random.Random(hash(wbs + person_name))`。两个全新 Python 进程对同一字符串的 hash 已得到不同值，证明不能承诺跨重启、跨机器或跨 Python 版本完全一致；同一进程内节点复用仍是当前设计。项目约定禁止随意改种子，因此本方案不擅自替换算法或设置 PYTHONHASHSEED。推荐本轮部署保持现有逻辑并明确限制；如果要求跨重启一致，应先确认，再单独选择固定启动环境或稳定摘要种子的方案，验证生成数据影响。
2. **假日年份**：cn_holidays.py 内置 2024—2026 年假日/调休，其他年份实际退化为周一至周五判断，未发现旧设计所说的超范围 Warning。推荐本轮明确准确假日日历覆盖范围，不自动联网拉取。若交付业务包含其他年份，须先确认并离线审核补充日历；不能把部署验收当成日历正确性背书。
3. **公式缓存和查看**：openpyxl 不计算公式。输入样例有公式，data_only=True 会读取已保存的缓存；输出公式也需要 Excel 打开重算。禁止为省去 Office 而直接删除公式或改成常量。WPS/LibreOffice/网页版不能默认视为完全兼容，若使用需单独验收。参考：[openpyxl 公式说明](https://openpyxl.readthedocs.io/en/stable/simple_formulae.html)。

上述事项涉及业务结果，按 AGENTS.md 的冲突处理要求等待确认，当前不修改相关逻辑。Python hash 行为参考：[PYTHONHASHSEED](https://docs.python.org/3/using/cmdline.html#envvar-PYTHONHASHSEED)。

## 5. Windows 兼容性检查与最小修改清单

### 5.1 必须项（批准后才实施）

| 编号 | 文件/资源 | 原因及建议 | 预期影响 |
|---|---|---|---|
| M1 | run.py；新增 start_windows.bat | 回环监听 127.0.0.1，8002，reload=False，单进程；批处理以自身目录定位、检查 .venv\Scripts\python.exe、调用 run.py 并保留错误窗口 | 不开放局域网访问，减少重载子进程和重复启动；业务路由不变 |
| M2 | start_windows.bat；部署说明 | 通过 cd /d "%~dp0" 固定工作目录，检查写权限；从其他目录运行或双击都使用同一项目根目录 | 现有 DB_PATH/generated/ 静态路径保持兼容；无需引入路径配置框架 |
| M3 | app/static/index.html；新增 app/static/vendor/Alpine 固定资源及许可证 | 消除主界面 CDN，固定资源版本 | 断网且无浏览器缓存仍能使用全部界面 |
| M4 | app/main.py | 关闭默认 /docs、/redoc；保留 OpenAPI JSON | 消除隐藏 CDN/字体请求，主界面/API 功能不变 |
| M5 | requirements.txt；新增 requirements-windows.lock.txt | 选择经验证的直接依赖版本（含显式 Pydantic），锁定全部传递依赖；lock 仅包含包名与精确版本、无 URL/本机路径 | 后续由统一 lock 生成 wheelhouse 和安装；不是对本机 pip freeze 的照搬 |
| M6 | app/routers/export.py；针对性测试 | 在写文件前校验 Windows 不支持的字符、控制字符、目录分隔/路径穿越、长度边界；合法现有名称保持不变，非法名称明确报错；捕获文件占用/无写权限并返回可操作提示 | 避免合法 Linux/macOS 名称在 Windows 失败或写出目录；不擅自重命名业务阶段/Sheet |
| M7 | .gitignore；新增 examples/ 或 tests/fixtures/ 下的获准资源 | 补齐 SQLite -wal/-shm/-journal 等派生文件、交付构建产物的忽略；对审核后的必要 xlsx 做精确白名单 | 既排除业务数据，又不丢运行验证资源；不删除任何现有文件 |
| M8 | 新增 WINDOWS_DEPLOYMENT.md；验收记录；最少量部署回归测试 | 明确联网准备/离线安装/首次启动/故障处理/验收；覆盖修改引入的风险 | 部署人员可按同一套命令复现结果 |

M2 选择“固定启动目录”满足最小变动原则；不必同时改 main.py、database.py、export.py 的全部路径为绝对路径。若将来要求任意 ASGI 命令在任意工作目录运行，再统一为基于 __file__ 的路径，这不是本轮承诺。

Windows 路径使用 os.path/Path 和完整引号，避免 shell 字符转义问题；启动脚本采用 Windows 可识别的 CRLF，脚本命令尽量 ASCII，Python/HTML/说明统一 UTF-8。环境使用 `python`/指定解释器路径，不能直接照抄 macOS 的 python3 命令。默认通过显式 .venv 解释器运行，无需 PowerShell Activate.ps1，也无需降低执行策略。

### 5.2 可选项及不扩展范围

- **导出原子写入**：当前直接覆盖同路径。可选使用同目录临时文件、关闭句柄后 os.replace，并处理 Windows 目标被 Excel 占用；首轮不引入复杂文件队列。若实施必须验证失败不破坏旧文件。
- **减少依赖**：验证后可从 uvicorn[standard] 改为基础 uvicorn、移除未用 sse-starlette，减少 wheels；首轮优先保持清单语义，不作为交付前提。
- **可诊断日志/安装脚本**：可加最简单日志说明和离线安装批处理；无需服务化或自动升级。
- **数据库迁移异常处理**：init_db 的 ALTER 捕获所有 Exception，建议将来细化；全新交付先验证真实表结构，不扩展历史迁移工程。
- **文件句柄**：导入用 BytesIO、输出用 BytesIO，实际落盘使用 with open；运行代码未发现命名临时文件、子进程、shell 命令、COM、AppleScript、POSIX 专用调用。未发现需要为路径兼容重写 importer 的证据。若 Windows 验证发现句柄问题，再做对应最小修复。
- **性能**：上传整文件/构建完整 Workbook 都在内存中，多阶段生成由前端逐阶段 fetch；单机规模先测代表性项目耗时和峰值内存。不新增缓存、后台队列、多 worker 或异步任务平台。
- 不做容器、EXE 打包、自动更新、Windows 服务、计划任务、登录权限系统、局域网多人访问、功能重构或新增 CI。

## 6. 交付资源清单

推荐区分“私有源码仓库”和“给目标机的干净交付包”；交付包由明确允许的文件构建，不能直接压缩当前工作目录。

| 类别 | 需要内容 | 处理方式 |
|---|---|---|
| 运行源码 | run.py、app/ 全部必要模块（含目前未跟踪的 progress_narratives.py） | 经过审核后纳入源码版本和交付包 |
| 静态资源 | index.html、本地 Alpine.js、对应许可证/来源记录 | 随代码交付，无运行时下载 |
| 数据定义/业务种子 | database.py 的 DDL 与兼容列逻辑；cn_holidays.py；report_templates.py；meeting_narratives.py；progress_narratives.py；excel_builder/ 样式/公式 | 资源在源码中，无需额外种子 SQL、预置 DB 或外部模板 |
| 依赖清单 | requirements.txt、Windows 精确 lock；若使用测试再单独列开发依赖 | 与验收记录一致，不把 pytest/httpx 当最终用户必装依赖 |
| 启动/部署 | start_windows.bat、WINDOWS_DEPLOYMENT.md、版本/commit 与 SHA-256 清单、验收记录 | 启动前不自动安装依赖、不访问网络 |
| 输入模板 | 经审核的空白字段模板及可测试的脱敏样例 | 原件含业务内容，不能未经确认分发；保留必需 Sheet/公式缓存 |
| 格式参照 | 02.需求分析阶段.xlsx 的获准版本，或脱敏副本 | 是维护/验收资源，运行不依赖；不随包分发现有真实输出 |
| 测试源码/fixtures | 私有仓库保存完整测试及可用 fixture | 运行包可选；开发交付/验收包必须能独立验证 |
| Python 安装程序 | 候选 python-3.14.9-amd64.exe，官方来源及校验信息 | 部署人员在联网 Windows 下载，不把二进制安装包提交 Git |
| wheelhouse | 对应 Windows x64/CPython 版本的完整 wheels，清单及哈希 | 部署人员准备与转运；不提交大体积 wheels 到源码 Git |

必须排除：measurem.db 及其日志/副本、generated/、temp/、全部历史业务导出和业务输入、本机 .venv/venv、__pycache__、pytest 缓存、.git/、.claude/、.gstack/、.DS_Store、凭据/.env、本机日志。AGENTS.md/CLAUDE.md/DEV_PLAN.md 可留在私有源码仓库；部署包是否包含按维护需要选择。0815.md、promptinit.md、QualityM_promptinit.md、temp.md 等过程文档不作为运行必需项，不默认纳入交付包；不删除原件。

检查时两份根目录样例均被 *.xlsx 忽略，Git 当前未跟踪。输入样例 6 个 Sheet：项目信息、项目进度计划、里程碑、配置项清单、基线组成、代码模块；输出样例 8 个 Sheet、469 个公式。两份均未发现 externalLinks 部件；这不等于已审核所有单元格/超链接的隐私和安全。测试中的固定计数/日期依赖当前样例，替换 fixture 时须同时调整期望并保留语义，不能只复制个空白模板就宣称测试可复现。

## 7. Git 与 GitHub 处理方式

只读核验结果：

- git rev-parse --show-toplevel 返回本项目绝对路径，根目录存在 .git，未误用父目录仓库。
- 当前 main 跟踪 origin/main；fetch/push URL 均为 https://github.com/AriaNova315/5000B-MeasureM.git。
- 本地 HEAD、缓存 origin/main 以及在线 git ls-remote 的远程 main/HEAD 均为 `606abf21b24fc166f565f6f19f01ff8f93ce86b2`，截至核验时没有已提交内容的领先/落后。
- gh repo view 返回 `isPrivate: true`、默认分支 main、当前认证身份 viewerPermission=ADMIN。仓库名与 MeasureM 对应，技术上可以直接沿用。
- 初始工作区 20 个已跟踪文件修改、7 个未跟踪文件；包括多项报告公式/样式/导入业务改动，而非只有部署内容。未跟踪的 progress_narratives.py 已被现有模块导入，若只打包 HEAD 会丢失当前功能。
- 未跟踪项：0815.md、QualityM_promptinit.md、promptinit.md、app/services/progress_narratives.py、tests/test_excel_visual_styles.py、tests/test_meeting_attendees.py、tests/test_workday_calendar.py。
- git ls-files 检查未发现当前受跟踪的 .db、.xlsx、generated/ 或 temp/ 文件。未全面审计旧提交里是否出现过已删除的敏感数据，正式推送前应只读检查历史和待提交内容；私有状态不能代替内容审查。

推荐：沿用该私有仓库，不重设 remote、不新建仓库、不改可见性。实施阶段先形成明确纳入文件清单，保留并审核原有修改；不能 git add -A 打包所有过程文件，也不能回退用户改动。待批准后按 /review、适用测试和 /ship 流程执行获准的提交/推送，禁用不属于本次授权的版本/CI 扩展；不 force-push。

需要用户确认：AriaNova315/5000B-MeasureM 是否就是最终交付归属。若用户指定其他组织或仓库，先明确 owner/repo/权限再做关联；当前没有新建仓库的必要。若将来确需新建，必须经授权以 private 创建并验证后才能首次推送，本轮不执行。

推送前检查远端现有 Actions 配置/工作流和触发条件；当前本地受跟踪文件未见 .github/workflows。用户本轮禁止触发 Actions，因此后续批准推送时也不能擅自新增或触发工作流；如发现相关触发条件，先说明再处理。

## 8. 部署说明拟包含的内容及关键命令

以下是批准实施后的说明草案，不是本阶段已执行的命令。所有 `requirements-windows.lock.txt`、start_windows.bat 等新增文件须先完成和验证。以下 PowerShell 命令从项目根目录执行。

### 8.1 联网 Windows：部署人员准备

1. 下载指定完整 Python x64 安装程序，保留 pip/venv，记录来源、签名/哈希、确切版本。不要只下载需要联网取运行时的管理器。
2. 使用同一 Python 和架构，获取审核后的源码包及 Windows lock。lock 应在干净 Windows 环境解析/验证，固定直接和传递依赖，不含 URL/VCS/本机路径。
3. 用 lock 下载完整 wheels（不用 --no-deps）；随后在另一干净 venv 中禁网安装验证。

```powershell
py -3.14 -c "import sys,struct; print(sys.version); print(struct.calcsize('P')*8)"
py -3.14 -m venv .prepare-venv
.\.prepare-venv\Scripts\python.exe -m pip --disable-pip-version-check download --only-binary=:all: --dest .\wheelhouse -r .\requirements-windows.lock.txt
Get-ChildItem .\wheelhouse\*.whl | Get-FileHash -Algorithm SHA256
Get-FileHash .\python-3.14.9-amd64.exe -Algorithm SHA256
```

上述下载命令在联网 Windows 上执行，不能以 macOS 下载结果替代；候选 Python 补丁版必须与记录一致。所有 wheels 必须下载成功且无源码包；wheelhouse 不足则准备阶段失败。一般无需额外升级 pip；若确需固定 pip 版本，必须把对应 pip wheel 纳入准备和离线验证，不能临时在线升级。

pip 官方支持 download 归集依赖，再使用 --no-index/--find-links 安装：[pip download](https://pip.pypa.io/en/stable/cli/pip_download/)、[离线安装说明](https://pip.pypa.io/en/stable/user_guide/#installing-from-local-packages)。

### 8.2 离线目标机：安装和运行

1. 检查完整安装程序、源码 ZIP、wheelhouse 与清单哈希；Python 安装保留 pip/venv，按用户权限安装。解压源码到新的本地可写目录，确认包里没有 measurem.db/历史 generated。
2. 不复制联网准备机的 .venv；每台目标机重新创建。遇到 Windows 应用别名/py 未安装，使用说明中确认的 Python 完整路径。
3. 使用以下命令离线安装。--no-index 防止查在线索引，--no-cache-dir 排除本机缓存掩盖缺包，--only-binary 防止目标机编译；lock 中不得含网络 URL。

```powershell
Set-Location "$env:LOCALAPPDATA\MeasureM"
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip --disable-pip-version-check install --no-index --no-cache-dir --only-binary=:all: --find-links=.\wheelhouse -r .\requirements-windows.lock.txt
.\.venv\Scripts\python.exe -m pip check
.\start_windows.bat
```

无需激活 venv。启动后浏览器打开 http://127.0.0.1:8002；如果不自动打开就手工输入。首次启动创建空 DB，页面无项目；用户自行上传新项目 Excel。关闭浏览器不会停止服务，应在服务窗口 Ctrl+C 正常退出，结束前不要直接断电。

部署说明还应包含：目录与可写位置、端口占用诊断（Get-NetTCPConnection -LocalPort 8002 或 netstat -ano | findstr :8002）、缺 wheel/错误架构、无法写文件、Excel 占用导出文件、输入公式缓存、日志位置/控制台信息、重复启动提示、备份与卸载说明。备份是目标机使用后的维护动作，不意味着迁移已有数据；卸载/覆盖数据必须另经用户同意。端口冲突时不杀未知进程、不改用 8000/8001。

## 9. 验证方案与失败验收

### 9.1 本阶段已执行的安全检查

- 只读检查 Git 路径、状态、远程、分支、commit；在线只读查询 remote HEAD/main 和仓库私有性。
- 对 run.py、app/、tests/ 共 30 个 Python 文件使用 ast.parse 检查语法，通过；未导入应用、未触发 lifespan。
- 用 zipfile/XML 只读检查两个样例的 Sheet、公式数与外链部件，无业务初始化或文件改写。
- 两个短时 Python 子进程比较同一字符串 hash，确认跨进程随机化限制；未运行项目算法或修改环境变量。
- 未运行 pytest，因为未安装，且本轮禁止安装；未启动 HTTP 服务、测试导入或生成 Excel。未在 Windows 上测试，未下载 wheels。

### 9.2 批准后的验证矩阵

| 层次 | 验证内容 | 通过条件 |
|---|---|---|
| 干净打包 | 从审核的最终源码清单构建包，检查未跟踪必需模块和资源完整性 | 无真实 DB/导出/凭据；脱离原工作区、同级项目也能运行 |
| 离线依赖 | Windows 新用户/干净 venv，禁网、无 pip cache，按上述命令安装 | 无缺包/编译/网络访问，pip check 通过，安装清单与 lock 一致 |
| 启动入口 | 中文/空格路径、从其他工作目录调用和双击批处理 | 使用正确 .venv 与固定根目录，127.0.0.1:8002 可用；无 0.0.0.0/[::] 监听 |
| 首次初始化 | 在隔离副本中无 DB 启动，检查 8 张业务表及必需列 | 列表为空，DDL 正确；重启不丢测试数据、不重复种子 |
| 路径与权限 | 普通用户可写路径；只读目录、端口已占用、缺 .venv | 有明确诊断；不在其他目录意外新建 DB，不擅自变端口 |
| 完整离线界面 | 无缓存浏览器禁网访问，查看 Network/控制台，切换功能区 | Alpine 加载成功；主页面资源和业务请求均同源；文档 UI 已关闭或完全本地化 |
| 核心业务链 | 用脱敏 fixture 导入、项目列表、阶段/节点、叙述编辑/保存/恢复、单阶段及全阶段生成下载、重复导入 | API 与页面状态一致，人工覆盖生效，下载文件完整 |
| Excel 结构 | openpyxl 重新打开生成文件；覆盖阶段1、双周、普通阶段、里程碑 | Sheet/辅助隐藏表、编号、列序、合并、字体、公式、日期范围保留；偏差格式不变 |
| Office 行为 | 目标 Excel 版本打开重算保存，再以 data_only=True 读关键结果 | 无修复提示、#NAME?/意外错误；关键缓存和版式符合既有规则 |
| 文件边界 | 正常中文名、非法 Windows 字符/目录分隔、长名、文件被 Excel 打开、无写权限、重复导出 | 合法名保持原样；错误清晰、无路径逃逸；若采用原子写入，失败不破坏旧文件 |
| 数据库/句柄 | 导入/导出后正常停机、再次启动、关闭后移动测试副本 | 无遗留进程/句柄或持续锁库；连接正常关闭 |
| 业务限制 | 同一进程重复生成的节点工时一致性，跨重启表现记录；覆盖范围外年份行为 | 与最终获准的确定性及日历规则一致，不偷偷改变种子 |
| 资源占用 | 代表性完整项目各阶段连续生成，记录时间、内存、文件体积 | 目标机可完成且未失去响应；不以 macOS 数据代替 Windows 指标 |
| 回归 | 在隔离副本和单独测试依赖环境运行 python -m pytest -q | 原有及针对性部署测试通过；测试依赖不用装到生产 venv |
| 仓库交付 | push 后只读核对 remote commit、private、文件清单；核对 ZIP 版本 | 远程包含获准源码/资源，仍为 private，交付包可追溯同一版本 |

测试数据流与主要失败点：

```text
启动脚本 -> 解释器/目录/权限检查 -> 单进程服务 -> 空库 DDL -> GET 项目列表
  | 缺环境/端口占用/无权限：清晰失败，停止，不更改其他系统
浏览器 -> 本地 Alpine -> 上传 fixture -> importer(data_only) -> SQLite commit
  | 缺字段/损坏文件/缓存缺失：按既有校验反馈，不遗留半次导入
  -> 叙述读取/覆盖/恢复 -> 导出 builder -> generated 写入 -> 下载
  | 非法名称/文件占用/磁盘不可写：反馈原因，保持输出兼容
  -> openpyxl 结构检查 -> Excel 重算保存 -> 关键结果/版式确认
```

现有测试覆盖导入列表、叙述有效内容、报告公式、会议逻辑、样式和工作日日历，部分尚未跟踪，且多数依赖根目录项目相关信息表.xlsx。没有现成的 Windows 离线安装、启动器、回环监听或浏览器离线资源测试，需补足上述验收；不得把 AST 通过写成测试全通过。真实页面验证阶段使用 /qa 或 /qa-only，并联合检查 API、模型、页面状态流转。

## 10. 建议执行顺序与阶段验收

| 阶段 | 工作 | 验收/停止条件 |
|---|---|---|
| P0 当前 | 只读检查与本方案 | 用户收到明确方案；停下等待批准 |
| P1 归属/基线 | 确认现有私有仓库、目标 Windows/Office 和样例授权；审核当前未提交修改 | 必需模块清单明确；原修改保留；业务限制有决定 |
| P2 最小改动 | M1—M8；按需补少量部署测试；不改工时/报告业务规则 | 代码 diff 范围可审阅；静态资源完全本地；无新增 CI |
| P3 Windows 依赖冻结 | 干净 Windows 解析候选依赖、形成精确 lock；部署人员下载完整安装程序与 wheels | x64/标准 CPython 组合明确，所有下载齐全、哈希记录齐全 |
| P4 干净离线验收 | 新目录/新 venv/无数据库，禁网安装启动并走完整链路 | 第9节相关项目通过；无 Windows 环境则明确“待验收”，不能提前签收 |
| P5 提交与推送 | 按获准范围 review/ship、运行检查、提交并推送现有私有仓库 | 远端 commit 对应交付版本、private 保持；不新增/触发未经允许的 Actions |
| P6 封包交付 | 从已审核版本构建干净源码包，附文档/清单/验收结果，由部署人员配套转运安装包和 wheels | 再解压确认包内无业务数据；最终目标机独立运行，不依赖原机器 |

实施建议串行完成：启动器/静态资源/依赖锁定/测试与文档彼此关联，当前工作区已有大量业务修改，没有必要引入多 worktree、并行代理或额外合并。若 Windows 验证发现业务规则和样例矛盾，暂停对应修改，先获得决定。

## 11. 待用户决定的信息

请集中确认以下事项，其他常规细节按本文推荐执行：

1. **仓库归属**：是否沿用现有 private 仓库 AriaNova315/5000B-MeasureM？推荐沿用。
2. **目标环境**：是否可按 Windows 11 x64 + 标准 Python 3.14 + 桌面 Excel 验收？若目标机实际系统/架构或表格软件不同，请给出实际版本；不需要另行逐项挑选库版本。
3. **样例与业务范围**：现有输入/输出样例是否获准放入私有源码库和交付包？推荐仅交付审核后的空白模板与脱敏 fixture，原件保留；同时明确是否需要处理 2024—2026 之外的项目年份。
4. **确定性要求**：是否接受本轮保持现有“同进程一致、跨重启不保证完全一致”的实现限制？推荐部署改动不改种子；若必须跨重启一致，先确认该业务要求，再形成单独的最小调整方案。
5. **实施批准**：是否批准按本方案进入后续实施？当前没有执行任何实施项；上述待决信息不自动视为已同意。

## GSTACK REVIEW REPORT

| Review | Trigger | Why | Runs | Status | Findings |
|---|---|---|---|---|---|
| Eng Review | /plan-eng-review | Windows 单机离线交付 | 1 | DONE_WITH_CONCERNS，待用户批准 | 已检查架构/代码质量/测试/性能；提出 M1—M8；Windows 实测和业务边界未决 |
| Outside Voice | 未执行 | 本轮只读检查与单一方案文档范围 | 0 | 未执行 | 未发送代码给外部模型，未生成额外评审产物 |

**VERDICT:** 方案已完成，不代表部署验收通过或获准实施。仅本方案文档为本阶段新增文件；不写技能状态/遥测/同步日志，不创建其他产物。

**UNRESOLVED DECISIONS:**
- 最终私有仓库归属、实际 Windows/Excel 环境、样例分发授权、项目年份覆盖、跨重启确定性要求及用户明确实施批准，见第11节。
