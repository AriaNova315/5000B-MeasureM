import aiosqlite

DB_PATH = "measurem.db"

DDL = """
CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    code TEXT,
    manager TEXT,
    team_members TEXT,  -- JSON数组
    qa_person TEXT,
    cm_person TEXT,
    source_filename TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS phases (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL REFERENCES projects(id),
    wbs TEXT NOT NULL,       -- '1', '2', '3'...
    name TEXT NOT NULL,
    start_date TEXT,         -- YYYY-MM-DD
    end_date TEXT
);

CREATE TABLE IF NOT EXISTS tracking_nodes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    phase_id INTEGER NOT NULL REFERENCES phases(id),
    date TEXT NOT NULL,      -- YYYY-MM-DD
    node_type TEXT NOT NULL  -- '双周' | '阶段'
);

CREATE TABLE IF NOT EXISTS wbs_tasks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    phase_id INTEGER NOT NULL REFERENCES phases(id),
    wbs TEXT NOT NULL,
    name TEXT NOT NULL,
    start_date TEXT,
    end_date TEXT,
    plan_hours REAL DEFAULT 0,
    resources TEXT DEFAULT '[]'  -- JSON数组
);

CREATE TABLE IF NOT EXISTS task_node_data (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_id INTEGER NOT NULL REFERENCES wbs_tasks(id),
    node_id INTEGER NOT NULL REFERENCES tracking_nodes(id),
    plan_should_complete REAL DEFAULT 0,
    pct_complete REAL DEFAULT 0,
    actual_complete REAL DEFAULT 0,
    UNIQUE(task_id, node_id)
);

CREATE TABLE IF NOT EXISTS narratives (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tracking_node_id INTEGER NOT NULL REFERENCES tracking_nodes(id),
    report_type TEXT NOT NULL,
    section TEXT NOT NULL,
    content TEXT DEFAULT '',
    UNIQUE(tracking_node_id, report_type, section)
);

CREATE TABLE IF NOT EXISTS config_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL REFERENCES projects(id),
    seq INTEGER,
    name TEXT NOT NULL,
    identifier TEXT,
    version TEXT,
    item_type TEXT,       -- '文档' | '代码' | '其他'
    actual_pages INTEGER, -- 实际页数（文档类）
    plan_pages INTEGER,   -- 估算规模，默认=actual_pages
    phase_wbs TEXT,       -- 所属阶段 WBS
    review_date TEXT      -- 评审完成日期 YYYY-MM-DD
);

CREATE TABLE IF NOT EXISTS code_modules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL REFERENCES projects(id),
    seq INTEGER,
    name TEXT NOT NULL,
    actual_loc INTEGER,   -- 实际行数
    plan_loc INTEGER,     -- 估算行数，默认=actual_loc
    phase_wbs TEXT,       -- 所属阶段 WBS
    completion_date TEXT  -- 完成日期 YYYY-MM-DD
);
"""


async def init_db():
    async with aiosqlite.connect(DB_PATH) as db:
        await db.executescript(DDL)
        # 兼容旧数据库：新列已存在时 SQLite 会报错，忽略即可
        for sql in [
            "ALTER TABLE projects ADD COLUMN source_filename TEXT",
            "ALTER TABLE config_items ADD COLUMN review_date TEXT",
            "ALTER TABLE code_modules ADD COLUMN phase_wbs TEXT",
            "ALTER TABLE code_modules ADD COLUMN completion_date TEXT",
        ]:
            try:
                await db.execute(sql)
            except Exception:
                pass
        await db.commit()


async def get_db() -> aiosqlite.Connection:
    db = await aiosqlite.connect(DB_PATH)
    db.row_factory = aiosqlite.Row
    return db
