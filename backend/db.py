"""
SQLite 数据库初始化和 CRUD 辅助函数。
自动创建 ppt_reader.db 和所需的表。
"""
import sqlite3
import os
from paths import data_dir

DB_PATH = os.path.join(data_dir(), "ppt_reader.db")


def get_connection():
    """获取数据库连接，启用 WAL 模式和外键约束。"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    """创建所有需要的表（如果不存在）。"""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS presentations (
            id TEXT PRIMARY KEY,
            filename TEXT NOT NULL,
            original_filename TEXT NOT NULL,
            slide_count INTEGER NOT NULL DEFAULT 0,
            title TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS highlights (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ppt_id TEXT NOT NULL REFERENCES presentations(id) ON DELETE CASCADE,
            slide_idx INTEGER NOT NULL,
            highlighted_text TEXT NOT NULL,
            segments_json TEXT NOT NULL,
            color TEXT DEFAULT '#FFEB3B',
            note TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        )
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_highlights_ppt_slide
        ON highlights(ppt_id, slide_idx)
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS chat_messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ppt_id TEXT NOT NULL REFERENCES presentations(id) ON DELETE CASCADE,
            role TEXT NOT NULL CHECK(role IN ('user', 'assistant')),
            content TEXT NOT NULL,
            created_at TEXT DEFAULT (datetime('now'))
        )
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_chat_ppt ON chat_messages(ppt_id)
    """)

    conn.commit()
    conn.close()
