"""
高亮服务 — SQLite CRUD 操作。
"""
import json
from db import get_connection
from models import HighlightCreate, HighlightUpdate, HighlightResponse


def create_highlight(h: HighlightCreate) -> HighlightResponse:
    """创建新的高亮记录。"""
    segments_json = json.dumps(
        [s.model_dump() for s in h.segments], ensure_ascii=False
    )
    conn = get_connection()
    cursor = conn.execute(
        """INSERT INTO highlights (ppt_id, slide_idx, highlighted_text, segments_json, color, note)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (h.ppt_id, h.slide_idx, h.highlighted_text, segments_json, h.color, h.note),
    )
    conn.commit()
    row = conn.execute("SELECT * FROM highlights WHERE id=?", (cursor.lastrowid,)).fetchone()
    conn.close()
    return _row_to_response(row)


def get_highlights_for_slide(ppt_id: str, slide_idx: int) -> list[HighlightResponse]:
    """获取某页所有高亮。"""
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM highlights WHERE ppt_id=? AND slide_idx=? ORDER BY created_at ASC",
        (ppt_id, slide_idx),
    ).fetchall()
    conn.close()
    return [_row_to_response(r) for r in rows]


def update_highlight(highlight_id: int, h: HighlightUpdate) -> HighlightResponse:
    """更新高亮的颜色或笔记。"""
    conn = get_connection()
    existing = conn.execute("SELECT * FROM highlights WHERE id=?", (highlight_id,)).fetchone()
    if not existing:
        conn.close()
        raise ValueError(f"高亮 {highlight_id} 不存在")

    new_color = h.color if h.color is not None else existing["color"]
    new_note = h.note if h.note is not None else existing["note"]

    conn.execute(
        "UPDATE highlights SET color=?, note=?, updated_at=datetime('now') WHERE id=?",
        (new_color, new_note, highlight_id),
    )
    conn.commit()
    row = conn.execute("SELECT * FROM highlights WHERE id=?", (highlight_id,)).fetchone()
    conn.close()
    return _row_to_response(row)


def delete_highlight(highlight_id: int):
    """删除高亮记录。"""
    conn = get_connection()
    conn.execute("DELETE FROM highlights WHERE id=?", (highlight_id,))
    conn.commit()
    conn.close()


def _row_to_response(row) -> HighlightResponse:
    """将 sqlite3.Row 转为 Pydantic 响应模型。"""
    if row is None:
        raise ValueError("记录不存在")
    return HighlightResponse(
        id=row["id"],
        ppt_id=row["ppt_id"],
        slide_idx=row["slide_idx"],
        highlighted_text=row["highlighted_text"],
        segments_json=row["segments_json"],
        color=row["color"],
        note=row["note"],
        created_at=row["created_at"],
        updated_at=row["updated_at"],
    )
