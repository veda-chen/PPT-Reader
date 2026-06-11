"""
高保真渲染服务 — 调用 export_worker.py 子进程，把 PPT 导出为高清 PNG。

设计：
  - 上传时一次性导出全部页，缓存到 uploads/<ppt_id>/render/slide{N}.png
  - 子进程隔离 COM，超时/失败时降级（前端回退到 HTML 文字渲染）
  - 内存记录每个 ppt 的导出状态和原始像素尺寸（用于透明文字层定位）
"""
import os
import sys
import json
import asyncio
from paths import data_dir

UPLOADS_DIR = os.path.join(data_dir(), "uploads")
# WORKER 仅 dev 模式使用（export_worker.py 在 backend/ 下）；
# 打包(frozen)模式改用 sys.executable + "--export-worker" 入口，见 _run_export_subprocess。
WORKER = os.path.join(os.path.dirname(__file__), "export_worker.py")

EXPORT_ZOOM = 2.0  # PNG 渲染倍率

# 内存状态: { ppt_id: {"ok": bool, "count": int, "page_width": ..., "page_height": ...} }
_render_status: dict[str, dict] = {}
# 文字层缓存: { ppt_id: {"page_width": W, "page_height": H, "slides": [[span,...],...]} }
_spans_cache: dict[str, dict] = {}


def render_dir(ppt_id: str) -> str:
    return os.path.join(UPLOADS_DIR, ppt_id, "render")


def has_rendered(ppt_id: str) -> bool:
    """该 PPT 是否已成功导出高清图。"""
    return _render_status.get(ppt_id, {}).get("ok", False)


def get_render_info(ppt_id: str) -> dict:
    """返回导出信息（含页面 point 尺寸，用于坐标换算）。"""
    return _render_status.get(ppt_id, {"ok": False})


def slide_image_path(ppt_id: str, slide_idx: int) -> str | None:
    """返回某页 PNG 的磁盘路径（不存在返回 None）。"""
    path = os.path.join(render_dir(ppt_id), f"slide{slide_idx}.png")
    return path if os.path.exists(path) else None


def _load_spans(ppt_id: str) -> dict:
    """加载并缓存某 PPT 的文字层 JSON。"""
    if ppt_id in _spans_cache:
        return _spans_cache[ppt_id]
    path = os.path.join(render_dir(ppt_id), "spans.json")
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        _spans_cache[ppt_id] = data
        return data
    except Exception:
        return {}


def get_slide_spans(ppt_id: str, slide_idx: int) -> dict:
    """
    返回某页的精确文字层数据。
    { "page_width": W_pt, "page_height": H_pt, "spans": [ {text,x,y,w,h,size,block,line,span}, ... ] }
    无数据时返回空 spans。
    """
    data = _load_spans(ppt_id)
    if not data:
        return {"page_width": 960, "page_height": 540, "spans": []}
    slides = data.get("slides", [])
    spans = slides[slide_idx] if 0 <= slide_idx < len(slides) else []
    return {
        "page_width": data.get("page_width", 960),
        "page_height": data.get("page_height", 540),
        "spans": spans,
    }


def _run_export_subprocess(pptx_path: str, out_dir: str, timeout: int) -> dict:
    """
    同步调用子进程导出（在独立线程中运行，避免阻塞事件循环）。
    用 subprocess.run 而非 asyncio 子进程——后者在 Windows SelectorEventLoop 下不支持。
    """
    import subprocess
    # 打包(frozen)后 sys.executable 是 exe 本身、WORKER 脚本不可直接执行，
    # 改用 exe 的 "--export-worker" 入口（仍是独立进程，隔离 COM）。
    if getattr(sys, "frozen", False):
        cmd = [sys.executable, "--export-worker", pptx_path, out_dir, str(EXPORT_ZOOM)]
    else:
        cmd = [sys.executable, WORKER, pptx_path, out_dir, str(EXPORT_ZOOM)]
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            timeout=timeout,
        )
        out = proc.stdout.decode("utf-8", errors="replace").strip()
        last_line = out.splitlines()[-1] if out else "{}"
        try:
            result = json.loads(last_line)
        except Exception:
            result = {"ok": False, "error": f"解析失败: {out[:200]}"}

        if not result.get("ok"):
            err = proc.stderr.decode("utf-8", errors="replace").strip()
            result.setdefault("error", err[:300] if err else "未知错误")
        return result

    except subprocess.TimeoutExpired:
        return {"ok": False, "error": "导出超时"}
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}


async def export_ppt_images(ppt_id: str, timeout: int = 120) -> dict:
    """
    导出 PPT 全部页为 PNG（在线程池中跑子进程，不阻塞事件循环）。
    返回状态 dict；失败时 ok=False（前端据此降级到 HTML 渲染）。
    """
    pptx_path = os.path.join(UPLOADS_DIR, ppt_id, "original.pptx")
    if not os.path.exists(pptx_path):
        status = {"ok": False, "error": "PPT 文件不存在"}
        _render_status[ppt_id] = status
        return status

    out_dir = render_dir(ppt_id)
    result = await asyncio.to_thread(_run_export_subprocess, pptx_path, out_dir, timeout)
    _render_status[ppt_id] = result
    _spans_cache.pop(ppt_id, None)  # 清旧缓存
    return result
