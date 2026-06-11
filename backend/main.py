"""
FastAPI 主入口 — 智能 PPT 阅读器后端服务。
"""
import os
import sys
import io

# 确保 backend 目录在导入路径中（dev 与打包通用），供 from paths/db/... import
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from paths import data_dir, resource_dir

# 尝试加载 .env（可写数据目录优先：dev=项目根 / 打包=exe 同目录；再回退 backend 目录）
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(data_dir(), ".env"))
    load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))
except ImportError:
    pass

# 确保 Windows 下 stdout 使用 UTF-8 编码
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import json
from contextlib import asynccontextmanager

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware

from db import init_db, get_connection
from models import (
    PPTInfo, FullTextResponse,
    HighlightCreate, HighlightUpdate, HighlightResponse,
    TranslateRequest, TranslateResponse,
    SummarizeRequest, SummarizeResponse,
    ChatRequest, ChatResponse, ChatMessageResponse,
    VisionRequest, VisionResponse,
    ErrorResponse,
)
from ppt_parser import save_uploaded_ppt, parse_ppt, load_presentation
from ppt_to_html import render_slide, render_translated_fidelity_slide
from highlight_service import (
    create_highlight, get_highlights_for_slide,
    update_highlight, delete_highlight,
)
from llm_service import (
    translate_text, summarize_text, chat,
    translate_all_slides, get_translated_paragraphs, get_translation_status,
)
from vision_service import describe_image
from render_service import (
    export_ppt_images, has_rendered, get_render_info, slide_image_path,
)

# ── 路径配置 ──────────────────────────────────────────────
# 可写数据（uploads）走 data_dir()；只读资源（前端）走 resource_dir()
BASE_DIR = data_dir()
UPLOADS_DIR = os.path.join(data_dir(), "uploads")
FRONTEND_DIR = os.path.join(resource_dir(), "frontend")
STATIC_DIR = os.path.join(FRONTEND_DIR, "static")
TEMPLATES_DIR = os.path.join(FRONTEND_DIR, "templates")


@asynccontextmanager
async def lifespan(app: FastAPI):
    """应用启动/关闭时的生命周期管理。"""
    init_db()
    os.makedirs(UPLOADS_DIR, exist_ok=True)
    print(f"[OK] Database initialized: {os.path.join(BASE_DIR, 'ppt_reader.db')}")
    print(f"[OK] Upload dir: {UPLOADS_DIR}")
    # 打包(frozen)模式：服务起来后自动打开浏览器（dev 模式由 启动.bat 负责，避免 reload 双开）
    if getattr(sys, "frozen", False):
        import threading
        import webbrowser
        threading.Timer(1.5, lambda: webbrowser.open("http://127.0.0.1:8800")).start()
    yield


app = FastAPI(
    title="智能PPT阅读器",
    description="支持高亮笔记、一键翻译、智能总结与文档对话",
    version="1.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# 禁用静态资源缓存，避免改动 JS/CSS 后浏览器仍用旧版本（本地单用户应用，无需缓存）
@app.middleware("http")
async def _no_cache_static(request, call_next):
    response = await call_next(request)
    if request.url.path.startswith("/static") or request.url.path == "/":
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response


# ── 静态文件 ──────────────────────────────────────────────
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    """主页面。"""
    html_path = os.path.join(TEMPLATES_DIR, "index.html")
    if not os.path.exists(html_path):
        return HTMLResponse("<h1>index.html 未找到</h1>", status_code=404)
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read())


# ── PPT API ───────────────────────────────────────────────

@app.post("/api/upload", response_model=PPTInfo)
async def api_upload(file: UploadFile = File(...)):
    """上传 .pptx 文件，并触发 PowerPoint 高保真导出。"""
    if not file.filename.endswith(".pptx"):
        raise HTTPException(400, "只支持 .pptx 格式文件")

    content = await file.read()
    ppt_id, file_path = save_uploaded_ppt(content, file.filename)
    info = parse_ppt(file_path, ppt_id)

    conn = get_connection()
    conn.execute(
        "INSERT INTO presentations (id, filename, original_filename, slide_count, title) VALUES (?, ?, ?, ?, ?)",
        (ppt_id, "original.pptx", file.filename, info["slide_count"], info["title"]),
    )
    conn.commit()
    conn.close()

    # 一次性导出全部页为高清 PNG（PowerPoint COM，子进程隔离）
    # 失败时自动降级到 HTML 渲染，不阻断上传
    await export_ppt_images(ppt_id)

    return PPTInfo(
        id=ppt_id,
        original_filename=file.filename,
        slide_count=info["slide_count"],
        title=info["title"],
    )


@app.get("/api/ppt/{ppt_id}", response_model=PPTInfo)
async def api_get_ppt(ppt_id: str):
    """获取 PPT 元信息。"""
    conn = get_connection()
    row = conn.execute("SELECT * FROM presentations WHERE id=?", (ppt_id,)).fetchone()
    conn.close()
    if not row:
        raise HTTPException(404, "PPT 不存在")
    return PPTInfo(
        id=row["id"],
        original_filename=row["original_filename"],
        slide_count=row["slide_count"],
        title=row["title"],
        created_at=row["created_at"],
    )


@app.get("/api/ppt/{ppt_id}/slides/{slide_idx}")
async def api_get_slide(ppt_id: str, slide_idx: int):
    """获取单页幻灯片的 HTML（原文页：优先高保真 PNG+透明文字层）。"""
    pres = load_presentation(ppt_id)
    if pres is None:
        raise HTTPException(404, "PPT 不存在")
    if slide_idx < 0 or slide_idx >= len(pres.slides):
        raise HTTPException(404, f"幻灯片页码无效 (0~{len(pres.slides)-1})")

    # 若已成功导出高清图，用高保真模式；否则降级 HTML 渲染
    use_fidelity = has_rendered(ppt_id)
    html = render_slide(pres, slide_idx, ppt_id, fidelity=use_fidelity)
    return HTMLResponse(html)


@app.get("/api/ppt/{ppt_id}/render/slide{slide_idx}.png")
async def api_get_render_png(ppt_id: str, slide_idx: int):
    """提供 PowerPoint 导出的高清幻灯片 PNG。"""
    path = slide_image_path(ppt_id, slide_idx)
    if not path:
        raise HTTPException(404, "高清图不存在")
    return FileResponse(path, media_type="image/png")


@app.get("/api/ppt/{ppt_id}/render-status")
async def api_render_status(ppt_id: str):
    """查询高保真导出状态。"""
    return get_render_info(ppt_id)


@app.get("/api/ppt/{ppt_id}/text", response_model=FullTextResponse)
async def api_get_full_text(ppt_id: str):
    """获取全文档文本。"""
    pres = load_presentation(ppt_id)
    if pres is None:
        raise HTTPException(404, "PPT 不存在")

    file_path = os.path.join(UPLOADS_DIR, ppt_id, "original.pptx")
    info = parse_ppt(file_path, ppt_id)

    return FullTextResponse(
        full_text="\n\n---\n\n".join(
            f"Slide {i+1}:\n{t}" for i, t in enumerate(info["slide_texts"])
        ),
        slide_texts=info["slide_texts"],
    )


@app.get("/api/ppt/{ppt_id}/images/{filename:path}")
async def api_get_image(ppt_id: str, filename: str):
    """提供提取的图片文件。"""
    filepath = os.path.join(UPLOADS_DIR, ppt_id, "images", filename)
    if not os.path.exists(filepath):
        raise HTTPException(404, "图片不存在")
    return FileResponse(filepath)


# ── PPT 全文翻译 API（双语对照核心）─────────────────────

@app.post("/api/ppt/{ppt_id}/translate-all")
async def api_translate_all(ppt_id: str):
    """翻译整个 PPT 的全部文本。"""
    try:
        result = await translate_all_slides(ppt_id)
        return result
    except ValueError as e:
        raise HTTPException(404, str(e))
    except RuntimeError as e:
        raise HTTPException(500, str(e))


@app.get("/api/ppt/{ppt_id}/slides/{slide_idx}/translated")
async def api_get_translated_slide(ppt_id: str, slide_idx: int):
    """获取翻译后的幻灯片 HTML（结构与原文一致，文本替换为译文）。"""
    pres = load_presentation(ppt_id)
    if pres is None:
        raise HTTPException(404, "PPT 不存在")
    if slide_idx < 0 or slide_idx >= len(pres.slides):
        raise HTTPException(404, f"幻灯片页码无效")

    # 获取译文段落列表（与渲染段落组 1:1 对齐）
    translated_paragraphs = get_translated_paragraphs(ppt_id, slide_idx)
    if not translated_paragraphs:
        await translate_all_slides(ppt_id)
        translated_paragraphs = get_translated_paragraphs(ppt_id, slide_idx)

    # 高保真可用 → 用 spans.json 原生渲染译文（PNG 背景 + 可见译文层）
    # 不可用 → 降级 python-pptx 渲染 + 文本替换
    if has_rendered(ppt_id):
        print(f"[译文] ppt={ppt_id} slide={slide_idx} 高保真路径, 译文行数={len(translated_paragraphs)}")
        html = render_translated_fidelity_slide(ppt_id, slide_idx, translated_paragraphs)
    else:
        print(f"[译文] ppt={ppt_id} slide={slide_idx} 降级路径, 译文段数={len(translated_paragraphs)}")
        html = render_slide(pres, slide_idx, ppt_id)
        html = _replace_paragraphs_with_translation(html, translated_paragraphs)

    return HTMLResponse(html)


def _replace_paragraphs_with_translation(html: str, translated_paragraphs: list[str]) -> str:
    """
    按段落组替换译文，避免中英混杂。

    HTML 中每个 <span> 带 data-shape-id / data-para-idx / data-run-idx。
    同一 (shape_id, para_idx) 的多个 span 属于同一段落（多 run）。
    策略：每遇到一个新段落组，取下一条译文放入该组第一个 span，
    同组其余 span 清空——这样无论原文段落被拆成几个 run 都不会残留原文。
    """
    import re

    if not translated_paragraphs:
        return html

    span_re = re.compile(
        r'(<span\s[^>]*?data-shape-id="([^"]*)"\s+data-para-idx="([^"]*)"\s+data-run-idx="([^"]*)"[^>]*>)(.*?)(</span>)',
        re.DOTALL
    )

    state = {"last_key": None, "idx": -1}

    def repl(m):
        open_tag, shape_id, para_idx, run_idx, text, close = m.groups()
        key = (shape_id, para_idx)

        # 判断是否为空段落占位（零宽空格）
        is_placeholder = text.strip() in ("", "&#8203;", "​")

        if key != state["last_key"]:
            # 新段落组的第一个 span
            state["last_key"] = key
            if is_placeholder:
                # 空段落保持原样，不消耗译文
                return m.group(0)
            state["idx"] += 1
            if state["idx"] < len(translated_paragraphs):
                zh = translated_paragraphs[state["idx"]]
                # HTML 转义
                zh = zh.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                return open_tag + zh + close
            # 译文不足，清空避免残留原文
            return open_tag + close
        else:
            # 同段落的后续 run → 清空（译文已全部放进第一个 span）
            if is_placeholder:
                return m.group(0)
            return open_tag + close

    return span_re.sub(repl, html)


@app.get("/api/ppt/{ppt_id}/translation-status")
async def api_translation_status(ppt_id: str):
    """查询翻译是否已完成。"""
    return get_translation_status(ppt_id)


# ── Highlights API ────────────────────────────────────────

@app.post("/api/highlights", response_model=HighlightResponse)
async def api_create_highlight(h: HighlightCreate):
    """创建高亮。"""
    return create_highlight(h)


@app.get("/api/highlights/{ppt_id}/{slide_idx}")
async def api_get_highlights(ppt_id: str, slide_idx: int):
    """获取某页所有高亮。"""
    return get_highlights_for_slide(ppt_id, slide_idx)


@app.put("/api/highlights/{highlight_id}", response_model=HighlightResponse)
async def api_update_highlight(highlight_id: int, h: HighlightUpdate):
    """更新高亮（颜色或笔记）。"""
    return update_highlight(highlight_id, h)


@app.delete("/api/highlights/{highlight_id}")
async def api_delete_highlight(highlight_id: int):
    """删除高亮。"""
    delete_highlight(highlight_id)
    return {"success": True}


# ── LLM API ───────────────────────────────────────────────

@app.post("/api/llm/translate", response_model=TranslateResponse)
async def api_translate(req: TranslateRequest):
    """翻译文本。"""
    result = await translate_text(req.text, req.source_lang, req.target_lang)
    return TranslateResponse(translated_text=result)


@app.post("/api/llm/summarize", response_model=SummarizeResponse)
async def api_summarize(req: SummarizeRequest):
    """总结文本。"""
    result = await summarize_text(req.text, req.scope, req.style)
    return SummarizeResponse(summary=result)


@app.post("/api/chat", response_model=ChatResponse)
async def api_chat(req: ChatRequest):
    """文档对话。"""
    reply = await chat(req.ppt_id, req.message)
    return ChatResponse(reply=reply)


@app.get("/api/chat/{ppt_id}/history")
async def api_chat_history(ppt_id: str):
    """获取对话历史。"""
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM chat_messages WHERE ppt_id=? ORDER BY created_at ASC",
        (ppt_id,),
    ).fetchall()
    conn.close()
    return [
        ChatMessageResponse(
            id=r["id"], role=r["role"], content=r["content"], created_at=r["created_at"]
        )
        for r in rows
    ]


@app.delete("/api/chat/{ppt_id}/history")
async def api_clear_chat_history(ppt_id: str):
    """清空对话历史。"""
    conn = get_connection()
    conn.execute("DELETE FROM chat_messages WHERE ppt_id=?", (ppt_id,))
    conn.commit()
    conn.close()
    return {"success": True}


# ── Vision API ────────────────────────────────────────────

@app.post("/api/vision/describe", response_model=VisionResponse)
async def api_vision_describe(req: VisionRequest):
    """图片识图（调用豆包模型）。"""
    description = await describe_image(req.ppt_id, req.image_path)
    return VisionResponse(description=description)


# ── 启动入口 ──────────────────────────────────────────────

def _first_run_bootstrap() -> bool:
    """
    打包(frozen)首次运行：在 exe 同目录生成可编辑的 .env，并建好 uploads 目录。
    返回 True 表示本次刚生成了 .env（需用户先填 key），调用方应提示后退出，
    避免带着占位符 key 启动服务（占位符进程不会自动重载用户后填的 key）。
    """
    if not getattr(sys, "frozen", False):
        return False
    os.makedirs(os.path.join(data_dir(), "uploads"), exist_ok=True)
    env_path = os.path.join(data_dir(), ".env")
    if os.path.exists(env_path):
        return False
    example = os.path.join(resource_dir(), ".env.example")
    if not os.path.exists(example):
        return False
    try:
        import shutil
        shutil.copy(example, env_path)
        try:
            os.startfile(env_path)  # Windows：自动用记事本打开供填写
        except Exception:
            pass
        return True
    except Exception as e:
        print(f"[首次运行] 生成 .env 失败：{e}")
        return False


if __name__ == "__main__":
    # 打包(frozen)模式下的高保真导出子进程入口：render_service 以
    # [exe, "--export-worker", <pptx> <out_dir> <zoom>] 调起本程序，独立进程隔离 COM。
    if "--export-worker" in sys.argv:
        _i = sys.argv.index("--export-worker")
        import json as _json
        from export_worker import export as _export
        try:
            _r = _export(sys.argv[_i + 1], sys.argv[_i + 2], float(sys.argv[_i + 3]))
            print(_json.dumps(_r))
        except Exception as _e:
            print(_json.dumps({"ok": False, "error": f"{type(_e).__name__}: {_e}"}))
            sys.exit(1)
        sys.exit(0)

    import uvicorn
    if getattr(sys, "frozen", False):
        if _first_run_bootstrap():
            # 首次运行：刚生成 .env，先让用户填 key，不带占位符启动服务
            print("=" * 56)
            print("  首次运行：已在程序目录生成 .env 配置文件。")
            print("  请在弹出的记事本里填入你的 API Key，保存后，")
            print("  再次双击本程序即可启动。")
            print("=" * 56)
            try:
                input("按回车键退出...")
            except Exception:
                pass
            sys.exit(0)
        # 冻结态不能用 reload / import-string，直接传 app 对象
        uvicorn.run(app, host="127.0.0.1", port=8800)
    else:
        uvicorn.run("main:app", host="127.0.0.1", port=8800, reload=True)
