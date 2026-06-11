"""
智谱 GLM API 集成 — 翻译、总结、文档对话、全文PPT翻译。
使用 OpenAI 兼容接口。
"""
import os
import re
import json
from openai import OpenAI
from paths import data_dir

# ── 客户端 ───────────────────────────────────────────────
_client = None

def _get_client():
    """延迟初始化 OpenAI 客户端（指向千问/DashScope 端点）。"""
    global _client
    if _client is None:
        # 优先 QWEN_*，回退旧的 ZHIPU_*（历史命名）以兼容旧 .env
        api_key = (os.environ.get("QWEN_API_KEY") or os.environ.get("ZHIPU_API_KEY", "")).strip()
        base_url = (
            os.environ.get("QWEN_BASE_URL")
            or os.environ.get("ZHIPU_BASE_URL")
            or "https://dashscope.aliyuncs.com/compatible-mode/v1"
        )
        if not api_key or api_key in ("your_api_key_here", "sk-xxxxxxxx"):
            raise RuntimeError("尚未配置 API Key：请在 .env 中把 QWEN_API_KEY 改成你自己的 Key")
        # 含非 ASCII（如未替换的中文占位符）会让 HTTP 鉴权头编码失败，提前给出可读报错
        try:
            api_key.encode("ascii")
        except UnicodeEncodeError:
            raise RuntimeError("API Key 含非英文字符（可能还是占位符）：请在 .env 填入真实的 QWEN_API_KEY（纯英文数字）")
        _client = OpenAI(api_key=api_key, base_url=base_url)
    return _client


MODEL = os.environ.get("QWEN_MODEL") or os.environ.get("ZHIPU_MODEL", "qwen-plus")

# 全文翻译并发上限（DashScope 等服务商有 QPS 限制，过高会触发 429 限流）
TRANSLATE_CONCURRENCY = int(os.environ.get("TRANSLATE_CONCURRENCY", "3"))

# ── 全文档文本缓存 ────────────────────────────────────────
_text_cache: dict[str, str] = {}

# ── PPT 全文翻译缓存 ───────────────────────────────────────
_translation_cache: dict[str, list[str]] = {}


async def _get_full_text(ppt_id: str) -> str:
    """获取全文档文本（带缓存）。"""
    import os as _os
    if ppt_id in _text_cache:
        return _text_cache[ppt_id]

    from ppt_parser import parse_ppt
    uploads = _os.path.join(data_dir(), "uploads")
    file_path = _os.path.join(uploads, ppt_id, "original.pptx")
    if _os.path.exists(file_path):
        info = parse_ppt(file_path, ppt_id)
        full = "\n\n---\n\n".join(
            f"第{i+1}页:\n{t}" for i, t in enumerate(info["slide_texts"])
        )
        _text_cache[ppt_id] = full
        return full
    return ""


def _call_llm(messages: list[dict], max_tokens: int = 1024, temperature: float = 0.3) -> str:
    """通用 LLM 调用封装。"""
    client = _get_client()
    resp = client.chat.completions.create(
        model=MODEL,
        max_tokens=max_tokens,
        temperature=temperature,
        messages=messages,
    )
    return resp.choices[0].message.content.strip()


# ── 翻译 ─────────────────────────────────────────────────

async def translate_text(text: str, source_lang: str = "auto", target_lang: str = "zh") -> str:
    """翻译文本。"""
    lang_map = {"zh": "中文", "en": "English", "ja": "日本語"}
    target_name = lang_map.get(target_lang, target_lang)
    return _call_llm([{
        "role": "user",
        "content": f"请将以下文本翻译为{target_name}。只输出译文，不要解释：\n\n{text}"
    }], temperature=0)


# ── 总结 ─────────────────────────────────────────────────

async def summarize_text(text: str, scope: str = "slide", style: str = "brief") -> str:
    """总结文本。无论原文是什么语言，都用简体中文输出。"""
    scope_name = "当前幻灯片" if scope == "slide" else "整份文档"
    style_prompt = "简明扼要地" if style == "brief" else "详细地"
    return _call_llm([
        {
            "role": "system",
            "content": "你是一个文档总结助手。无论原文是中文、英文还是其他语言，你的总结必须始终使用简体中文输出。"
        },
        {
            "role": "user",
            "content": f"请用简体中文{style_prompt}总结以下{scope_name}的内容，提取关键要点（务必用中文回答）：\n\n{text}"
        }
    ])


# ── 文档对话 ─────────────────────────────────────────────

async def chat(ppt_id: str, message: str) -> str:
    """文档对话 — 以全文档文本为上下文。"""
    from db import get_connection

    full_text = await _get_full_text(ppt_id)

    # 获取历史消息
    conn = get_connection()
    rows = conn.execute(
        "SELECT role, content FROM chat_messages WHERE ppt_id=? ORDER BY created_at ASC",
        (ppt_id,),
    ).fetchall()
    conn.close()

    # 构建消息列表
    messages = []
    if full_text:
        messages.append({
            "role": "system",
            "content": f"你是一个PPT文档问答助手。以下是这份PPT的完整内容，请基于它回答用户问题。如果问题与文档无关，请诚实告知。回答使用中文。\n\n=== PPT内容 ===\n{full_text}"
        })

    for r in rows[-20:]:  # 最多保留最近20条历史
        messages.append({"role": r["role"], "content": r["content"]})

    messages.append({"role": "user", "content": message})

    # 保存用户消息
    conn = get_connection()
    conn.execute(
        "INSERT INTO chat_messages (ppt_id, role, content) VALUES (?, 'user', ?)",
        (ppt_id, message),
    )
    conn.commit()
    conn.close()

    try:
        reply = _call_llm(messages, max_tokens=2048)
    except Exception as e:
        reply = f"对话出错: {str(e)}"

    # 保存 AI 回复
    conn = get_connection()
    conn.execute(
        "INSERT INTO chat_messages (ppt_id, role, content) VALUES (?, 'assistant', ?)",
        (ppt_id, reply),
    )
    conn.commit()
    conn.close()

    return reply


# ── 全文 PPT 翻译（双语对照核心功能）────────────────────
# 缓存结构：{ ppt_id: [ [para0_zh, para1_zh, ...],  # slide 0
#                       [para0_zh, ...],            # slide 1
#                       ... ] }
# 每张幻灯片是一个译文段落列表，与 ppt_to_html.extract_render_paragraphs 严格 1:1 对齐。


# ── 专有名词纠错词典 ─────────────────────────────────────
# 模型翻译专有名词（地名、人名、机构名）时容易出错，译后强制替换。
# key = 错误译文中的词, value = 正确译文
# 注意：仅替换匹配到的子串，不只针对整段，确保部分匹配也能纠正。
_PROPER_NOUN_FIX: dict[str, str] = {
    # Zhanjiang → 湛江（模型常错译为汕尾/汕头等广东其他城市，或保留拼音）
    "汕汕": "湛江",
    "汕尾": "湛江",
    "山江": "湛江",
    "展江": "湛江",
    "Zhanjiang": "湛江",
    "zhanjiang": "湛江",
}


def _fix_proper_nouns(text: str) -> str:
    """译后纠错：用词典替换模型翻译错误的专有名词。"""
    for wrong, correct in _PROPER_NOUN_FIX.items():
        if wrong in text:
            print(f"[纠错] 替换 \"{wrong}\" → \"{correct}\"")
            text = text.replace(wrong, correct)
    return text


def _translate_single(text: str, max_retries: int = 3) -> str:
    """翻译单段文本，失败时自动重试（限流额外退避），彻底失败返回原文。"""
    import time
    last_error = None
    for attempt in range(max_retries + 1):
        try:
            resp = _call_llm([
                {"role": "system", "content": "你是专业的PPT翻译助手。将用户输入的文本翻译为简体中文。只输出译文，不要解释。专有名词（地名、人名、机构名）务必翻译为正确的中文，不确定时保留原文。"},
                {"role": "user", "content": f"请翻译为简体中文（只输出译文）：\n{text}"},
            ], max_tokens=1024, temperature=0)
            fixed = _fix_proper_nouns(resp)
            if resp != fixed:
                print(f"[翻译] 原文=\"{text[:60]}\" → 模型=\"{resp[:60]}\" → 纠错=\"{fixed[:60]}\"")
            return fixed
        except Exception as e:
            last_error = e
            msg = str(e)
            # 配置/鉴权类永久错误：重试无意义，直接放弃（避免逐行刷屏）
            permanent = isinstance(e, (RuntimeError, UnicodeEncodeError)) or any(
                k in msg for k in ("ascii", "401", "invalid_api_key", "Unauthorized", "API Key")
            )
            if permanent:
                print(f"[翻译] 配置/鉴权错误，停止重试: {msg[:120]}")
                break
            if attempt < max_retries:
                # 限流（429）退避更久，其它错误用较短退避
                is_rate_limit = "429" in str(e) or "limit_requests" in str(e)
                wait = (3.0 if is_rate_limit else 1.0) * (attempt + 1)
                tag = "限流" if is_rate_limit else "错误"
                print(f"[翻译] {tag}重试 {attempt+1}/{max_retries}: \"{text[:40]}\" 等待 {wait:.1f}s...")
                time.sleep(wait)
            else:
                print(f"[翻译] 彻底失败: \"{text[:40]}\" → {last_error}")
    # 所有重试耗尽，返回原文（保留原文好过留空）
    return text


def _translate_paragraphs(paragraphs: list[str]) -> list[str]:
    """
    逐段翻译，保证返回列表与输入严格 1:1 对齐。
    空段落（仅空白）原样保留。
    """
    result = list(paragraphs)
    for i, text in enumerate(paragraphs):
        if text.strip():
            result[i] = _translate_single(text)
    return result


def _extract_fidelity_slides_lines(ppt_id: str) -> list[list[str]]:
    """
    从高保真 spans.json 提取所有幻灯片的文本行列表。
    每行 = 同一 (block, line) 下所有 span 的文本拼接。
    返回 [[slide0_lines], [slide1_lines], ...]，与 render_translated_fidelity_slide 的分组对齐。
    """
    import os as _os
    import json as _json

    base = data_dir()
    path = _os.path.join(base, "uploads", ppt_id, "render", "spans.json")
    if not _os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8") as f:
        data = _json.load(f)

    result = []
    for slide_spans in data.get("slides", []):
        groups: dict[tuple, list[dict]] = {}
        for sp in slide_spans:
            key = (sp["block"], sp["line"])
            groups.setdefault(key, []).append(sp)

        lines = []
        for key in sorted(groups.keys()):
            line_spans = sorted(groups[key], key=lambda s: s["span"])
            text = "".join(s["text"] for s in line_spans)
            lines.append(text)
        result.append(lines)
    return result


async def translate_all_slides(ppt_id: str) -> dict:
    """
    翻译整个 PPT 的全部幻灯片（按段落/行结构化翻译）。
    高保真可用时优先从 spans.json 提取文本（PowerPoint 导出，100% 保真），
    否则回退 python-pptx 提取。
    结果缓存在内存中，重复调用直接返回缓存。
    """
    import asyncio
    global _translation_cache

    if ppt_id in _translation_cache:
        return {"slide_count": len(_translation_cache[ppt_id]), "translating": False}

    # 提前校验 API Key：无效时快速失败（避免对上百行逐行重试刷屏）
    _get_client()

    import os as _os
    uploads = _os.path.join(data_dir(), "uploads")
    file_path = _os.path.join(uploads, ppt_id, "original.pptx")
    if not _os.path.exists(file_path):
        raise ValueError("PPT 文件不存在")

    from render_service import has_rendered

    if has_rendered(ppt_id):
        # 高保真模式：从 spans.json 提取（PowerPoint 导出 PDF → PyMuPDF 提取 bbox）
        slides_paragraphs = _extract_fidelity_slides_lines(ppt_id)
        total_lines = sum(len(paras) for paras in slides_paragraphs)
        print(f"[翻译] ppt={ppt_id} 高保真路径, {len(slides_paragraphs)} 页, {total_lines} 行待翻译")
    else:
        # 降级模式：python-pptx 提取
        from pptx import Presentation
        from ppt_to_html import extract_render_paragraphs
        pres = Presentation(file_path)
        slides_paragraphs = [extract_render_paragraphs(slide) for slide in pres.slides]
        total_paras = sum(len(paras) for paras in slides_paragraphs)
        print(f"[翻译] ppt={ppt_id} 降级路径, {len(slides_paragraphs)} 页, {total_paras} 段待翻译")

    # 展平所有段落，逐段翻译（每段独立 API 调用，质量对齐网页版）
    all_paras: list[tuple[int, int, str]] = []  # (slide_idx, para_idx, text)
    for si, paras in enumerate(slides_paragraphs):
        for pi, text in enumerate(paras):
            if text.strip():
                all_paras.append((si, pi, text))

    # 限制并发数，避免触发服务商限流（DashScope 默认 QPS 较低，全量并发会 429）
    sem = asyncio.Semaphore(TRANSLATE_CONCURRENCY)

    async def _translate_one(si: int, pi: int, text: str) -> tuple[int, int, str]:
        async with sem:
            return (si, pi, await asyncio.to_thread(_translate_single, text))

    tasks = [_translate_one(si, pi, text) for si, pi, text in all_paras]
    results = await asyncio.gather(*tasks)

    # 组装回 slides_paragraphs 结构
    translations = [list(paras) for paras in slides_paragraphs]  # 深拷贝
    for si, pi, translated in results:
        translations[si][pi] = translated

    _translation_cache[ppt_id] = list(translations)
    return {"slide_count": len(translations), "translating": False}


def get_translated_paragraphs(ppt_id: str, slide_idx: int) -> list[str]:
    """获取单张幻灯片的译文段落列表（与渲染段落组 1:1 对齐）。"""
    if ppt_id not in _translation_cache:
        return []
    cache = _translation_cache[ppt_id]
    if slide_idx < 0 or slide_idx >= len(cache):
        return []
    return cache[slide_idx]


def get_translation_status(ppt_id: str) -> dict:
    """查询翻译状态。"""
    if ppt_id in _translation_cache:
        return {"translated": True, "slide_count": len(_translation_cache[ppt_id])}
    return {"translated": False, "slide_count": 0}
