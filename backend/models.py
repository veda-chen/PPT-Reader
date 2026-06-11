"""
Pydantic 请求/响应模型定义。
"""
from pydantic import BaseModel
from typing import Optional, List


# --- PPT ---

class PPTInfo(BaseModel):
    id: str
    original_filename: str
    slide_count: int
    title: str
    created_at: str = ""


class FullTextResponse(BaseModel):
    full_text: str
    slide_texts: List[str]


# --- Highlights ---

class HighlightSegment(BaseModel):
    shape_id: int
    para_idx: int
    run_idx: int
    char_start: int
    char_end: int


class HighlightCreate(BaseModel):
    ppt_id: str
    slide_idx: int
    highlighted_text: str
    segments: List[HighlightSegment]
    color: str = "#FFEB3B"
    note: str = ""


class HighlightUpdate(BaseModel):
    color: Optional[str] = None
    note: Optional[str] = None


class HighlightResponse(BaseModel):
    id: int
    ppt_id: str
    slide_idx: int
    highlighted_text: str
    segments_json: str
    color: str
    note: str
    created_at: str
    updated_at: str


# --- LLM ---

class TranslateRequest(BaseModel):
    text: str
    source_lang: str = "auto"
    target_lang: str = "zh"


class TranslateResponse(BaseModel):
    translated_text: str


class SummarizeRequest(BaseModel):
    text: str
    scope: str = "slide"   # "slide" | "document"
    style: str = "brief"   # "brief" | "detailed"


class SummarizeResponse(BaseModel):
    summary: str


class ChatRequest(BaseModel):
    ppt_id: str
    message: str


class ChatResponse(BaseModel):
    reply: str


class ChatMessageResponse(BaseModel):
    id: int
    role: str
    content: str
    created_at: str


# --- Vision ---

class VisionRequest(BaseModel):
    ppt_id: str
    image_path: str


class VisionResponse(BaseModel):
    description: str


# --- Error ---

class ErrorResponse(BaseModel):
    detail: str
