# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

智能PPT阅读器 — 本地桌面应用，浏览器访问 `http://127.0.0.1:8800`。Python FastAPI 后端 + 原生 HTML/CSS/JS 前端。核心功能：PPT 高保真预览、高亮笔记、双语对照翻译、文档对话、图片识图。

## 启动/停止

```bash
# 启动（项目根目录）
cd "D:\desk\PPT Reader\PPT Reader\backend"
python main.py

# 停止所有相关进程
powershell -Command "Get-Process python*,POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force"
```

首次启动前需 `pip install -r backend/requirements.txt`（已含 `PyMuPDF`）。普通用户可直接双击根目录 **`启动.bat`**（自动装依赖、缺 `.env` 时从 `.env.example` 生成、起服务、开浏览器）。

dev 模式（`python main.py`，非 frozen）以 `uvicorn.run("main:app", reload=True)` 启动，**改 backend 下 `.py` 自动重启**；frozen（exe）模式走 `uvicorn.run(app, reload=False)`。两种模式内存缓存都会随重启清空。

## 分发与打包

- **源码分发**：发整个项目目录，对方双击 `启动.bat` 即可（需 Python 3.10+）。
- **免安装 exe**：双击根目录 **`打包.bat`** → PyInstaller 按 `build.spec` 打成 `dist/智能PPT阅读器.exe`（onefile，含控制台看日志）。把 `dist/` 整个发给别人,首次运行自动在 exe 同目录生成 `.env` 供填 key。
- **路径解析 `backend/paths.py`**：`resource_dir()`（只读：前端/`vision.js`，frozen 时=`sys._MEIPASS`）与 `data_dir()`（可写：`uploads`/`ppt_reader.db`/`.env`，frozen 时=exe 同目录）。dev 下两者都=项目根，**行为与改造前一致**。新增涉及读写文件的代码必须用这两个函数,不要再写 `dirname(dirname(__file__))`。
- **COM 子进程在 exe 内**：`export_worker.py` 不能作为脚本被 exe 执行。frozen 时 `render_service` 改用 `[sys.executable, "--export-worker", ...]` 调起 exe 自身,由 `main.py` 的 `__main__` 顶部分发到 `export_worker.export()` 后退出（仍是独立进程,隔离 COM）。
- **识图依赖 Node**：`vision.js` 已随仓库/打包分发（`vision_service.py` 用 `resource_dir()/vision.js`,不再是写死的 `~/vision.js`）。exe 不含 Node,缺失时识图返回友好提示,不影响其他功能。

## 环境配置

- `.env` 文件在项目根目录（`main.py` 会同时尝试根目录和 `backend/` 子目录），启动时自动加载
- 必需变量：`QWEN_API_KEY`、`QWEN_BASE_URL`、`QWEN_MODEL`（当前 `.env` 指向阿里云 DashScope 千问模型，OpenAI 兼容接口）
- **兼容**：`llm_service.py` 优先读 `QWEN_*`，找不到时回退旧的 `ZHIPU_*`（历史命名，原指智谱）。`QWEN_MODEL` 缺省回退默认 `qwen-plus`
- 可选变量：`TRANSLATE_CONCURRENCY`（默认 `3`）控制全文翻译并发数。DashScope 等服务商 QPS 较低，调高易触发 429 限流
- 高保真渲染需要本机安装 **Microsoft PowerPoint**（Office 16.0/365），通过 `pywin32` COM 自动化调用。PowerPoint 不可用时自动降级为 python-pptx 文字渲染

## 架构

### 两种渲染模式

| 模式 | 触发条件 | 实现 |
|------|---------|------|
| **高保真**（默认） | PowerPoint COM 可用 | `export_worker.py` 子进程：COM 导出 PDF → PyMuPDF 渲染 PNG + 提取每字精确 bbox (`spans.json`) → `ppt_to_html.py: _render_fidelity_slide()` 构建**PNG 背景 + 透明可选中文字层** |
| **降级** | COM 失败/未装 Office | `ppt_to_html.py: render_slide(fidelity=False)` 用 python-pptx 解析 shape → HTML span |

高保真模式的文字层从 PDF 提取的精确 bbox 定位，与 PNG 同源同坐标系，像素级对齐。每个 span 用 `data-shape-id/data-para-idx/data-run-idx` 映射 PDF 的 block/line/span 索引。

**导出管线**：上传时 `main.py → render_service.export_ppt_images()` → `asyncio.to_thread(subprocess.run)` 调用 `export_worker.py`（独立进程隔离 COM，避免污染 uvicorn）。输出到 `uploads/<ppt_id>/render/`：`slide{N}.png` + `spans.json`。

**文字层双通道提取**（`export_worker.py`）：对每页同时跑 `dict`（`_extract_spans_dict`，结构化 block/line/span，**含真实字号**）和 `words`（平铺单词级，捕获 dict 遗漏的边缘文字），再按 bbox 重叠度去重合并（`_has_overlap`，阈值 0.5），dict 优先。改文字层提取逻辑时两个通道都要顾及。
> 注意：必须用 `get_text("dict")` 而非 `"rawdict"`——rawdict 的 span 没有 `text` 字段（文字在 `chars` 逐字符列表里），早期误用导致该通道恒空、全靠 words 撑且字号失真。

**栅格化文本恢复**（`export_worker._recover_rasterized_spans`）：带特效（发光/阴影/渐变/图片填充等）的文本框被 PowerPoint 导出 PDF 时栅格化成位图，从 PDF 文字层消失（PyMuPDF 提取不到）。该步用 python-pptx 读原文,凡顶层文本框段落文字未出现在已提取文字层者,按其 EMU 坐标（`EMU/12700=pt`，与 spans.json 同坐标系）合成 span 注入,`block` 从 `_SYNTH_BLOCK_BASE=100000` 起编号（带 `synthetic:true` 标记）。这样恢复的文字自动获得可选中/高亮/翻译/双语显示能力,**三处下游消费 spans.json 的代码零改动**。v1 仅处理顶层 shape，GROUP 内暂不恢复（坐标需组变换）。

### LLM 路由

- **文本任务**（翻译、总结、对话、全文 PPT 翻译）：OpenAI 兼容接口 → `llm_service.py`。provider 由 `QWEN_BASE_URL` 决定（当前指向阿里云 DashScope 千问模型；旧 `ZHIPU_*` 变量仍作回退兼容）
- **图片识图**：豆包 VL → `vision_service.py`（subprocess 调用 `resource_dir()/vision.js`，即项目根目录的 `vision.js`，非 OpenAI 协议，仅识图用）

### 双语对照翻译

`bilingual.js` 触发 → `POST /api/ppt/<id>/translate-all` → 后端展平所有段落/行，每段**独立一次 API 调用**确保 1:1 对齐，用 `asyncio.Semaphore(TRANSLATE_CONCURRENCY)` 限流并发翻译（`_translate_single` 自带 429 退避重试，彻底失败时返回原文而非留空）。高保真模式下源文本从 `spans.json` 提取（按 block/line 分组的行），降级模式下从 `python-pptx` 提取（按 shape/paragraph 分组的段落）。

**译后纠错**：`llm_service._PROPER_NOUN_FIX` 是硬编码的专有名词替换词典（如各种误译 → `湛江`），每段译文都会过一遍 `_fix_proper_nouns()`。模型常错译地名/人名，新增纠错项往这里加。

译文页有两条渲染路径：高保真可用时走 `render_translated_fidelity_slide()`（PNG 背景 + 半透明白色覆盖原文 + 可见译文），不可用时走 `_replace_paragraphs_with_translation()` 纯文字替换。翻译结果缓存在内存字典中（`llm_service._translation_cache`），服务器重启后丢失。

### 高亮系统

选中文本 → `highlight.js: captureSelection()` 遍历选区 TreeWalker，对每个命中的 `[data-run-idx]` span 记录 `(shape_id, para_idx, run_idx, char_start, char_end)` → POST `/api/highlights` → SQLite `highlights` 表（`segments_json` 字段存 JSON） → `_renderSingle()` 用 `document.createRange() + surroundContents(<mark>)` 恢复 DOM 高亮。

高保真模式下 span 来自 PDF 的 block/line/span 映射，与普通模式不同但接口一致，`highlight.js` 无需感知模式差异。

**关键注意**：`captureSelection()` 重建文本时不同 span 之间需要插入空格（PDF 提取的 span 常无尾部空格），否则笔记文字粘连。

### 数据库

SQLite 单文件 `ppt_reader.db`，三个表：`presentations`、`highlights`、`chat_messages`。`db.py` 提供 `get_connection()`（WAL 模式 + 外键约束）。

### Uploads 目录结构

每个上传的 PPT 在 `uploads/<ppt_id>/` 下：

```
uploads/<ppt_id>/
├── original.pptx          # 原始上传文件
├── images/                # 提取的图片（上传时 python-pptx 提取）
│   ├── slide0_shape42.jpg
│   └── ...
└── render/                # 高保真导出产物（export_worker.py 输出）
    ├── slides.pdf         # COM 导出的 PDF（中间产物）
    ├── slide0.png         # 高清 PNG（EXPORT_ZOOM=2.0 控制倍率）
    ├── slide1.png
    └── spans.json         # 每字精确 bbox（PyMuPDF 提取，与 PNG 同坐标系）
```

图片提取（`ppt_parser.extract_images()`）与渲染导出（`export_worker.py`）是**两个独立管线**：前者从 python-pptx 提取内嵌图片二进制，后者通过 PowerPoint COM 导出整页像素。

### 翻译缓存（注意）

三处内存缓存均**不持久化**，服务器重启后丢失：
- `llm_service._translation_cache` — 全文翻译结果（按页的段落列表）。重启后需重新触发翻译
- `llm_service._text_cache` — 文档全文（对话/翻译共用，避免重复解析 pptx）
- `render_service._render_status` / `_spans_cache` — 高保真导出状态与文字层 JSON

重启后 `has_rendered()` 返回 `False`，即使磁盘上 `render/` 产物仍在，也会先走降级模式直到重新上传/导出。高保真模式下翻译源文本取自 `spans.json`（保证 100% 文本保真），降级模式下取自 `python-pptx`。

### 前端 JS 模块（无构建工具，IIFE 模式）

| 文件 | 职责 |
|------|------|
| `app.js` | 主控制器：上传、状态管理、工具栏启用 |
| `api.js` | fetch 封装 |
| `slide-renderer.js` | 幻灯片加载/缓存/`zoom` 自适应缩放 |
| `navigation.js` | 缩略图条、键盘翻页 |
| `highlight.js` | 选区捕获、高亮 CRUD、笔记面板渲染 |
| `bilingual.js` | 双语对照模式：触发翻译、同步翻页滚动、`rescaleAll()` |
| `chat-panel.js` | 文档对话面板 |
| `llm-panel.js` | 选中翻译/总结按钮 |
| `image-viewer.js` | 点击图片识图（豆包） |

## 识图能力

底层模型不具备原生识图。遇到图片时用项目根目录的 `vision.js`：

```
node vision.js "<图片路径>" "用中文描述这张图片"
```

服务：火山引擎 Ark，模型 doubao-seed-2-0-lite-260428（`vision.js` 自带默认 Key，可用 `.env` 的 `ARK_API_KEY`/`ARK_BASE_URL`/`VISION_MODEL` 覆盖）。需本机安装 Node.js。

## Windows 特别注意事项

- 终端用 bash（Git Bash），路径用 Unix 风格 `/d/desk/...`
- Python stdout/stderr 在 `main.py` 启动时已强制 UTF-8（`io.TextIOWrapper`）
- 静态资源（JS/CSS）设置了 `Cache-Control: no-cache`，**不需要硬刷新**
- 两条子进程调用路径不同，**不要统一**：
  - `render_service` 高保真导出用 `subprocess.run` + `asyncio.to_thread`（同步隔离 COM，长任务、需超时控制）
  - `vision_service` 识图用 `asyncio.create_subprocess_exec`（依赖 Windows 默认的 ProactorEventLoop；若改回 SelectorEventLoop 会报 `NotImplementedError`）
- PowerPoint COM 导出后需确保 `POWERPNT.EXE` 进程被正确清理（`export_worker.py` 的 finally 块），否则残留僵尸进程会阻塞后续导出
