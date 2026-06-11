#  智能PPT阅读器

> 一个本地运行的智能 PPT 阅读工具：**像素级高保真预览** + **划词高亮笔记** + **一键双语对照翻译** + **文档对话** + **图片识图**。

![platform](https://img.shields.io/badge/platform-Windows-blue)
![python](https://img.shields.io/badge/Python-3.10%2B-3776AB)
![backend](https://img.shields.io/badge/backend-FastAPI-009688)
![local](https://img.shields.io/badge/run-100%25%20本地-success)

打开浏览器访问 `http://127.0.0.1:8800`，上传一个 `.pptx`，即可在网页里高保真地阅读、批注和翻译你的演示文稿。数据全部留在本机，不上传云端。

---

##  它能做什么

| 功能 | 说明 |
|------|------|
| **高保真预览** | 调用本机 PowerPoint 引擎渲染，字体、排版、图表、SmartArt 与原 PPT 像素级一致 |
| **高亮笔记** | 选中文字即可高亮、加批注，刷新/翻页后自动恢复 |
| **双语对照** | 一键全文翻译，左右分屏同步翻页，原文译文对照阅读 |
| **文档对话** | 基于整份 PPT 内容的智能问答 |
| **智能总结** | 当前页或全文一键总结 |
| **划词翻译** | 选中任意文字即时翻译 |
| **图片识图** | 点击幻灯片里的图片，AI 自动描述内容 |

> 即便是带特效、被「压成图片」的文字，也能被还原出来，照样可选中、可翻译。

---

##  快速开始

### 方式一：免安装版（推荐给普通用户）

1. 拿到打包好的 `dist` 文件夹，双击 **`智能PPT阅读器.exe`**；
2. 首次运行会自动生成 `.env` 并打开记事本，把 `QWEN_API_KEY` 改成你自己的 Key 后保存；
3. 再次双击 exe，浏览器会自动打开 `http://127.0.0.1:8800`，上传 `.pptx` 即可使用。

> 无需安装 Python。上传记录、数据库都保存在 exe 同目录，可随程序一起拷走。

### 方式二：源码运行

```bash
# 1. 安装依赖
pip install -r backend/requirements.txt

# 2. 配置 Key：复制 .env.example 为 .env，填入你的 API Key

# 3. 启动（Windows 可直接双击「启动.bat」一键完成上述步骤）
cd backend && python main.py
```

启动后浏览器打开 **http://127.0.0.1:8800**。

### 配置 API Key

在 `.env` 中填写（兼容任意 OpenAI 协议的服务商，默认指向阿里云 DashScope 千问）：

```ini
QWEN_API_KEY=你的_API_Key
QWEN_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1
QWEN_MODEL=qwen-max
```

> Key 获取：[DashScope 控制台](https://dashscope.console.aliyun.com)

### 环境要求

- **Windows 10/11**
- **高保真预览**：需安装 **Microsoft PowerPoint**（Office 2016/365）。没有也能用，会自动降级为文字渲染。
- **图片识图**（可选）：需安装 **[Node.js](https://nodejs.org/)**；未安装时该功能会提示不可用，不影响其他功能。

---

## 📖 使用指南

| 操作 | 方式 |
|------|------|
| 上传 PPT | 点右上角「上传PPT」或拖拽 `.pptx` |
| 翻页 | 左侧缩略图 / 键盘 ← → / 底部按钮 |
| 高亮 | 选中文字 → 选颜色 |
| 加笔记 | 点已有高亮 → 「💬 添加笔记」 |
| 双语对照 | 顶部「📖 双语对照」 |
| 划词翻译 | 选中文字 → 「🌐 翻译」 |
| 总结 | 「📝 总结本页」/「📄 总结全文」 |
| 文档对话 | 右侧「对话」面板提问 |
| 图片识图 | 点击幻灯片里的图片 |

---

## 🛠️ 技术实现

纯 **Python FastAPI 后端 + 原生 HTML/CSS/JS 前端**，无前端构建工具，本地单文件运行。

| 层面 | 技术 |
|------|------|
| 后端 | Python · FastAPI · Uvicorn |
| 高保真渲染 | PowerPoint COM（pywin32）+ PyMuPDF |
| PPT 解析 | python-pptx |
| 大语言模型 | 千问 Qwen / 任意 OpenAI 兼容接口 |
| 图片识图 | 豆包 VL（火山引擎 Ark） |
| 数据存储 | SQLite |
| 打包 | PyInstaller（单文件 exe） |

**核心思路——怎么做到「和原 PPT 一模一样」**：不靠解析重排，而是**让 PowerPoint 自己渲染自己**——

```
PowerPoint 导出 PDF ──► PyMuPDF 渲染成高清 PNG（作背景）
                    └─► PyMuPDF 提取每个字的精确坐标（作透明文字层）
背景图 + 透明文字层叠加 = 像素级保真 ＋ 文字可选中/可翻译
```

两者坐标同源于同一个 PDF，天然对齐；PowerPoint 不可用时自动降级为 python-pptx 文字渲染。

---

## 📂 目录结构

```
智能PPT阅读器/
├── backend/                # FastAPI 后端
│   ├── main.py             # 入口：路由、FastAPI 应用、打包后的 exe 入口
│   ├── ppt_parser.py       # python-pptx 解析、提取图片
│   ├── ppt_to_html.py      # 幻灯片渲染成 HTML（高保真 / 降级两种模式）
│   ├── export_worker.py    # 独立子进程：PowerPoint COM 导出 PDF/PNG + 提取文字坐标
│   ├── render_service.py   # 调度高保真导出、缓存渲染状态
│   ├── llm_service.py      # 翻译 / 总结 / 对话（OpenAI 兼容接口）
│   ├── vision_service.py   # 识图：subprocess 调用 vision.js
│   ├── highlight_service.py# 高亮笔记的增删查
│   ├── db.py / models.py   # SQLite 连接、数据模型
│   ├── paths.py            # 源码运行 / 打包路径解析
│   └── requirements.txt    # Python 依赖
├── frontend/               # 原生 HTML/CSS/JS 前端（无构建工具）
│   ├── templates/index.html
│   └── static/css, static/js   # 上传、翻页、高亮、双语、对话、识图等模块
├── vision.js               # 识图脚本（Node，调用豆包 VL）
├── .env.example            # 配置模板（复制为 .env 后填 Key）
├── 启动.bat                # 源码版一键启动（装依赖 + 起服务 + 开浏览器）
├── 打包.bat + build.spec   # 用 PyInstaller 打包成免安装 exe
├── 使用说明.txt            # 给最终用户的简易说明
└── README.md / CLAUDE.md   # 项目说明 / 开发者指南
```

> 运行后还会生成 `uploads/`（上传的 PPT 与渲染产物）和 `ppt_reader.db`（高亮、对话记录）——这些是你的本地数据，已在 `.gitignore` 中排除，不会上传。
