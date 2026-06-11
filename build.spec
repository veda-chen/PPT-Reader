# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包配置 — 智能PPT阅读器 单文件 exe。
用法：python -m PyInstaller build.spec --noconfirm --clean
产物：dist/智能PPT阅读器.exe
"""
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# ── 隐藏导入（PyInstaller 静态分析易漏的动态依赖）──
hiddenimports = [
    # PowerPoint COM 自动化
    'win32com', 'win32com.client', 'pythoncom', 'pywintypes',
    # 渲染 / 解析 / LLM
    'fitz', 'pptx', 'openai', 'dotenv', 'multipart',
    # 本项目 backend 模块（确保都打进去；export_worker 走 --export-worker 入口）
    'paths', 'db', 'models', 'ppt_parser', 'ppt_to_html',
    'render_service', 'export_worker', 'highlight_service',
    'llm_service', 'vision_service',
]
hiddenimports += collect_submodules('uvicorn')   # uvicorn 的 loops/protocols/lifespan 子模块

# ── 随程序打包的资源 ──
datas = [
    ('frontend', 'frontend'),     # 前端静态资源 + 模板
    ('vision.js', '.'),           # 识图脚本（需用户单装 Node 运行）
    ('.env.example', '.'),        # 配置模板（首次运行复制到 exe 同目录）
]
datas += collect_data_files('pptx')   # python-pptx 自带模板

# ── 排除环境里被误收、本项目用不到的重型库（大幅瘦身、加速打包/启动）──
# 本项目只用 fastapi/uvicorn/openai/pydantic/python-pptx/lxml/Pillow/fitz/win32com/dotenv/sqlite3，
# 下面这些（科学计算 / 绘图 / Jupyter / GUI 工具包）都不在 import 链里，安全排除。
excludes = [
    'matplotlib', 'pandas', 'numpy', 'scipy',
    'IPython', 'ipykernel', 'jupyter', 'jupyter_client', 'jupyter_core',
    'notebook', 'nbformat', 'nbconvert', 'jedi', 'parso',
    'zmq', 'pyzmq', 'tkinter', '_tkinter', 'Tkinter',
    'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'wx',
    'pytest', 'sphinx', 'docutils', 'lark',
    'jsonschema', 'jsonschema_specifications', 'psutil',
]

a = Analysis(
    ['backend/main.py'],
    pathex=['backend'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='智能PPT阅读器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,            # 保留控制台，便于查看翻译/导出日志
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
