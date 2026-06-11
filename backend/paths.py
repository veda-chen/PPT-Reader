"""
路径解析 — 区分"只读资源"与"可写数据"，兼容源码运行与 PyInstaller 打包。

- 源码运行：两者都 = 项目根目录（与历史行为完全一致）。
- 打包成 onefile exe（sys.frozen=True）：
    resource_dir() = sys._MEIPASS（临时解包目录，只读，退出即清）
    data_dir()     = exe 所在目录（可写、持久）
  这样前端/脚本等只读资源随 exe 打包，而 uploads/数据库/.env 落在 exe 旁边并持久保留。
"""
import os
import sys


def _frozen() -> bool:
    return getattr(sys, "frozen", False)


def resource_dir() -> str:
    """只读资源根：frontend、export_worker、vision.js 等随程序分发的文件。"""
    if _frozen():
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def data_dir() -> str:
    """可写数据根：uploads、ppt_reader.db、.env 等运行期产生/需用户编辑的文件。"""
    if _frozen():
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
