"""
图片识图服务 — 通过 subprocess 调用 vision.js 使用豆包模型。
仅在识图场景使用豆包，文本任务由 Claude API 处理。
"""
import os
import subprocess
import asyncio
from paths import resource_dir, data_dir

# vision.js 随程序分发（仓库根 / 打包内），不再写死用户家目录
VISION_SCRIPT = os.path.join(resource_dir(), "vision.js")


async def describe_image(ppt_id: str, image_path: str) -> str:
    """
    调用 vision.js 对图片进行识别描述。

    参数:
        ppt_id: PPT 唯一标识
        image_path: 图片在 uploads/<ppt_id>/ 下的相对路径（如 "images/slide0_shape42.jpg"）

    返回:
        图片的中文描述文本
    """
    # 构建图片的完整磁盘路径
    base_upload = os.path.join(data_dir(), "uploads")

    # 智能解析图片路径 — 支持多种格式
    # image_path 可能是：
    #   1. "shape42" (只有 shape_id，来自 data-image-path)
    #   2. "images/slide0_shape42.jpg" (完整相对路径)
    image_full_path = None

    if not image_path.startswith("images/"):
        # 按 shape_id 搜索匹配的图片文件
        images_dir = os.path.join(base_upload, ppt_id, "images")
        if os.path.isdir(images_dir):
            # 搜索包含该 shape_id 的文件
            for fname in os.listdir(images_dir):
                if image_path in fname:
                    image_full_path = os.path.join(images_dir, fname)
                    break
    else:
        # 完整相对路径
        image_full_path = os.path.join(base_upload, ppt_id, image_path)

    if not image_full_path or not os.path.exists(image_full_path):
        return f"未找到图片文件: {image_path}"

    # 调用 vision.js
    try:
        proc = await asyncio.create_subprocess_exec(
            "node", VISION_SCRIPT, image_full_path, "请详细描述这张图片的内容。",
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        stdout, stderr = await asyncio.wait_for(
            proc.communicate(), timeout=30
        )
        if proc.returncode != 0:
            err_msg = stderr.decode("utf-8", errors="replace").strip()
            return f"识图出错: {err_msg}"
        return stdout.decode("utf-8", errors="replace").strip()
    except asyncio.TimeoutError:
        return "识图超时，请重试"
    except FileNotFoundError:
        # 未安装 Node.js（exe 免安装版不内置 Node）
        return "识图功能需安装 Node.js 后可用：https://nodejs.org/"
    except NotImplementedError:
        # 极端情况：事件循环不支持子进程
        return "识图功能当前不可用（子进程不受支持）"
    except Exception as e:
        return f"识图调用失败: {str(e)}"
