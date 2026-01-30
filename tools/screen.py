"""
屏幕截图和摄像头工具

提供屏幕截图、显示器枚举和摄像头拍照功能
"""

import mss
import base64
import os
import io
import glob
import json
import httpx
from PIL import Image
from datetime import datetime
from mcp.server.fastmcp import FastMCP
from utils.config import load_config, get_config_value


# 全局 VLM 配置
VLM_CONFIG = None


def load_vlm_config():
    """从 config.json 加载 VLM 配置"""
    global VLM_CONFIG
    config = load_config()
    VLM_CONFIG = config.get("vlm", {})


def call_vlm_api(image_path: str) -> str:
    """
    调用 VLM API 分析图像
    兼容 OpenAI 格式（如 GPT-4o, GLM-4V, 本地 LLM）
    """
    if not VLM_CONFIG:
        load_vlm_config()

    # 检查是否启用 VLM
    if not VLM_CONFIG or not VLM_CONFIG.get("enabled"):
        return "AI 分析未在 config.json 中启用。"

    if not VLM_CONFIG.get("api_key"):
        return "错误: 在 config.json 中未配置 VLM API key。"

    try:
        # 优化图像以进行 AI 分析
        img = Image.open(image_path)

        # 如果需要转换为 RGB（例如来自 PNG 的 RGBA）
        if img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')

        # 从配置获取参数
        max_dim = get_config_value(VLM_CONFIG, "vlm_max_dim", 1000)
        quality = get_config_value(VLM_CONFIG, "vlm_quality", 60)

        # 如果尺寸 > max_dim 则调整大小（保持纵横比）
        if max(img.size) > max_dim:
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)

        # 保存为内存 JPEG 格式并压缩
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=quality)
        buffer.seek(0)

        # 编码
        base64_image = base64.b64encode(buffer.read()).decode('utf-8')
        mime_type = "image/jpeg"

        url = f"{VLM_CONFIG.get('base_url', '').rstrip('/')}/chat/completions"
        api_key = VLM_CONFIG.get("api_key")
        model = VLM_CONFIG.get("model", "glm-4.6v")
        prompt = VLM_CONFIG.get("prompt", "图片中有什么？")

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }

        payload = {
            "model": model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": prompt
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:{mime_type};base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            "max_tokens": 1000
        }

        # 从配置获取超时
        timeout = get_config_value(VLM_CONFIG, "vlm_timeout", 60)

        # 使用 httpx 进行同步请求，超时时间更长
        with httpx.Client(timeout=timeout) as client:
            response = client.post(url, headers=headers, json=payload)
            response.raise_for_status()
            result = response.json()

            content = result['choices'][0]['message']['content']
            return content

    except httpx.TimeoutException:
        return "AI 分析错误: 请求超时（图像可能太大或网络慢）。"
    except Exception as e:
        return f"AI 分析错误: {str(e)}"


def cleanup_screenshots(directory: str, max_files: int = 20):
    """只保留最新的 `max_files` 张图片"""
    try:
        files = glob.glob(os.path.join(directory, "*.png")) + glob.glob(os.path.join(directory, "*.jpg"))

        # 按修改时间排序（最旧的在前）
        files.sort(key=os.path.getmtime)

        # 如果文件数超过 max_files，删除最旧的
        if len(files) > max_files:
            files_to_delete = files[:-max_files]
            for f in files_to_delete:
                try:
                    os.remove(f)
                except Exception:
                    pass
    except Exception as e:
        print(f"警告: 清理截图失败: {e}")


def register_screen_tools(mcp: FastMCP, screenshots_dir: str, base_url: str, config: dict = None):
    """注册屏幕工具"""

    # 初始加载配置
    load_vlm_config()

    # 从配置获取参数
    max_screenshots = get_config_value(config, "screen.max_screenshots", 20)

    @mcp.tool(name="MyPC-take_screenshot")
    def take_screenshot(display_index: int = 1, ai_analysis: bool = False) -> str:
        """
        截取指定显示器的屏幕截图

        参数:
            display_index: 要捕获的显示器索引（默认: 1 表示主显示器）
                           使用 1 表示第一个显示器，2 表示第二个，以此类推
            ai_analysis: 如果为 True，使用 AI (VLM) 分析图像内容并提取文本（默认: False）

        返回:
            截图的 URL，以及请求的 AI 分析
        """
        try:
            with mss.mss() as sct:
                monitors = sct.monitors

                if display_index >= len(monitors):
                    return f"错误: 显示器索引 {display_index} 超出范围。可用: 1-{len(monitors)-1}"

                selected_monitor = monitors[display_index]
                info = f"显示器 {display_index}: {selected_monitor['width']}x{selected_monitor['height']} 位于 ({selected_monitor['left']},{selected_monitor['top']})"

                # 捕获截图
                sct_img = sct.grab(selected_monitor)

                # 转换为 PIL Image
                img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")

                # 生成文件名
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"screenshot_{timestamp}.png"
                filepath = os.path.join(screenshots_dir, filename)
                os.makedirs(screenshots_dir, exist_ok=True)

                # 保存新截图前清理旧截图
                cleanup_screenshots(screenshots_dir, max_screenshots)

                # 保存文件（全分辨率 PNG）
                img.save(filepath)

                # 只返回 URL
                url = f"{base_url}/screenshots/{filename}"
                response = f"截图捕获成功！\n\n[信息: {info}]\n\nURL: {url}"

                if ai_analysis:
                    analysis = call_vlm_api(filepath)
                    response += f"\n\n=== AI 分析 ===\n{analysis}"

                return response

        except Exception as e:
            import traceback
            return f"截图错误: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"

    @mcp.tool(name="MyPC-list_monitors")
    def list_monitors() -> str:
        """
        列出可用的显示器及其尺寸

        返回每个显示器的详细信息，包括索引、尺寸和推荐的截图索引
        """
        with mss.mss() as sct:
            monitors = sct.monitors
            info = ["可用显示器:"]

            for i, monitor in enumerate(monitors):
                if i == 0:
                    desc = "所有显示器合并"
                    recommend = "❌ 不推荐"
                else:
                    desc = f"显示器 {i}"
                    # 检查是否是主显示器（通常坐标为 0,0）
                    is_primary = (monitor['left'] == 0 and monitor['top'] == 0)
                    if is_primary:
                        recommend = "✅ 主显示器（推荐）"
                    else:
                        recommend = "✅ 可用"

                info.append(f"\n{i}: {desc}")
                info.append(f"   尺寸: {monitor['width']}x{monitor['height']}")
                info.append(f"   位置: ({monitor['left']}, {monitor['top']})")
                info.append(f"   {recommend}")

            info.append("\n💡 提示: 使用不带参数的 take_screenshot() 自动检测主显示器。")
            return "\n".join(info)

    @mcp.tool(name="MyPC-take_webcam_photo")
    def take_webcam_photo(camera_index: int = 0, ai_analysis: bool = False) -> str:
        """
        使用摄像头拍照

        参数:
            camera_index: 摄像头索引（默认: 0 表示主摄像头）
            ai_analysis: 如果为 True，使用 AI (VLM) 分析图像内容（默认: False）

        返回:
            捕获照片的 URL
        """
        try:
            import cv2
        except ImportError:
            return "错误: 未安装 opencv-python。请安装它以使用摄像头功能。"

        try:
            # 初始化摄像头
            cap = cv2.VideoCapture(camera_index)

            if not cap.isOpened():
                return f"错误: 无法打开索引为 {camera_index} 的摄像头。请检查摄像头是否连接且未被使用。"

            # 读取一帧
            ret, frame = cap.read()

            # 立即释放摄像头
            cap.release()

            if not ret:
                return "错误: 从摄像头捕获帧失败。"

            # 生成文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"webcam_{timestamp}.jpg"
            filepath = os.path.join(screenshots_dir, filename)
            os.makedirs(screenshots_dir, exist_ok=True)

            # 保存新截图前清理旧截图
            cleanup_screenshots(screenshots_dir, max_screenshots)

            # 使用 OpenCV 保存图像
            cv2.imwrite(filepath, frame)

            # 返回 URL
            url = f"{base_url}/screenshots/{filename}"
            response = f"摄像头照片拍摄成功！\n\nURL: {url}"

            if ai_analysis:
                analysis = call_vlm_api(filepath)
                response += f"\n\n=== AI 分析 ===\n{analysis}"

            return response

        except Exception as e:
            return f"拍照错误: {str(e)}"
