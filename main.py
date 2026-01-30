"""
MYPC-MCP - Windows PC Control Server

基于 Model Context Protocol (MCP) 的 Windows 电脑控制服务器
为 AI 智能体提供全面的本地 Windows PC 控制能力
"""

import os
import socket
from mcp.server.fastmcp import FastMCP

# 导入配置工具
from utils.config import (
    load_config,
    get_config_value,
    get_safe_zones,
    get_drives,
)

# 导入工具模块
from tools.screen import register_screen_tools
from tools.system import register_system_tools
from tools.files import register_file_tools
from tools.ssh import register_ssh_tools
from tools.window import register_window_tools
from tools.search import register_search_tools

# 新增工具（从原始项目迁移）
from tools.bash import register_bash_tools
from tools.keyboard_mouse import register_keyboard_mouse_tools
from tools.detector import register_detector_tools

# Starlette 组件
from starlette.staticfiles import StaticFiles
from starlette.responses import Response, FileResponse
from urllib.parse import parse_qs


def get_local_ip():
    """获取本地 IPv4 地址"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            hostname = socket.gethostname()
            return socket.gethostbyname(hostname)
        except Exception:
            return "127.0.0.1"


def get_local_ipv6():
    """获取本地 IPv6 地址"""
    try:
        s = socket.socket(socket.AF_INET6, socket.SOCK_DGRAM)
        s.connect(("2001:4860:4860::8888", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "::1"


# ============================================================================
# 配置加载
# ============================================================================
config = load_config()

# 服务器配置
server_enabled = get_config_value(config, "server.enabled", True)
server_name = get_config_value(config, "server.name", "MyPC-MCP")
server_port = get_config_value(config, "server.port", 9999)
server_domain = get_config_value(config, "server.domain", "localhost")
server_host = get_config_value(config, "server.host", "0.0.0.0")
screenshots_dir = get_config_value(config, "server.screenshots_dir", "./screenshots")
local_host_header = get_config_value(config, "server.local_host_header", "localhost:9999")

# 创建截图目录
SCREENSHOTS_DIR = os.path.abspath(screenshots_dir)
os.makedirs(SCREENSHOTS_DIR, exist_ok=True)

# 安全区配置（自动展开环境变量）
SAFE_ZONES = get_safe_zones(config)

# SSH 配置
SSH_CONFIG = config.get("ssh", None)

# 网络配置
LOCAL_IP = get_local_ip()
LOCAL_IPV6 = get_local_ipv6()
BASE_URL = f"http://{server_domain}:{server_port}"

# ============================================================================
# 安全检查
# ============================================================================
def is_safe_path(path: str) -> bool:
    """检查路径是否在安全区内"""
    if not SAFE_ZONES:
        # 默认安全区
        default_zones = [
            os.path.expanduser("~/Documents"),
            os.path.expanduser("~/Downloads"),
            os.path.expanduser("~/Desktop"),
        ]
        zones = [os.path.abspath(z) for z in default_zones]
    else:
        zones = [os.path.abspath(z) for z in SAFE_ZONES]

    try:
        abs_path = os.path.abspath(path)
        norm_path = os.path.normcase(abs_path)

        for zone in zones:
            abs_zone = os.path.abspath(zone)
            norm_zone = os.path.normcase(abs_zone)

            # 检查是否是安全区本身
            if norm_path == norm_zone:
                return True

            # 检查是否是安全区内的文件
            if not norm_zone.endswith(os.sep):
                norm_zone += os.sep

            if norm_path.startswith(norm_zone):
                return True

        return False
    except Exception:
        return False


# ============================================================================
# MCP 服务器初始化
# ============================================================================
print(f"Initializing {server_name}...")
mcp = FastMCP(server_name)

# 注册所有工具
print("Registering tools:")

# 屏幕工具
print("  - Screen tools (screenshot, webcam)")
register_screen_tools(mcp, SCREENSHOTS_DIR, BASE_URL, config)

# 系统工具
print("  - System tools (volume, power, status)")
register_system_tools(mcp)

# 文件工具
print("  - File tools (read, write, search)")
register_file_tools(mcp, safe_zones=SAFE_ZONES, base_url=BASE_URL)

# 窗口工具
print("  - Window tools (process, clipboard, notification)")
register_window_tools(mcp, SCREENSHOTS_DIR, BASE_URL)

# 搜索工具
print("  - Search tools (Everything)")
register_search_tools(mcp, config=config)

# Git Bash 工具
print("  - Git Bash tools")
register_bash_tools(mcp)

# 键鼠工具
print("  - Keyboard/Mouse tools")
register_keyboard_mouse_tools(mcp)

# 文件检测工具
print("  - File detector tools")
register_detector_tools(mcp)

# SSH 工具
if SSH_CONFIG:
    print("  - SSH tools")
    register_ssh_tools(mcp, ssh_config=SSH_CONFIG)

print("All tools registered successfully!")


# ============================================================================
# HTTP 中间件和应用
# ============================================================================
class HostHeaderMiddleware:
    """
    中间件：重写 Host 头为 localhost 以满足 MCP 验证
    允许外部连接同时满足 MCP 内部检查
    """
    def __init__(self, app, local_host):
        self.app = app
        self.local_host = local_host.encode()

    async def __call__(self, scope, receive, send):
        if scope["type"] == "http":
            new_headers = []
            for key, value in scope.get("headers", []):
                if key == b"host":
                    new_headers.append((b"host", self.local_host))
                else:
                    new_headers.append((key, value))
            scope = dict(scope)
            scope["headers"] = new_headers
        await self.app(scope, receive, send)


class CombinedApp:
    """
    组合应用：整合 MCP SSE、静态文件服务和安全下载
    """
    def __init__(self, mcp_app, static_dir):
        self.mcp_app = mcp_app
        self.static_app = StaticFiles(directory=static_dir)

    async def __call__(self, scope, receive, send):
        path = scope.get("path", "")
        query_string = scope.get("query_string", b"").decode("utf-8")

        if path.startswith("/screenshots"):
            # 静态文件服务
            scope = dict(scope)
            scope["path"] = path[len("/screenshots"):]
            await self.static_app(scope, receive, send)

        elif path == "/download":
            # 安全文件下载
            params = parse_qs(query_string)
            file_path_list = params.get("path")

            if not file_path_list:
                response = Response("Error: Missing 'path' parameter.", status_code=400)
                await response(scope, receive, send)
                return

            file_path = file_path_list[0]

            # 安全检查
            if not os.path.exists(file_path):
                response = Response("Error: File not found.", status_code=404)
                await response(scope, receive, send)
                return

            if not is_safe_path(file_path):
                response = Response("Error: Access denied. File is not in a Safe Zone.", status_code=403)
                await response(scope, receive, send)
                return

            # 提供文件
            filename = os.path.basename(file_path)
            response = FileResponse(file_path, filename=filename)
            await response(scope, receive, send)

        else:
            # 转发到 MCP
            await self.mcp_app(scope, receive, send)


# ============================================================================
# 主程序
# ============================================================================
if __name__ == "__main__":
    import uvicorn

    print("\n" + "=" * 60)
    print(f"Starting {server_name} on port {server_port} (SSE mode)...")
    print("=" * 60)
    print(f"Server Name: {server_name}")
    print(f"Domain: {server_domain}")
    print(f"Local IPv4: {LOCAL_IP}")
    print(f"Local IPv6: {LOCAL_IPV6}")
    print(f"Screenshots URL: {BASE_URL}/screenshots/")
    print()

    if SAFE_ZONES:
        print("Safe Zones (from config.json):")
        for zone in SAFE_ZONES:
            print(f"  - {zone}")
    else:
        print("Safe Zones: Using defaults (Documents, Downloads, Desktop)")

    print()
    print("=" * 60)
    print("Accepting connections from all hosts (IPv4 + IPv6)...")
    print("Press Ctrl+C to stop the server")
    print("=" * 60)
    print()

    # 创建组合应用
    mcp_app = mcp.sse_app()
    combined = CombinedApp(mcp_app, SCREENSHOTS_DIR)
    app = HostHeaderMiddleware(combined, local_host_header)

    # 使用 "::" 监听 IPv4 和 IPv6
    uvicorn.run(app, host=server_host, port=server_port)
