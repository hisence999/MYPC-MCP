"""
文件搜索工具 - 使用 Everything

提供基于 Everything 的快速文件搜索功能
"""

import subprocess
import os
from mcp.server.fastmcp import FastMCP
from utils.config import load_config, find_executable, get_config_value


def find_es_executable(config):
    """查找 Everything 命令行接口 (es.exe)"""
    everything_paths = config.get("paths", {}).get("everything", [
        r"C:\Program Files\Everything\es.exe",
        r"C:\Program Files (x86)\Everything\es.exe",
        r"D:\APP\Everything\es.exe",
        "es.exe"
    ])

    return find_executable(everything_paths)


def register_search_tools(mcp: FastMCP, config: dict = None):
    """
    注册使用 Everything (es.exe) 的文件搜索工具

    Args:
        mcp: FastMCP 实例
        config: 配置字典
    """

    # 获取 es.exe 路径
    es_path = find_es_executable(config) if config else None
    default_limit = get_config_value(config, "search.default_limit", 20) if config else 20

    @mcp.tool(name="MyPC-search_files")
    def search_files(query: str, limit: int = None) -> str:
        """
        使用 'Everything' (es.exe) 搜索文件和文件夹

        此工具支持 Everything 的强大搜索语法：
        - 通配符: "*.py", "log*.txt"
        - 扩展名: "ext:png;jpg", "ext:doc"
        - 类型宏: "pic:", "audio:", "video:", "exe:"
        - 逻辑: "foo bar" (AND), "foo | bar" (OR), "!foo" (NOT)
        - 路径: "D:\Downloads\ *.zip"

        参数:
            query: 搜索查询（例如 "config.json", "*.py", "项目笔记"）
            limit: 返回的最大结果数（默认: 20）

        返回:
            匹配文件路径列表
        """
        if limit is None:
            limit = default_limit

        if not es_path:
            return """错误: 未找到 'es.exe' (Everything CLI)。

请安装 Everything 搜索工具：
1. 下载: https://www.voidtools.com/
2. 安装到默认位置或在 config.json 中配置路径

配置示例:
{
    "paths": {
        "everything": [
            "C:\\\\Program Files\\\\Everything\\\\es.exe",
            "D:\\\\APP\\\\Everything\\\\es.exe"
        ]
    }
}"""

        try:
            # 构建命令
            # -n <num>: 限制结果数
            cmd = [es_path, str(query), "-n", str(limit)]

            # 运行搜索
            # creationflags=0x08000000 (CREATE_NO_WINDOW) 防止 cmd 窗口弹出
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='replace',
                creationflags=0x08000000
            )

            if result.returncode != 0:
                return f"搜索错误: {result.stderr}"

            output = result.stdout.strip()

            if not output:
                return f"未找到 '{query}' 的结果"

            # 格式化输出
            lines = output.split('\n')
            count = len(lines)

            response = [f"找到 {count} 个 '{query}' 的结果:"]
            response.extend(lines)

            if count >= limit:
                response.append(f"\n(显示前 {limit} 个结果)")

            return "\n".join(response)

        except Exception as e:
            return f"执行搜索时出错: {str(e)}"
