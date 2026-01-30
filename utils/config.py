"""
MYPC-MCP 配置加载工具

提供配置文件读取和环境变量扩展功能
"""

import os
import json
import re
from typing import Any, Dict, List, Optional


def expand_env_vars(path: str) -> str:
    """
    扩展路径中的环境变量

    支持以下格式:
    - Windows 风格: %USERPROFILE%, %APPDATA%
    - Unix 风格: $HOME, ${HOME}
    - 波浪号: ~ (用户主目录)

    Args:
        path: 可能包含环境变量的路径

    Returns:
        扩展后的绝对路径
    """
    if not path:
        return path

    # Windows 风格: %USERPROFILE%
    path = re.sub(r'%([^%]+)%', lambda m: os.environ.get(m.group(1), m.group(0)), path)

    # Unix 风格: $HOME, ${HOME}
    path = os.path.expandvars(path)

    # ~ 展开
    path = os.path.expanduser(path)

    return path


def expand_env_in_list(paths: List[str]) -> List[str]:
    """
    扩展列表中的所有路径的环境变量

    Args:
        paths: 路径列表

    Returns:
        扩展后的路径列表
    """
    return [expand_env_vars(p) for p in paths] if paths else []


def expand_env_in_config(config: Dict[str, Any]) -> Dict[str, Any]:
    """
    递归扩展配置中的所有环境变量

    Args:
        config: 配置字典

    Returns:
        环境变量扩展后的配置
    """
    if not isinstance(config, dict):
        return config

    result = {}
    for key, value in config.items():
        if isinstance(value, str):
            result[key] = expand_env_vars(value)
        elif isinstance(value, list):
            # 处理列表中的字符串
            result[key] = [
                expand_env_vars(item) if isinstance(item, str) else item
                for item in value
            ]
        elif isinstance(value, dict):
            # 递归处理嵌套字典
            result[key] = expand_env_in_config(value)
        else:
            result[key] = value

    return result


def load_config(config_file: str = "config.json") -> Dict[str, Any]:
    """
    加载配置文件

    Args:
        config_file: 配置文件路径

    Returns:
        配置字典，如果加载失败返回空字典
    """
    config_path = os.path.join(os.path.dirname(__file__), "..", config_file)
    config_path = os.path.abspath(config_path)

    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            # 展开环境变量
            return expand_env_in_config(config)
        except json.JSONDecodeError as e:
            print(f"Warning: Failed to parse {config_file}: {e}")
        except Exception as e:
            print(f"Warning: Failed to load {config_file}: {e}")
    else:
        print(f"Warning: {config_file} not found, using default configuration")

    return {}


def get_config_value(config: Dict[str, Any], key_path: str, default: Any = None) -> Any:
    """
    获取嵌套配置值

    Args:
        config: 配置字典
        key_path: 配置键路径，使用点号分隔，如 "server.port"
        default: 默认值

    Returns:
        配置值，如果不存在则返回默认值
    """
    keys = key_path.split('.')
    value = config

    for key in keys:
        if isinstance(value, dict) and key in value:
            value = value[key]
        else:
            return default

    return value


def get_safe_zones(config: Dict[str, Any]) -> List[str]:
    """
    获取安全区列表，展开环境变量

    Args:
        config: 配置字典

    Returns:
        安全区路径列表
    """
    safe_zones = config.get("safe_zones", [])
    if not safe_zones:
        # 默认安全区
        safe_zones = [
            "~/Documents",
            "~/Downloads",
            "~/Desktop"
        ]

    return expand_env_in_list(safe_zones)


def find_executable(paths: List[str]) -> Optional[str]:
    """
    在路径列表中查找第一个存在的可执行文件

    Args:
        paths: 可执行文件路径列表

    Returns:
        第一个存在的可执行文件路径，如果都不存在则返回 None
    """
    for path in paths:
        expanded = expand_env_vars(path)
        if os.path.exists(expanded):
            return expanded

    return None


def get_drives(config: Dict[str, Any]) -> List[str]:
    """
    获取驱动器列表

    Args:
        config: 配置字典

    Returns:
        驱动器列表，如 ["C:", "D:", "E:"]
    """
    drives = config.get("paths", {}).get("drives", [])

    if not drives:
        # 自动检测驱动器
        drives = []
        for letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
            drive = f"{letter}:"
            if os.path.exists(drive):
                drives.append(drive)

    return drives


def get_workspace(config: Dict[str, Any]) -> str:
    """
    获取工作区目录

    Args:
        config: 配置字典

    Returns:
        工作区路径，默认为用户目录下的 Workspace
    """
    workspace = config.get("paths", {}).get("workspace") or config.get("files", {}).get("default_workspace")

    if workspace:
        return expand_env_vars(workspace)

    # 默认工作区
    return os.path.expanduser("~/Workspace")
