"""
键鼠模拟工具

提供基础的键盘输入、鼠标点击、快捷键功能
"""

import pyautogui
from mcp.server.fastmcp import FastMCP
from utils.config import load_config, get_config_value


# 从配置加载安全设置
config = load_config()
failsafe = get_config_value(config, "keyboard_mouse.failsafe", True)
pause = get_config_value(config, "keyboard_mouse.pause", 0.1)
type_interval = get_config_value(config, "keyboard_mouse.type_interval", 0.02)

# 设置安全参数
pyautogui.FAILSAFE = failsafe  # 鼠标移到左上角可紧急中断
pyautogui.PAUSE = pause  # 每个操作后暂停


def register_keyboard_mouse_tools(mcp: FastMCP):
    """注册键鼠模拟工具"""

    @mcp.tool()
    async def type_text(text: str, enter: bool = False) -> str:
        """
        在当前焦点位置输入文字（支持中文）。

        参数:
            text: 要输入的文字内容
            enter: 输入完成后是否按回车键 (默认 False)

        返回:
            操作结果信息
        """
        try:
            import pyperclip

            # 如果全是 ASCII 字符，使用原生输入
            if text.isascii():
                pyautogui.typewrite(text, interval=type_interval)
            else:
                # 含有非 ASCII 字符（如中文），使用剪贴板中转
                pyperclip.copy(text)
                pyautogui.hotkey('ctrl', 'v')

            # 如果需要回车
            if enter:
                pyautogui.press('enter')

            return f"✅ 已输入文字: '{text}'" + (" 并按下回车" if enter else "")

        except Exception as e:
            return f"❌ 输入失败: {str(e)}"

    @mcp.tool()
    async def hotkey(keys: list) -> str:
        """
        按下组合快捷键。

        参数:
            keys: 按键列表，如 ["ctrl", "c"] 表示 Ctrl+C

        常用按键:
            - 修饰键: ctrl, alt, shift, win
            - 功能键: enter, tab, esc, space, backspace, delete
            - 方向键: up, down, left, right
            - 功能键: f1-f12

        示例:
            - 复制: ["ctrl", "c"]
            - 粘贴: ["ctrl", "v"]
            - 全选: ["ctrl", "a"]
            - 保存: ["ctrl", "s"]
            - 撤销: ["ctrl", "z"]
            - 切换窗口: ["alt", "tab"]

        返回:
            操作结果信息
        """
        try:
            pyautogui.hotkey(*keys)

            key_combo = " + ".join(keys)
            return f"✅ 已按下快捷键: {key_combo}"

        except Exception as e:
            return f"❌ 快捷键执行失败: {str(e)}"
