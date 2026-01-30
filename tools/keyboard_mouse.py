"""
键鼠模拟工具 - 精简版
提供基础的键盘输入、鼠标点击、快捷键功能
"""

import pyautogui
import win32gui

# 安全设置
pyautogui.FAILSAFE = True  # 鼠标移到左上角可紧急中断
pyautogui.PAUSE = 0.1  # 每个操作后暂停0.1秒


def register_keyboard_mouse_tools(mcp):
    """注册键鼠模拟工具"""

    @mcp.tool()
    async def type_text(text: str, enter: bool = False) -> str:
        """
        在当前焦点位置输入文字（支持中文）。

        Args:
            text: 要输入的文字内容
            enter: 输入完成后是否按回车键 (默认 False)

        Returns:
            操作结果信息
        """
        try:
            import pyperclip
            
            # 如果全是 ASCII 字符，使用原生输入
            if text.isascii():
                pyautogui.typewrite(text, interval=0.02)
            else:
                # 含有非 ASCII 字符（如中文），使用剪贴板中转
                # 先保存当前剪贴板内容（可选，但这里为了简单直接覆盖）
                pyperclip.copy(text)
                pyautogui.hotkey('ctrl', 'v')
            
            # 如果需要回车
            if enter:
                pyautogui.press('enter')
            
            return f"✅ 已输入文字: '{text}'" + (" 并按下回车" if enter else "")
        
        except Exception as e:
            return f"❌ 输入失败: {str(e)}"

    @mcp.tool()
    async def hotkey(keys: list[str]) -> str:
        """
        按下组合快捷键。

        Args:
            keys: 按键列表，如 ["ctrl", "c"] 表示 Ctrl+C

        Returns:
            操作结果信息
        """
        try:
            pyautogui.hotkey(*keys)
            
            key_combo = " + ".join(keys)
            return f"✅ 已按下快捷键: {key_combo}"

        except Exception as e:
            return f"❌ 快捷键执行失败: {str(e)}"

    @mcp.tool()
    async def get_selected_text(smart_terminal: bool = True, restore_clipboard: bool = True) -> str:
        """
        【获取用户选中的文本】通过剪贴板获取当前选中的文本内容

        用户提到"选中" → get_selected_text（主动复制）
                    ↓
        用户提到"复制"或"剪贴板" → get_clipboard（被动读取）
                    ↓
        用户提到"当前文件" → detect_active_file（文件感知）
                    ↓
        用户提到"光标/输入框" → get_focused_control（UI 感知）
                    ↓

        Args:
            smart_terminal: 是否智能识别终端类型并自动切换快捷键（默认 True）
            restore_clipboard: 操作完成后是否恢复原剪贴板（默认 True）

        Returns:
            选中的文本内容，或错误提示
        """
        try:
            import pyperclip
            import time

            # 1. 保存原剪贴板
            old_clipboard = ""
            if restore_clipboard:
                try:
                    old_clipboard = pyperclip.paste()
                except:
                    pass

            # 2. 获取当前窗口类名
            hwnd = win32gui.GetForegroundWindow()
            class_name = win32gui.GetClassName(hwnd)
            window_title = win32gui.GetWindowText(hwnd)

            # 3. 智能选择复制快捷键
            if smart_terminal:
                if 'CASCADIA' in class_name:
                    # Windows Terminal
                    hotkey = ['ctrl', 'shift', 'c']
                    terminal_type = "Windows Terminal"
                elif 'ConsoleWindowClass' in class_name:
                    # Legacy CMD
                    hotkey = ['enter']
                    terminal_type = "CMD"
                elif 'mintty' in class_name:
                    # Git Bash
                    hotkey = ['ctrl', 'insert']
                    terminal_type = "Git Bash"
                else:
                    # 默认 Ctrl+C
                    hotkey = ['ctrl', 'c']
                    terminal_type = "默认"
            else:
                hotkey = ['ctrl', 'c']
                terminal_type = "默认"

            # 4. 清空剪贴板并执行复制
            pyperclip.copy("")
            time.sleep(0.05)

            if len(hotkey) == 1:
                pyautogui.press(hotkey[0])
            else:
                pyautogui.hotkey(*hotkey)

            time.sleep(0.1)

            # 5. 获取复制的文本
            selected_text = pyperclip.paste()

            # 6. 恢复原剪贴板
            if restore_clipboard:
                pyperclip.copy(old_clipboard)

            # 7. 返回结果
            if selected_text:
                return f"""✅ 成功获取选中文本

【窗口信息】
类名: {class_name}
标题: {window_title}
检测类型: {terminal_type}

【选中内容】
{selected_text}

【统计】
字符数: {len(selected_text)}
行数: {len(selected_text.splitlines())}"""
            else:
                return f"""❌ 未获取到文本内容
"""

        except Exception as e:
            # 尝试恢复剪贴板
            if restore_clipboard:
                try:
                    pyperclip.copy(old_clipboard)
                except:
                    pass
            return f"❌ 获取选中文本失败: {str(e)}"
