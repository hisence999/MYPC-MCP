"""
Office tools using pywin32 - operates on already-open Word & PowerPoint.
"""
import win32com.client as win32
from mcp.server.fastmcp import FastMCP
import json


def register_office_tools(mcp: FastMCP):
    """Register Office manipulation tools (Word & PPT)."""

    # ==========================================
    # 📝 Word 万能执行器
    # ==========================================
    @mcp.tool(name="MyPC-execute_word_code")
    def execute_word_code(code: str) -> str:
        """
        在已打开的 Word 中执行 Python 代码 (基于 pywin32)。

        【可用变量】
        - app: Word 应用程序对象 (Word.Application)
        - doc: 当前活动文档 (ActiveDocument)
        - selection: 当前光标/选区 (Selection)
        - win32: win32com.client 模块

        【常用操作】
        - 读取选中文字: selection.Text
        - 写入文字: selection.TypeText("Hello")
        - 插入段落: doc.Paragraphs.Add()
        - 全文替换: content = doc.Content.Text; doc.Content.Find.Execute(FindText="旧", ReplaceWith="新", Replace=2)
        - 设置格式: selection.Font.Bold = True | selection.Font.Color = RGB(255,0,0)
        - 获取全文: doc.Content.Text
        - 字数统计: doc.Words.Count

        【注意事项】
        1. 仅操作已打开的 Word
        2. 操作实时可见
        """
        try:
            # 连接 Word 实例
            try:
                app = win32.GetActiveObject("Word.Application")
            except Exception:
                return "错误: 未检测到运行中的 Word 进程，请先打开 Word。"

            # 获取上下文
            try:
                doc = app.ActiveDocument
            except:
                return "错误: Word 已打开，但没有活动的文档。"

            selection = app.Selection

            # 准备执行环境
            exec_globals = {
                'win32': win32,
                'app': app,
                'doc': doc,
                'selection': selection,
                'json': json,
                '__builtins__': __builtins__
            }

            # 执行逻辑
            result = None
            try:
                result = eval(code, exec_globals)
            except SyntaxError:
                try:
                    exec(code, exec_globals)
                except Exception as e:
                    return f"执行错误: {str(e)}"

                # 智能返回
                if '=' in code and not code.strip().startswith(('if', 'for', 'while', 'def', 'class', 'try', 'with')):
                    var_part = code.split('=')[0].strip()
                    var_name = var_part.split()[0]
                    if var_name in exec_globals and var_name not in ['app', 'doc', 'selection', 'win32']:
                        result = exec_globals[var_name]

            # 结果格式化
            if result is None:
                user_vars = {k: v for k, v in exec_globals.items()
                            if k not in ['win32', 'app', 'doc', 'selection', 'json', '__builtins__']
                            and not k.startswith('_')}
                if user_vars:
                    last_var = list(user_vars.keys())[-1]
                    result = exec_globals[last_var]
                else:
                    return "执行成功 (无返回值)"

            return str(result)

        except Exception as e:
            return f"系统错误: {str(e)}"

    # ==========================================
    # 📊 PowerPoint 万能执行器
    # ==========================================
    @mcp.tool(name="MyPC-execute_ppt_code")
    def execute_ppt_code(code: str) -> str:
        """
        在已打开的 PowerPoint 中执行 Python 代码 (基于 pywin32)。

        【可用变量】
        - app: PPT 应用程序对象
        - pres: 当前演示文稿 (ActivePresentation)
        - slide: 当前选中的幻灯片 (ActiveSlide)
        - view: 当前视图 (ActiveWindow.View)

        【常用操作】
        - 读取备注: slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        - 新建幻灯片: pres.Slides.Add(Index=pres.Slides.Count+1, Layout=2)
        - 插入标题: slide.Shapes.Title.TextFrame.TextRange.Text = "标题"
        - 插入文本框: slide.Shapes.AddTextbox(1, 100, 100, 200, 50).TextFrame.TextRange.Text = "内容"
        - 获取幻灯片数量: pres.Slides.Count
        - 遍历幻灯片: [s.Name for s in pres.Slides]

        【注意事项】
        1. 仅操作已打开的 PowerPoint
        2. 操作实时可见
        """
        try:
            # 连接 PPT 实例
            try:
                app = win32.GetActiveObject("PowerPoint.Application")
            except Exception:
                return "错误: 未检测到运行中的 PowerPoint 进程。"

            # 获取上下文
            try:
                pres = app.ActivePresentation
            except:
                return "错误: PPT 已打开，但没有活动的演示文稿。"

            # 获取当前幻灯片
            try:
                slide = app.ActiveWindow.View.Slide
            except:
                slide = None

            # 准备执行环境
            exec_globals = {
                'win32': win32,
                'app': app,
                'pres': pres,
                'slide': slide,
                'view': app.ActiveWindow.View,
                'json': json,
                '__builtins__': __builtins__
            }

            # 执行逻辑
            result = None
            try:
                result = eval(code, exec_globals)
            except SyntaxError:
                try:
                    exec(code, exec_globals)
                except Exception as e:
                    return f"执行错误: {str(e)}"

                if '=' in code and not code.strip().startswith(('if', 'for', 'while', 'def', 'class', 'try', 'with')):
                    var_part = code.split('=')[0].strip()
                    var_name = var_part.split()[0]
                    if var_name in exec_globals and var_name not in ['app', 'pres', 'slide', 'view', 'win32']:
                        result = exec_globals[var_name]

            # 结果格式化
            if result is None:
                user_vars = {k: v for k, v in exec_globals.items()
                            if k not in ['win32', 'app', 'pres', 'slide', 'view', 'json', '__builtins__']
                            and not k.startswith('_')}
                if user_vars:
                    last_var = list(user_vars.keys())[-1]
                    result = exec_globals[last_var]
                else:
                    return "执行成功 (无返回值)"

            return str(result)

        except Exception as e:
            return f"系统错误: {str(e)}"
