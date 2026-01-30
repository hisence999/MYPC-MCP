"""
Excel tools using xlwings - operates on already-open Excel files.
"""
import xlwings as xw
from mcp.server.fastmcp import FastMCP


def register_excel_tools(mcp: FastMCP):
    """Register Excel manipulation tools."""

    @mcp.tool(name="MyPC-execute_excel_code")
    def execute_excel_code(
        code: str,
        file_name: str = None
    ) -> str:
        """
        在已打开的 Excel 中执行 xlwings 代码。

        【参数】
        - code (必需): Python 代码
        - file_name (可选): Excel 文件名，不指定则用活动工作簿

        【可用变量】
        - xw: xlwings 模块
        - wb: 当前工作簿
        - sheet: 当前活动工作表
        - pd: pandas (如已安装)
        - np: numpy (如已安装)

        【常用操作】
        读取: sheet.range('A1').value | sheet.range('A1:C10').value | sheet.used_range.value
        写入: sheet.range('A1').value = 'Hello' | sheet.range('A1:C3').value = [[1,2],[3,4],[5,6]]
        工作表: wb.sheets.add('名') | [s.name for s in wb.sheets] | len(wb.sheets)
        格式: sheet.range('A1').color = (255,0,0) | sheet.range('A1').api.Font.Bold = True

        【注意事项】
        1. 仅操作已打开的 Excel，不会打开新文件
        2. 多个 Excel 打开时，建议指定 file_name
        3. 赋值语句会自动返回被赋值的值
        """
        try:
            # 获取目标工作簿
            target_wb = None

            if file_name:
                # 按文件名查找工作簿
                for book in xw.books:
                    if file_name in book.name:
                        target_wb = book
                        break
                if target_wb is None:
                    available = [b.name for b in xw.books]
                    return f"错误: 找不到名为 '{file_name}' 的工作簿。已打开的工作簿: {available}"
            else:
                # 使用活动工作簿
                if xw.books:
                    target_wb = xw.books.active
                else:
                    return "错误: 没有已打开的 Excel 工作簿。请先打开 Excel 文件。"

            # 获取活动工作表
            sheet = target_wb.sheets.active

            # 准备执行环境 - 预置常用库和内置函数
            try:
                import pandas as pd
                import numpy as np
            except ImportError:
                pd = None
                np = None

            exec_globals = {
                'xw': xw,
                'wb': target_wb,
                'sheet': sheet,
                'pd': pd,          # pandas
                'np': np,          # numpy
                'json': __import__('json'),
                # 开放所有内置函数
                '__builtins__': __builtins__,
            }

            # 执行代码
            result = None

            # 先尝试 eval（表达式，有返回值）
            try:
                result = eval(code, exec_globals)
            except SyntaxError:
                # 如果是语句，使用 exec
                try:
                    exec(code, exec_globals)
                except Exception as e:
                    return f"执行错误: {str(e)}"

                # 智能返回：如果是赋值语句，返回被赋值的变量
                if '=' in code and not code.strip().startswith(('if', 'for', 'while', 'def', 'class', 'try', 'with')):
                    # 提取赋值的变量名
                    var_part = code.split('=')[0].strip()
                    var_name = var_part.split()[0]
                    if var_name in exec_globals and var_name not in ['xw', 'wb', 'sheet', 'pd', 'np', 'json']:
                        result = exec_globals[var_name]

            # 处理返回结果
            if result is None:
                # 尝试返回最后定义的变量
                user_vars = {k: v for k, v in exec_globals.items()
                            if k not in ['xw', 'wb', 'sheet', 'pd', 'np', 'json', '__builtins__']
                            and not k.startswith('_')}
                if user_vars:
                    last_var = list(user_vars.keys())[-1]
                    result = exec_globals[last_var]
                else:
                    return "执行成功"

            if isinstance(result, (list, tuple)):
                import json
                try:
                    return json.dumps(result, ensure_ascii=False, indent=2)
                except:
                    return str(result)
            elif isinstance(result, dict):
                import json
                try:
                    return json.dumps(result, ensure_ascii=False, indent=2)
                except:
                    return str(result)
            else:
                return str(result)

        except Exception as e:
            return f"执行错误: {str(e)}"

    @mcp.tool(name="MyPC-list_excel_books")
    def list_excel_books() -> str:
        """
        列出所有已打开的 Excel 工作簿。

        返回信息包括:
        - 工作簿名称
        - 工作表数量
        - 活动工作表名称
        """
        try:
            if not xw.books:
                return "没有已打开的 Excel 工作簿。"

            import json
            books_info = []

            for book in xw.books:
                info = {
                    "name": book.name,
                    "sheets_count": len(book.sheets),
                    "active_sheet": book.sheets.active.name,
                    "all_sheets": [s.name for s in book.sheets]
                }
                books_info.append(info)

            return json.dumps(books_info, ensure_ascii=False, indent=2)

        except Exception as e:
            return f"错误: {str(e)}"
