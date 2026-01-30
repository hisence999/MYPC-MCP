"""
Active Window File Detector - 智能识别当前窗口关联的文件

支持多种策略：
1. 资源管理器特殊处理（获取路径和选中项）
2. 标题路径提取（记事本等）
3. 软件特征库匹配
4. 进程句柄查询
5. 当前目录搜索
6. Everything 反查（需要 Everything 服务）

Author: MyPC-MCP
"""

import os
import re
import json
from datetime import datetime


def get_active_explorer_info(active_hwnd):
    """
    获取活动资源管理器窗口的路径和选中项

    Args:
        active_hwnd: 活动窗口句柄

    Returns:
        dict: 包含 current_path 和 selected_files，失败返回 None
    """
    try:
        import win32com.client
        import urllib.parse

        shell = win32com.client.Dispatch("Shell.Application")

        # 找到与活动窗口句柄匹配的资源管理器
        for window in shell.Windows():
            if window.HWND == active_hwnd:
                # 获取当前路径
                loc = window.LocationURL
                if loc.startswith("file:///"):
                    current_path = urllib.parse.unquote(loc[8:].replace("/", "\\"))
                else:
                    current_path = window.LocationName

                # 获取选中项
                selected_files = []
                try:
                    items = window.Document.SelectedItems()
                    for i in range(items.Count):
                        selected_files.append(items.Item(i).Path)
                except:
                    pass

                return {
                    "current_path": current_path,
                    "selected_files": selected_files
                }

        return None

    except Exception as e:
        print(f"[ERROR] get_active_explorer_info 失败: {e}")
        return None


def detect_active_file():
    """
    检测当前活动窗口关联的文件路径

    Returns:
        dict: 包含文件信息的字典，或错误信息
    """
    try:
        import win32gui
        import win32process
        import psutil
    except ImportError as e:
        return {"error": f"Missing dependencies: {e}"}

    # 获取活动窗口信息
    try:
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return {"error": "没有检测到活动窗口"}

        window_title = win32gui.GetWindowText(hwnd)

        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)
            process_name = process.name()
        except:
            process_name = "Unknown"

        print(f"[INFO] Detecting... Window: '{window_title}', Process: '{process_name}'")

    except Exception as e:
        return {"error": f"获取窗口信息失败: {e}"}

    candidate_path = None
    strategy_used = None

    # ========== 策略 Explorer: 资源管理器特殊处理 ==========
    if process_name.lower() == "explorer.exe":
        try:
            explorer_info = get_active_explorer_info(hwnd)
            if explorer_info:
                # 如果有选中的文件，返回第一个（主要结果）
                if explorer_info.get("selected_files"):
                    selected_path = explorer_info["selected_files"][0]
                    if os.path.exists(selected_path):
                        strategy_used = "资源管理器选中项"
                        print(f"[OK] 策略 Explorer 成功: {selected_path}")
                        # 如果是文件，返回标准文件信息；如果是目录，返回目录信息
                        if os.path.isfile(selected_path):
                            return build_file_info(selected_path, process_name, strategy_used, window_title, explorer_info)
                        else:
                            return {
                                "type": "directory",
                                "path": selected_path,
                                "filename": os.path.basename(selected_path),
                                "software": process_name,
                                "strategy": strategy_used,
                                "window_title": window_title,
                                "explorer_info": explorer_info
                            }
                # 如果没有选中文件，返回当前路径
                current_path = explorer_info.get("current_path")
                if current_path and os.path.exists(current_path):
                    return {
                        "type": "directory",
                        "path": current_path,
                        "filename": os.path.basename(current_path) if os.path.basename(current_path) else current_path,
                        "software": process_name,
                        "strategy": "资源管理器当前路径",
                        "window_title": window_title,
                        "note": "资源管理器当前路径（未选中文件）",
                        "explorer_info": explorer_info
                    }
        except Exception as e:
            print(f"[WARN] 资源管理器检测失败: {e}")

    # ========== 策略 A: 标题路径提取 ==========
    if not candidate_path:
        # Windows 路径: C:\Users\xxx\file.txt
        path_match = re.search(r'([a-zA-Z]:[\\/][^:*?"<>|\r\n]+)', window_title)
        if path_match:
            path = path_match.group(1)
            if os.path.exists(path):
                candidate_path = path
                strategy_used = "标题路径提取"
                print(f"[OK] 策略 A 成功: {candidate_path}")

    # ========== 策略 A-Plus: 压缩软件文件名提取 ==========
    if not candidate_path:
        # 检测是否是压缩软件（Bandizip, WinRAR, 7-Zip）
        archive_software = ['Bandizip', 'WinRAR', '7-Zip', 'WinZip']
        is_archive = any(sw in window_title for sw in archive_software)

        if is_archive:
            # 从标题提取文件名（压缩软件通常显示 "filename.zip - Software"）
            title_file = window_title.split(' - ')[0].strip()

            # 在下载目录和桌面搜索
            user_home = os.path.expanduser('~')
            search_dirs = [
                os.path.join(user_home, 'Downloads'),
                os.path.join(user_home, 'Desktop'),
                r'D:\DOWNLOAD',
                r'D:\Download',
                user_home
            ]

            for search_dir in search_dirs:
                if not os.path.exists(search_dir):
                    continue

                target_path = os.path.join(search_dir, title_file)
                if os.path.exists(target_path):
                    candidate_path = target_path
                    strategy_used = "压缩软件文件名"
                    print(f"[OK] 策略 A-Plus 成功: {candidate_path}")
                    break

                # 如果直接查找失败，尝试递归搜索（限制深度 3）
                try:
                    for root, dirs, files in os.walk(search_dir):
                        if title_file in files:
                            candidate_path = os.path.join(root, title_file)
                            strategy_used = "压缩软件递归搜索"
                            print(f"[OK] 策略 A-Plus 成功: {candidate_path}")
                            break
                except (PermissionError, OSError):
                    continue

                if candidate_path:
                    break

    # ========== 策略 B: 软件特征库匹配 ==========
    if not candidate_path:
        extracted = extract_path_from_title(window_title, process_name)
        if extracted and os.path.exists(extracted):
            candidate_path = extracted
            strategy_used = "软件特征库"
            print(f"[OK] 策略 B 成功: {candidate_path}")

    # ========== 策略 C: 进程句柄查询 ==========
    if not candidate_path:
        try:
            candidate_path = find_from_process_handles(pid, window_title)
            if candidate_path:
                strategy_used = "进程句柄查询"
                print(f"[OK] 策略 C 成功: {candidate_path}")
        except Exception as e:
            print(f"[WARN] 策略 C 失败: {e}")

    # ========== 策略 D: 当前目录搜索 ==========
    if not candidate_path:
        try:
            candidate_path = search_in_current_dir(window_title)
            if candidate_path:
                strategy_used = "当前目录搜索"
                print(f"[OK] 策略 D 成功: {candidate_path}")
        except Exception as e:
            print(f"[WARN] 策略 D 失败: {e}")

    # ========== 策略 E: Everything 反查（可选） ==========
    if not candidate_path:
        try:
            candidate_path = search_via_everything(window_title)
            if candidate_path:
                strategy_used = "Everything 搜索"
                print(f"[OK] 策略 E 成功: {candidate_path}")
        except Exception as e:
            print(f"[WARN] 策略 E 失败: {e}")

    # ========== 最终处理 ==========
    if candidate_path and os.path.exists(candidate_path):
        return build_file_info(candidate_path, process_name, strategy_used, window_title)
    else:
        return {
            "error": "无法识别文件路径",
            "window_title": window_title,
            "process_name": process_name,
            "suggestion": "建议手动提供文件路径"
        }


def extract_path_from_title(title, process_name):
    """
    根据软件特征从标题中提取路径

    Args:
        title: 窗口标题
        process_name: 进程名称

    Returns:
        str: 提取的绝对路径，或 None
    """
    # 软件特征库
    patterns = [
        # 记事本
        (r"(.*?)(?: - |—)记事本", lambda m: m.group(1)),
        # VS Code
        (r"(.*?)(?: - |—)Visual Studio Code", lambda m: m.group(1)),
        # Word
        (r"(.*?)(?: - |—)Word", lambda m: m.group(1)),
        # Excel
        (r"(.*?)(?: - |—)Excel", lambda m: m.group(1)),
        # PowerPoint
        (r"(.*?)(?: - |—)PowerPoint", lambda m: m.group(1)),
        # Notepad++
        (r"(.*?)(?: - |—)Notepad\+\+", lambda m: m.group(1)),
        # Sublime Text
        (r"(.*?)(?: - |—)Sublime Text", lambda m: m.group(1)),
    ]

    for pattern, extractor in patterns:
        match = re.search(pattern, title)
        if match:
            path = extractor(match)

            # 如果是绝对路径且存在，直接返回
            if os.path.isabs(path) and os.path.exists(path):
                return path

            # 如果是相对路径，尝试转换为绝对路径
            if not os.path.isabs(path):
                # 先尝试在当前工作目录查找
                abs_path = os.path.abspath(path)
                if os.path.exists(abs_path):
                    return abs_path

                # 再尝试在常见用户目录查找
                user_home = os.path.expanduser("~")
                common_dirs = [
                    os.path.join(user_home, "Desktop"),
                    os.path.join(user_home, "Downloads"),
                    os.path.join(user_home, "Documents"),
                ]

                for base_dir in common_dirs:
                    test_path = os.path.join(base_dir, path)
                    if os.path.exists(test_path):
                        return test_path

    return None


def find_from_process_handles(pid, window_title):
    """
    通过进程句柄查找打开的文件

    Args:
        pid: 进程 ID
        window_title: 窗口标题（用于匹配）

    Returns:
        str: 找到的文件路径，或 None
    """
    import psutil

    try:
        process = psutil.Process(pid)
        open_files = process.open_files()

        if not open_files:
            return None

        # 从标题提取文件名（去除软件名后缀）
        # 格式: "@AutomationLog.txt - Notepad" -> "@AutomationLog.txt"
        title_file = window_title.split(' - ')[0].split('—')[0].strip()

        # 优先级：完整匹配 > 扩展名匹配 > 最近修改
        candidates = {
            'exact': [],      # 文件名完全匹配
            'name_match': [], # 文件名部分匹配
            'recent': []      # 最近修改的文档文件
        }

        for f in open_files:
            path = f.path
            basename = os.path.basename(path)
            ext = os.path.splitext(path)[1].lower()

            # 排除系统文件和缓存
            if any(ext in path for ext in ['.dll', '.nls', '.mui', '.fon', '.exe', '.ttf', '.sys', '.drv']):
                continue
            if any(folder in path.lower() for folder in ['windows', 'system32', 'syswow64', 'nvidia', 'cache', 'temp']):
                continue

            # 优先选择文档类型
            if ext in ['.txt', '.py', '.md', '.json', '.docx', '.xlsx', '.pdf']:
                # 完整匹配
                if basename == title_file:
                    candidates['exact'].append(path)
                # 文件名匹配（不含扩展名）
                elif os.path.splitext(basename)[0] == os.path.splitext(title_file)[0]:
                    candidates['exact'].append(path)
                # 部分匹配
                elif title_file.lower() in basename.lower():
                    candidates['name_match'].append(path)
                # 最近文档
                else:
                    candidates['recent'].append(path)

        # 按优先级返回
        if candidates['exact']:
            return candidates['exact'][0]
        if candidates['name_match']:
            return candidates['name_match'][0]
        if candidates['recent']:
            # 按修改时间排序
            candidates['recent'].sort(key=lambda x: os.path.getmtime(x) if os.path.exists(x) else 0, reverse=True)
            return candidates['recent'][0]

        return None

    except (psutil.AccessDenied, psutil.NoSuchProcess):
        return None
    except Exception:
        return None


def search_in_current_dir(window_title):
    """
    在当前工作目录及常见用户目录中搜索文件
    优先匹配标题中的项目名

    Args:
        window_title: 窗口标题

    Returns:
        str: 找到的文件路径，或 None
    """
    import os
    import re

    # 从标题提取文件名和项目名
    # 格式: "● README.md - rubiks-cube - Visual Studio Code"
    title_parts = window_title.split(' - ')
    raw_file = title_parts[0].strip()

    # 清理文件名中的特殊字符（VS Code Git 状态标记等）
    # 移除: ●, ✨, ⚠, 🔥, 📝 等前缀符号和空格
    title_file = raw_file
    # 移除常见的 Git 状态符号
    for prefix in ['●', '✨', '⚠', '🔥', '📝', '◆', '◇', '○', '■', '□']:
        if title_file.startswith(prefix):
            title_file = title_file[len(prefix):].strip()
    # 移除所有特殊 Unicode 字符（保守处理）
    title_file = ''.join(c for c in title_file if ord(c) < 128 or c.isalnum() or c in '._-').strip()

    project_name = None

    # 尝试提取项目名（通常在文件名和软件名之间）
    if len(title_parts) >= 2:
        potential_project = title_parts[1].strip()
        # 排除已知的软件名
        known_software = ['Visual Studio Code', 'VS Code', 'Notepad', 'Word', 'Excel', 'PowerPoint',
                          'Notepad++', 'Sublime Text', 'PyCharm', 'IntelliJ IDEA', '未跟踪的', '未跟踪']
        if potential_project not in known_software:
            project_name = potential_project

    if not title_file or len(title_file) < 2:
        return None

    # 搜索路径列表（按优先级排序）
    search_paths = []

    # 1. 如果有项目名，优先搜索包含项目名的目录
    if project_name:
        current = os.getcwd()
        for _ in range(5):  # 最多搜索5层
            # 检查当前目录是否包含项目名
            if project_name.lower() in current.lower():
                search_paths.insert(0, current)  # 插入到最前面
            # 检查兄弟目录
            parent = os.path.dirname(current)
            if os.path.exists(parent):
                try:
                    for sibling in os.listdir(parent):
                        sibling_path = os.path.join(parent, sibling)
                        if os.path.isdir(sibling_path) and project_name.lower() in sibling.lower():
                            search_paths.append(sibling_path)
                except (PermissionError, OSError):
                    pass
            current = os.path.dirname(current)
            if not current or current == current[:-1]:  # 到达根目录
                break

    # 2. 当前工作目录及父目录
    if os.getcwd() not in search_paths:
        search_paths.append(os.getcwd())
    current = os.getcwd()
    for _ in range(2):  # 只搜索2层父目录
        parent = os.path.dirname(current)
        if parent and parent != current:
            search_paths.append(parent)
            current = parent
        else:
            break

    # 3. 扩展的用户目录（优先级较低）
    user_home = os.path.expanduser("~")
    common_dirs = [
        # 基础目录
        os.path.join(user_home, "Desktop"),
        os.path.join(user_home, "Downloads"),
        os.path.join(user_home, "Documents"),
        os.path.join(user_home, "Pictures"),
        os.path.join(user_home, "Videos"),
        os.path.join(user_home, "Music"),
        # 项目目录
        os.path.join(user_home, "Projects"),
        os.path.join(user_home, "Source"),
        os.path.join(user_home, "Repos"),
        os.path.join(user_home, "GitHub"),
        # 工作目录
        os.path.join(user_home, "Work"),
        os.path.join(user_home, "Workspace"),
    ]
    search_paths.extend(common_dirs)

    # 4. 驱动器根目录的常见子目录
    for drive in [r"C:\\", r"D:\\", r"E:\\", r"F:\\"]:
        if os.path.exists(drive):
            try:
                for item in os.listdir(drive):
                    item_path = os.path.join(drive, item)
                    if os.path.isdir(item_path):
                        # 常见的文件夹名
                        common_folder_names = [
                            'download', 'downloads', 'temp', 'tmp',
                            'work', 'project', 'projects', 'source',
                            'docs', 'documents', 'files', 'data'
                        ]
                        if item.lower() in common_folder_names:
                            search_paths.append(item_path)
            except (PermissionError, OSError):
                continue

    # 去重并保持顺序
    seen = set()
    unique_paths = []
    for path in search_paths:
        if path not in seen:
            seen.add(path)
            unique_paths.append(path)

    # 在每个路径中搜索
    best_match = None

    for search_path in unique_paths:
        if not os.path.exists(search_path):
            continue

        # 直接查找
        full_path = os.path.join(search_path, title_file)
        if os.path.exists(full_path):
            # 如果有项目名，优先返回包含项目名的路径
            if project_name and project_name.lower() in full_path.lower():
                return full_path
            # 保存为候选路径
            if not best_match:
                best_match = full_path

        # 递归搜索（限制深度3）
        try:
            for root, dirs, files in os.walk(search_path):
                # 限制搜索深度
                depth = root[len(search_path):].count(os.sep)
                if depth > 3:
                    continue

                if title_file in files:
                    found_path = os.path.join(root, title_file)
                    # 如果有项目名，优先返回包含项目名的路径
                    if project_name and project_name.lower() in found_path.lower():
                        return found_path
                    # 保存为候选路径
                    if not best_match:
                        best_match = found_path
        except (PermissionError, OSError):
            continue

    return best_match


def search_via_everything(window_title):
    """
    通过 Everything 搜索文件（需要 Everything 服务运行）

    Args:
        window_title: 窗口标题

    Returns:
        str: 找到的文件路径，或 None
    """
    try:
        import subprocess

        # 从标题提取可能的文件名
        filename = os.path.basename(window_title).split('—')[0].split('-')[0].strip()

        if len(filename) < 2:
            return None

        # Everything CLI 的完整路径
        # 优先检查常见安装位置
        possible_es_paths = [
            r"D:\APP\Everything\es.exe",
            r"C:\Program Files\Everything\es.exe",
            r"C:\Program Files (x86)\Everything\es.exe",
            os.path.expanduser(r"~\AppData\Local\Programs\Everything\es.exe"),
        ]

        es_path = None
        for path in possible_es_paths:
            if os.path.exists(path):
                es_path = path
                break

        if not es_path:
            return None

        # 检查 es.exe 是否存在
        if not os.path.exists(es_path):
            return None

        # 使用 es 命令行工具（语法: es.exe filename）
        cmd = f'"{es_path}" {filename}'

        result = subprocess.run(
            cmd,
            shell=True,
            capture_output=True,
            text=True,
            timeout=5,
            encoding='utf-8',
            errors='ignore'
        )

        if result.returncode == 0 and result.stdout.strip():
            # 取第一个结果
            lines = result.stdout.strip().split('\n')
            if lines and lines[0]:
                path = lines[0].strip()
                if os.path.exists(path):
                    return path

        return None

    except FileNotFoundError:
        # es.exe 未找到
        return None
    except subprocess.TimeoutExpired:
        return None
    except Exception:
        return None


def build_file_info(path, process_name, strategy, window_title, explorer_info=None):
    """
    构建文件信息字典

    Args:
        path: 文件路径
        process_name: 进程名称
        strategy: 使用的策略
        window_title: 窗口标题
        explorer_info: 可选的资源管理器信息（当从资源管理器选中时）

    Returns:
        dict: 文件信息
    """
    info = {
        "path": path,
        "filename": os.path.basename(path),
        "software": process_name,
        "strategy": strategy,
        "window_title": window_title,
        "size": os.path.getsize(path),
        "modified": datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S"),
        "is_text": is_text_file(path)
    }

    # 如果有资源管理器信息，添加到结果中
    if explorer_info:
        info["explorer_info"] = explorer_info
        # 如果选中了多个文件，添加提示
        if len(explorer_info.get("selected_files", [])) > 1:
            info["note"] = f"共选中 {len(explorer_info['selected_files'])} 个文件，返回第一个"
            info["all_selected"] = explorer_info["selected_files"]

    # 如果是文本文件，读取预览
    if info["is_text"]:
        try:
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(2000)
                info["preview"] = content
                info["preview_length"] = len(content)
        except Exception as e:
            info["preview"] = f"[预览失败: {e}]"
    else:
        info["preview"] = "[二进制文件，不支持预览]"

    return info


def is_text_file(path):
    """判断是否为文本文件"""
    text_extensions = {
        '.txt', '.py', '.md', '.json', '.js', '.html', '.css', '.log',
        '.ini', '.cfg', '.conf', '.xml', '.yaml', '.yml', '.toml',
        '.csv', '.tsv', '.sql', '.sh', '.bat', '.ps1', '.rb', '.go',
        '.rs', '.c', '.cpp', '.h', '.hpp', '.java', '.kt', '.swift',
        '.php', '.asp', '.aspx', '.jsp', '.jsx', '.tsx', '.vue',
        '.scss', '.sass', '.less', '.styl', '.dockerfile', '.gitignore',
        '.env', '.readme', '.license', '.changelog', '.txt'
    }
    return os.path.splitext(path)[1].lower() in text_extensions


# ========== 测试代码 ==========
if __name__ == "__main__":
    import sys

    # 设置 UTF-8 编码输出
    if sys.platform == "win32":
        import codecs
        sys.stdout = codecs.getwriter("utf-8")(sys.stdout.buffer, "strict")

    print("=" * 50)
    print("Active Window File Detector")
    print("=" * 50)
    print()

    result = detect_active_file()

    print()
    print("=" * 50)
    print("Detection Result:")
    print("=" * 50)
    print(json.dumps(result, ensure_ascii=False, indent=2))

    # 如果成功找到文件，显示详细信息
    if "path" in result:
        print()
        print("=" * 50)
        print("File Preview:")
        print("=" * 50)
        if result.get("is_text") and result.get("preview"):
            print(result["preview"])
        else:
            print(result.get("preview", "No preview"))
