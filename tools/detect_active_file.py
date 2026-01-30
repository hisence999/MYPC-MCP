"""
Active Window File Detector - 智能识别当前窗口关联的文件

支持多种策略：
1. 标题路径提取（记事本等）
2. Everything 反查（需要 Everything 服务）
3. 进程句柄查询
4. 当前目录搜索
"""

import os
import re
import json
from datetime import datetime
from utils.config import load_config, expand_env_in_list, find_executable, get_drives


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
        return {"error": f"缺少依赖: {e}"}

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

        print(f"[INFO] 检测中... 窗口: '{window_title}', 进程: '{process_name}'")

    except Exception as e:
        return {"error": f"获取窗口信息失败: {e}"}

    candidate_path = None
    strategy_used = None

    # 加载配置
    config = load_config()
    download_dirs = expand_env_in_list(config.get("paths", {}).get("download_dirs", []))
    drives = get_drives(config)
    common_dirs = config.get("detector", {}).get("common_dirs", [])
    search_depth = config.get("detector", {}).get("search_depth", 3)

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
            # 从标题提取文件名
            title_file = window_title.split(' - ')[0].strip()

            # 在下载目录和桌面搜索
            search_dirs = download_dirs + [
                os.path.join(os.path.expanduser("~"), "Desktop"),
                os.path.expanduser("~")
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

                # 递归搜索（限制深度）
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
        extracted = extract_path_from_title(window_title, process_name, common_dirs)
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
            candidate_path = search_in_current_dir(window_title, drives, common_dirs, search_depth)
            if candidate_path:
                strategy_used = "当前目录搜索"
                print(f"[OK] 策略 D 成功: {candidate_path}")
        except Exception as e:
            print(f"[WARN] 策略 D 失败: {e}")

    # ========== 策略 E: Everything 反查 ==========
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


def extract_path_from_title(title, process_name, common_dirs):
    """
    根据软件特征从标题中提取路径

    Args:
        title: 窗口标题
        process_name: 进程名称
        common_dirs: 常见目录列表

    Returns:
        str: 提取的绝对路径，或 None
    """
    # 软件特征库
    patterns = [
        (r"(.*?)(?: - |—)记事本", lambda m: m.group(1)),
        (r"(.*?)(?: - |—)Visual Studio Code", lambda m: m.group(1)),
        (r"(.*?)(?: - |—)Word", lambda m: m.group(1)),
        (r"(.*?)(?: - |—)Excel", lambda m: m.group(1)),
        (r"(.*?)(?: - |—)PowerPoint", lambda m: m.group(1)),
        (r"(.*?)(?: - |—)Notepad\+\+", lambda m: m.group(1)),
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
                abs_path = os.path.abspath(path)
                if os.path.exists(abs_path):
                    return abs_path

                # 在常见用户目录查找
                user_home = os.path.expanduser("~")
                for dir_name in common_dirs:
                    test_path = os.path.join(user_home, dir_name, path)
                    if os.path.exists(test_path):
                        return test_path

    return None


def find_from_process_handles(pid, window_title):
    """通过进程句柄查找打开的文件"""
    import psutil

    try:
        process = psutil.Process(pid)
        open_files = process.open_files()

        if not open_files:
            return None

        title_file = window_title.split(' - ')[0].split('—')[0].strip()

        candidates = {
            'exact': [],
            'name_match': [],
            'recent': []
        }

        for f in open_files:
            path = f.path
            basename = os.path.basename(path)
            ext = os.path.splitext(path)[1].lower()

            # 排除系统文件
            if any(ext in path for ext in ['.dll', '.nls', '.mui', '.fon', '.exe', '.ttf', '.sys', '.drv']):
                continue
            if any(folder in path.lower() for folder in ['windows', 'system32', 'syswow64', 'nvidia', 'cache', 'temp']):
                continue

            # 优先选择文档类型
            if ext in ['.txt', '.py', '.md', '.json', '.docx', '.xlsx', '.pdf']:
                if basename == title_file:
                    candidates['exact'].append(path)
                elif os.path.splitext(basename)[0] == os.path.splitext(title_file)[0]:
                    candidates['exact'].append(path)
                elif title_file.lower() in basename.lower():
                    candidates['name_match'].append(path)
                else:
                    candidates['recent'].append(path)

        if candidates['exact']:
            return candidates['exact'][0]
        if candidates['name_match']:
            return candidates['name_match'][0]
        if candidates['recent']:
            candidates['recent'].sort(key=lambda x: os.path.getmtime(x) if os.path.exists(x) else 0, reverse=True)
            return candidates['recent'][0]

        return None

    except (psutil.AccessDenied, psutil.NoSuchProcess):
        return None
    except Exception:
        return None


def search_in_current_dir(window_title, drives, common_dirs, search_depth):
    """在当前工作目录及常见用户目录中搜索文件"""
    import os
    import re

    title_parts = window_title.split(' - ')
    raw_file = title_parts[0].strip()

    # 清理文件名中的特殊字符
    title_file = raw_file
    for prefix in ['●', '✨', '⚠', '🔥', '📝', '◆', '◇', '○', '■', '□']:
        if title_file.startswith(prefix):
            title_file = title_file[len(prefix):].strip()
    title_file = ''.join(c for c in title_file if ord(c) < 128 or c.isalnum() or c in '._-').strip()

    project_name = None
    if len(title_parts) >= 2:
        potential_project = title_parts[1].strip()
        known_software = ['Visual Studio Code', 'VS Code', 'Notepad', 'Word', 'Excel', 'PowerPoint',
                          'Notepad++', 'Sublime Text', 'PyCharm', 'IntelliJ IDEA', '未跟踪的', '未跟踪']
        if potential_project not in known_software:
            project_name = potential_project

    if not title_file or len(title_file) < 2:
        return None

    # 搜索路径列表
    search_paths = []

    # 当前工作目录及父目录
    search_paths.append(os.getcwd())
    current = os.getcwd()
    for _ in range(2):
        parent = os.path.dirname(current)
        if parent and parent != current:
            search_paths.append(parent)
            current = parent
        else:
            break

    # 用户目录
    user_home = os.path.expanduser("~")
    for dir_name in common_dirs:
        search_paths.append(os.path.join(user_home, dir_name))

    # 驱动器根目录的常见子目录
    for drive in drives:
        drive_path = drive + "\\"
        if os.path.exists(drive_path):
            try:
                for item in os.listdir(drive_path):
                    item_path = os.path.join(drive_path, item)
                    if os.path.isdir(item_path):
                        if item.lower() in ['download', 'downloads', 'temp', 'tmp', 'work', 'project', 'projects']:
                            search_paths.append(item_path)
            except (PermissionError, OSError):
                continue

    # 去重
    seen = set()
    unique_paths = []
    for path in search_paths:
        if path not in seen:
            seen.add(path)
            unique_paths.append(path)

    # 搜索
    best_match = None
    for search_path in unique_paths:
        if not os.path.exists(search_path):
            continue

        full_path = os.path.join(search_path, title_file)
        if os.path.exists(full_path):
            if project_name and project_name.lower() in full_path.lower():
                return full_path
            if not best_match:
                best_match = full_path

        # 递归搜索
        try:
            for root, dirs, files in os.walk(search_path):
                depth = root[len(search_path):].count(os.sep)
                if depth > search_depth:
                    continue

                if title_file in files:
                    found_path = os.path.join(root, title_file)
                    if project_name and project_name.lower() in found_path.lower():
                        return found_path
                    if not best_match:
                        best_match = found_path
        except (PermissionError, OSError):
            continue

    return best_match


def search_via_everything(window_title):
    """通过 Everything 搜索文件"""
    try:
        import subprocess

        config = load_config()
        everything_paths = config.get("paths", {}).get("everything", [
            r"C:\Program Files\Everything\es.exe",
            r"C:\Program Files (x86)\Everything\es.exe",
            r"D:\APP\Everything\es.exe",
            "es.exe"
        ])

        es_path = find_executable(everything_paths)
        if not es_path:
            return None

        filename = os.path.basename(window_title).split('—')[0].split('-')[0].strip()
        if len(filename) < 2:
            return None

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
            lines = result.stdout.strip().split('\n')
            if lines and lines[0]:
                path = lines[0].strip()
                if os.path.exists(path):
                    return path

        return None

    except FileNotFoundError:
        return None
    except subprocess.TimeoutExpired:
        return None
    except Exception:
        return None


def build_file_info(path, process_name, strategy, window_title):
    """构建文件信息字典"""
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
        '.env', '.readme', '.license', '.changelog'
    }
    return os.path.splitext(path)[1].lower() in text_extensions
