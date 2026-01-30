"""
Git Bash 命令执行工具

提供在 Windows 上执行 Git Bash 命令的功能，带有黑名单安全机制。
"""

import subprocess
import os
import re
from mcp.server.fastmcp import FastMCP
from utils.config import load_config, find_executable


def find_git_bash() -> str:
    """从配置或常见位置查找 Git Bash 可执行文件"""
    config = load_config()
    git_bash_paths = config.get("paths", {}).get("git_bash", [
        r"C:\Program Files\Git\bin\bash.exe",
        r"C:\Program Files (x86)\Git\bin\bash.exe",
        r"D:\Program Files\Git\bin\bash.exe",
        r"E:\Program Files\Git\bin\bash.exe",
    ])

    return find_executable(git_bash_paths)


# =============================================================================
# 黑名单系统
# =============================================================================

# 单词命令黑名单（仅完整单词匹配）
BLOCKED_COMMANDS = {
    # 删除操作（使用 delete_file 工具代替）
    "rm", "rmdir", "shred", "del", "erase",

    # 写入操作（使用 write_file 工具代替）
    "tee",

    # 编辑器（使用 edit_file 工具代替）
    "vi", "vim", "nvim", "nano", "emacs", "ed", "notepad",

    # 移动/重命名（使用 move_file 工具代替）
    "mv", "rename", "ren",

    # 复制（使用 copy_file 工具代替）
    "cp", "xcopy", "robocopy",

    # 系统控制（太危险）
    "reboot", "shutdown", "halt", "poweroff",

    # 磁盘操作（太危险）
    "mkfs", "fdisk", "format", "parted", "gdisk",

    # 权限更改（可能破坏文件访问）
    "chmod", "chown", "chgrp", "attrib",

    # 用户管理（安全风险）
    "useradd", "userdel", "usermod", "passwd", "adduser", "deluser",

    # 强制杀死（可能杀死重要进程）
    "killall", "taskkill",

    # 网络防火墙（可能阻止访问）
    "iptables",
}

# 多词模式黑名单（精确子串匹配）
BLOCKED_PATTERNS = [
    # 输出重定向
    (r'\s+>\s+', "输出重定向被阻止。请使用 write_file 工具"),
    (r'\s+>>\s+', "追加重定向被阻止。请使用 write_file 工具"),
    (r'\s+2>\s+', "标准错误重定向被阻止。请使用 write_file 工具"),

    # 危险组合命令
    (r'\bkill\s+-9\b', "强制杀死 (-9) 被阻止"),
    (r'\bpkill\s+-9\b', "强制 pkill (-9) 被阻止"),
    (r'\binit\s+[06]\b', "init 0/6 被阻止"),
    (r'\bdd\s+if=', "dd 写操作被阻止"),
    (r'\bsystemctl\s+(poweroff|reboot|halt)\b', "systemctl 电源命令被阻止"),
    (r'\bshutdown\s+(now|-[hr])', "shutdown 命令被阻止"),
    (r'\bnetsh\s+advfirewall\b', "防火墙修改被阻止"),

    # 包卸载
    (r'\bapt(-get)?\s+(remove|purge)\b', "包卸载被阻止"),
    (r'\byum\s+(remove|erase)\b', "包卸载被阻止"),
    (r'\brpm\s+-e\b', "包卸载被阻止"),
    (r'\bpip\s+uninstall\b', "pip uninstall 被阻止"),
    (r'\bnpm\s+uninstall\b', "npm uninstall 被阻止"),

    # 压缩创建（查看是允许的）
    (r'\btar\s+.*-[cC]', "压缩创建被阻止（查看用 -t 是允许的）"),
    (r'\bzip\s+-r\b', "zip 创建被阻止"),
    (r'\b7z\s+a\b', "7z 压缩创建被阻止"),
]

# 被阻止命令的原因映射
BLOCK_REASONS = {
    "rm": "请使用 delete_file 工具",
    "rmdir": "请使用 delete_file 工具",
    "shred": "请使用 delete_file 工具",
    "del": "请使用 delete_file 工具",
    "erase": "请使用 delete_file 工具",
    "tee": "请使用 write_file 工具",
    "vi": "请使用 edit_file 工具",
    "vim": "请使用 edit_file 工具",
    "nvim": "请使用 edit_file 工具",
    "nano": "请使用 edit_file 工具",
    "emacs": "请使用 edit_file 工具",
    "ed": "请使用 edit_file 工具",
    "notepad": "请使用 edit_file 工具",
    "mv": "请使用 move_file 工具",
    "rename": "请使用 move_file 工具",
    "ren": "请使用 move_file 工具",
    "cp": "请使用 copy_file 工具",
    "xcopy": "请使用 copy_file 工具",
    "robocopy": "请使用 copy_file 工具",
    "chmod": "不允许更改权限",
    "chown": "不允许更改权限",
    "chgrp": "不允许更改权限",
    "attrib": "不允许更改权限",
    "reboot": "不允许系统控制命令",
    "shutdown": "不允许系统控制命令",
    "halt": "不允许系统控制命令",
    "poweroff": "不允许系统控制命令",
    "mkfs": "不允许磁盘格式化",
    "fdisk": "不允许磁盘操作",
    "format": "不允许磁盘格式化",
    "parted": "不允许磁盘操作",
    "gdisk": "不允许磁盘操作",
    "useradd": "不允许用户管理",
    "userdel": "不允许用户管理",
    "usermod": "不允许用户管理",
    "passwd": "不允许用户管理",
    "adduser": "不允许用户管理",
    "deluser": "不允许用户管理",
    "killall": "请使用 kill_process 工具",
    "taskkill": "请使用 kill_process 工具",
    "iptables": "不允许防火墙修改",
}


def extract_commands_from_input(command: str) -> list:
    """
    从输入中提取实际命令名，处理：
    - 管道: cmd1 | cmd2
    - 命令链: cmd1 && cmd2, cmd1 ; cmd2
    - 子shell: $(cmd), `cmd`

    返回命令名列表（每个命令的第一个词）。
    """
    commands = []

    # 按管道、&&、||、; 分割
    parts = re.split(r'\s*(?:\|{1,2}|&&|;)\s*', command)

    for part in parts:
        part = part.strip()
        if not part:
            continue

        # 处理子shell $(...) 或 `...`
        subshell_match = re.search(r'\$\(([^)]+)\)|`([^`]+)`', part)
        if subshell_match:
            inner = subshell_match.group(1) or subshell_match.group(2)
            commands.extend(extract_commands_from_input(inner))

        # 获取第一个词（实际命令）
        # 跳过变量赋值如 VAR=value
        if '=' in part.split()[0] if part.split() else False:
            # 可能是 VAR=value command，获取后面的命令
            tokens = part.split()
            for i, token in enumerate(tokens):
                if '=' not in token:
                    commands.append(token)
                    break
        else:
            tokens = part.split()
            if tokens:
                # 移除前导路径并获取命令名
                cmd = tokens[0]
                # 处理路径如 /usr/bin/rm 或 C:\path\cmd.exe
                cmd_name = os.path.basename(cmd)
                # 移除 .exe 扩展名（如果存在）
                if cmd_name.lower().endswith('.exe'):
                    cmd_name = cmd_name[:-4]
                commands.append(cmd_name.lower())

    return commands


def is_command_blocked(command: str) -> tuple:
    """
    使用改进的检测检查命令是否被阻止。

    使用:
    1. 单词边界匹配单个命令（不会在 'mcp' 中匹配 'cp'）
    2. 正则模式匹配复合危险命令

    返回:
        (is_blocked, reason)
    """
    # 首先检查正则模式（需要子串匹配）
    for pattern, reason in BLOCKED_PATTERNS:
        if re.search(pattern, command, re.IGNORECASE):
            return True, reason

    # 从输入中提取实际命令
    extracted_commands = extract_commands_from_input(command)

    # 检查任何提取的命令是否在黑名单中
    for cmd in extracted_commands:
        if cmd in BLOCKED_COMMANDS:
            reason = BLOCK_REASONS.get(cmd, "此命令不被允许")
            return True, f"'{cmd}' 被阻止。{reason}"

    return False, ""


def get_blocked_commands_str() -> str:
    """获取被阻止命令的格式化列表"""
    categories = {
        "删除 (使用 delete_file)": ["rm", "rmdir", "del", "shred"],
        "写入 (使用 write_file)": ["tee", ">", ">>"],
        "编辑 (使用 edit_file)": ["vim", "nano", "vi", "emacs"],
        "移动 (使用 move_file)": ["mv", "rename"],
        "复制 (使用 copy_file)": ["cp", "xcopy", "robocopy"],
        "系统控制": ["reboot", "shutdown", "halt"],
        "磁盘工具": ["mkfs", "fdisk", "format"],
        "权限": ["chmod", "chown", "chgrp"],
        "用户管理": ["useradd", "userdel", "passwd"],
        "进程杀死": ["killall", "taskkill", "kill -9"],
    }
    return "\n".join(f"  {cat}: {', '.join(cmds)}" for cat, cmds in categories.items())


def register_bash_tools(mcp: FastMCP):
    """
    注册 Git Bash 工具用于本地命令执行。

    使用改进的黑名单机制，使用单词边界匹配避免
    误报（例如 'mcp' 路径中的 'cp' 不会触发阻止）。
    """
    bash_path = find_git_bash()

    @mcp.tool(name="MyPC-bash")
    def bash_execute(command: str, timeout: int = 30) -> str:
        """
        在 Git Bash 中执行命令（本地 Linux 环境）。

        使用黑名单安全机制 - 危险命令被阻止，
        但大多数只读和信息命令是允许的。

        被阻止的命令（请使用文件工具代替）:
        - 删除: rm, rmdir, del → 使用 delete_file
        - 写入: tee, >, >> → 使用 write_file
        - 编辑: vim, nano → 使用 edit_file
        - 移动: mv → 使用 move_file
        - 复制: cp → 使用 copy_file
        - 系统: reboot, shutdown, mkfs, chmod 等

        允许的命令:
        - 只读: ls, cat, head, tail, grep, find, less, more
        - 信息: ps, top, df, du, free, uname, whoami, date
        - 网络: ping, curl, wget, netstat, ssh
        - 开发: git, npm, node, python, docker, docker-compose
        - 压缩查看: tar -tzf, unzip -l, zipinfo

        参数:
            command: 在 Git Bash 中执行的命令
            timeout: 命令超时时间（秒，默认 30）

        返回:
            命令输出（stdout + stderr）和退出码
        """
        if not bash_path:
            return "错误: 未找到 Git Bash。请从 https://git-scm.com/download/win 安装 Git for Windows"

        # 使用改进的检测检查黑名单
        is_blocked, reason = is_command_blocked(command)
        if is_blocked:
            return f"错误: {reason}"

        try:
            # 在 Git Bash 中执行命令
            result = subprocess.run(
                [bash_path, "-c", command],
                capture_output=True,
                text=True,
                timeout=timeout,
                encoding='utf-8',
                errors='replace'
            )

            output_lines = []
            if result.stdout:
                output_lines.append(f"STDOUT:\n{result.stdout}")
            if result.stderr:
                output_lines.append(f"STDERR:\n{result.stderr}")
            output_lines.append(f"退出码: {result.returncode}")

            return "\n".join(output_lines)

        except subprocess.TimeoutExpired:
            return f"错误: 命令在 {timeout} 秒后超时"
        except Exception as e:
            return f"执行命令时出错: {str(e)}"

    @mcp.tool(name="MyPC-bash_blocked")
    def bash_blocked() -> str:
        """
        列出所有被阻止的命令类别。

        被阻止的命令有专用工具或太危险。
        使用此工具了解哪些命令不被允许以及原因。

        返回:
            被阻止命令类别和原因的列表
        """
        return f"""被阻止的命令类别（请使用文件工具代替）:

{get_blocked_commands_str()}

对于文件操作，请改用这些工具:
- MyPC-read_file: 读取任何文件
- MyPC-write_file: 写入安全区中的文件
- MyPC-edit_file: 搜索和替换文件内容
- MyPC-copy_file: 复制文件（单个或批量）
- MyPC-move_file: 移动/重命名文件
- MyPC-delete_file: 删除文件（带回收站）

对于其他操作，大多数命令是允许的:
- 开发: git, npm, node, python, docker, docker-compose
- 信息: ls, cat, grep, find, ps, top, df, du, free
- 网络: ping, curl, wget, ssh, netstat
- 压缩查看: tar -tzf, unzip -l, less, gzip -cd"""

    @mcp.tool(name="MyPC-bash_status")
    def bash_status() -> str:
        """
        检查 Git Bash 安装状态。

        返回:
            Git Bash 路径和可用性信息
        """
        if bash_path:
            return f"""Git Bash 状态: 已安装

路径: {bash_path}

Git Bash 已准备好执行命令。

使用 MyPC-bash 执行命令，或使用 MyPC-bash_blocked 查看被阻止的命令。"""
        else:
            return """Git Bash 状态: 未找到

在以下常见位置未找到 Git Bash:
- C:\\Program Files\\Git\\bin\\bash.exe
- C:\\Program Files (x86)\\Git\\bin\\bash.exe
- D:\\Program Files\\Git\\bin\\bash.exe
- E:\\Program Files\\Git\\bin\\bash.exe

请从以下地址安装 Git for Windows:
https://git-scm.com/download/win

安装后，重启 MCP 服务器。"""
