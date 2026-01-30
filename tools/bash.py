import subprocess
import os
import re
import shlex
from mcp.server.fastmcp import FastMCP

# Find Git Bash executable
def find_git_bash() -> str:
    """Find Git Bash executable path."""
    possible_paths = [
        r"C:\Program Files\Git\bin\bash.exe",
        r"C:\Program Files (x86)\Git\bin\bash.exe",
        r"D:\Program Files\Git\bin\bash.exe",
        r"E:\Program Files\Git\bin\bash.exe",
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return path

    # Try to find via where command
    try:
        result = subprocess.run(
            ["where", "bash.exe"],
            capture_output=True,
            text=True,
            shell=True
        )
        if result.returncode == 0 and result.stdout:
            return result.stdout.strip().split('\n')[0]
    except Exception:
        pass

    return None


# =============================================================================
# IMPROVED BLACKLIST SYSTEM
# =============================================================================

# Single-word commands that should be blocked (matched as whole words only)
BLOCKED_COMMANDS = {
    # DELETE OPERATIONS (use delete_file tool instead)
    "rm", "rmdir", "shred", "del", "erase",
    
    # WRITE OPERATIONS (use write_file tool instead)
    "tee",
    
    # EDITORS (use edit_file tool instead)
    "vi", "vim", "nvim", "nano", "emacs", "ed", "notepad",
    
    # MOVE/RENAME (use move_file tool instead)
    "mv", "rename", "ren",
    
    # COPY (use copy_file tool instead)
    "cp", "xcopy", "robocopy",
    
    # SYSTEM CONTROL (too dangerous)
    "reboot", "shutdown", "halt", "poweroff",
    
    # DISK OPERATIONS (too dangerous)
    "mkfs", "fdisk", "format", "parted", "gdisk",
    
    # PERMISSION CHANGES (can break file access)
    "chmod", "chown", "chgrp", "attrib",
    
    # USER MANAGEMENT (security risk)
    "useradd", "userdel", "usermod", "passwd", "adduser", "deluser",
    
    # FORCE KILL (can kill important processes)
    "killall", "taskkill",
    
    # NETWORK FIREWALL (can block access)
    "iptables",
}

# Multi-word patterns that should be blocked (exact substring match is OK for these)
BLOCKED_PATTERNS = [
    # Output redirection
    (r'\s+>\s+', "Output redirection is blocked. Use write_file tool instead"),
    (r'\s+>>\s+', "Append redirection is blocked. Use write_file tool instead"),
    (r'\s+2>\s+', "Stderr redirection is blocked. Use write_file tool instead"),
    
    # Dangerous compound commands
    (r'\bkill\s+-9\b', "Force kill (-9) is blocked"),
    (r'\bpkill\s+-9\b', "Force pkill (-9) is blocked"),
    (r'\binit\s+[06]\b', "init 0/6 is blocked"),
    (r'\bdd\s+if=', "dd write operations are blocked"),
    (r'\bsystemctl\s+(poweroff|reboot|halt)\b', "systemctl power commands are blocked"),
    (r'\bshutdown\s+(now|-[hr])', "shutdown commands are blocked"),
    (r'\bnetsh\s+advfirewall\b', "Firewall modification is blocked"),
    
    # Package removal
    (r'\bapt(-get)?\s+(remove|purge)\b', "Package removal is blocked"),
    (r'\byum\s+(remove|erase)\b', "Package removal is blocked"),
    (r'\brpm\s+-e\b', "Package removal is blocked"),
    (r'\bpip\s+uninstall\b', "pip uninstall is blocked"),
    (r'\bnpm\s+uninstall\b', "npm uninstall is blocked"),
    
    # Archive creation (but viewing is allowed)
    (r'\btar\s+.*-[cC]', "Archive creation is blocked (viewing with -t is allowed)"),
    (r'\bzip\s+-r\b', "zip creation is blocked"),
    (r'\b7z\s+a\b', "7z archive creation is blocked"),
]

# Reason mapping for single-word blocked commands
BLOCK_REASONS = {
    "rm": "Use delete_file tool instead",
    "rmdir": "Use delete_file tool instead",
    "shred": "Use delete_file tool instead",
    "del": "Use delete_file tool instead",
    "erase": "Use delete_file tool instead",
    "tee": "Use write_file tool instead",
    "vi": "Use edit_file tool instead",
    "vim": "Use edit_file tool instead",
    "nvim": "Use edit_file tool instead",
    "nano": "Use edit_file tool instead",
    "emacs": "Use edit_file tool instead",
    "ed": "Use edit_file tool instead",
    "notepad": "Use edit_file tool instead",
    "mv": "Use move_file tool instead",
    "rename": "Use move_file tool instead",
    "ren": "Use move_file tool instead",
    "cp": "Use copy_file tool instead",
    "xcopy": "Use copy_file tool instead",
    "robocopy": "Use copy_file tool instead",
    "chmod": "Permission changes are not allowed",
    "chown": "Permission changes are not allowed",
    "chgrp": "Permission changes are not allowed",
    "attrib": "Permission changes are not allowed",
    "reboot": "System control commands are not allowed",
    "shutdown": "System control commands are not allowed",
    "halt": "System control commands are not allowed",
    "poweroff": "System control commands are not allowed",
    "mkfs": "Disk formatting is not allowed",
    "fdisk": "Disk operations are not allowed",
    "format": "Disk formatting is not allowed",
    "parted": "Disk operations are not allowed",
    "gdisk": "Disk operations are not allowed",
    "useradd": "User management is not allowed",
    "userdel": "User management is not allowed",
    "usermod": "User management is not allowed",
    "passwd": "User management is not allowed",
    "adduser": "User management is not allowed",
    "deluser": "User management is not allowed",
    "killall": "Use kill_process tool instead",
    "taskkill": "Use kill_process tool instead",
    "iptables": "Firewall modification is not allowed",
}


def extract_commands_from_input(command: str) -> list[str]:
    """
    Extract actual command names from the input, handling:
    - Pipes: cmd1 | cmd2
    - Command chaining: cmd1 && cmd2, cmd1 ; cmd2
    - Subshells: $(cmd), `cmd`
    
    Returns list of command names (first word of each command).
    """
    commands = []
    
    # Split by pipes, &&, ||, ;
    # This regex splits on |, &&, ||, ; while keeping the structure
    parts = re.split(r'\s*(?:\|{1,2}|&&|;)\s*', command)
    
    for part in parts:
        part = part.strip()
        if not part:
            continue
            
        # Handle subshell $(...) or `...`
        subshell_match = re.search(r'\$\(([^)]+)\)|`([^`]+)`', part)
        if subshell_match:
            inner = subshell_match.group(1) or subshell_match.group(2)
            commands.extend(extract_commands_from_input(inner))
        
        # Get the first word (the actual command)
        # Skip variable assignments like VAR=value
        if '=' in part.split()[0] if part.split() else False:
            # This might be VAR=value command, get the command after
            tokens = part.split()
            for i, token in enumerate(tokens):
                if '=' not in token:
                    commands.append(token)
                    break
        else:
            tokens = part.split()
            if tokens:
                # Remove any leading path and get just the command name
                cmd = tokens[0]
                # Handle paths like /usr/bin/rm or C:\path\cmd.exe
                cmd_name = os.path.basename(cmd)
                # Remove .exe extension if present
                if cmd_name.lower().endswith('.exe'):
                    cmd_name = cmd_name[:-4]
                commands.append(cmd_name.lower())
    
    return commands


def is_command_blocked(command: str) -> tuple[bool, str]:
    """
    Check if command is blocked using improved detection.
    
    Uses:
    1. Word-boundary matching for single commands (won't match 'cp' in 'mcp')
    2. Regex patterns for compound dangerous commands
    
    Returns:
        (is_blocked, reason)
    """
    # First, check regex patterns (these need substring matching)
    for pattern, reason in BLOCKED_PATTERNS:
        if re.search(pattern, command, re.IGNORECASE):
            return True, reason
    
    # Extract actual commands from the input
    extracted_commands = extract_commands_from_input(command)
    
    # Check if any extracted command is in the blocked list
    for cmd in extracted_commands:
        if cmd in BLOCKED_COMMANDS:
            reason = BLOCK_REASONS.get(cmd, "This command is not allowed")
            return True, f"'{cmd}' is blocked. {reason}"
    
    return False, ""


def get_blocked_commands_str() -> str:
    """Get formatted list of blocked commands."""
    categories = {
        "Delete (use delete_file)": ["rm", "rmdir", "del", "shred"],
        "Write (use write_file)": ["tee", ">", ">>"],
        "Edit (use edit_file)": ["vim", "nano", "vi", "emacs"],
        "Move (use move_file)": ["mv", "rename"],
        "Copy (use copy_file)": ["cp", "xcopy", "robocopy"],
        "System Control": ["reboot", "shutdown", "halt"],
        "Disk Tools": ["mkfs", "fdisk", "format"],
        "Permissions": ["chmod", "chown", "chgrp"],
        "User Management": ["useradd", "userdel", "passwd"],
        "Process Kill": ["killall", "taskkill", "kill -9"],
    }
    return "\n".join(f"  {cat}: {', '.join(cmds)}" for cat, cmds in categories.items())


def register_bash_tools(mcp: FastMCP):
    """
    Register Git Bash tools for local command execution.

    Uses improved blacklist with word-boundary matching to avoid
    false positives (e.g., 'cp' in 'mcp' path won't trigger block).
    """
    bash_path = find_git_bash()

    @mcp.tool(name="MyPC-bash")
    def bash_execute(command: str, timeout: int = 30) -> str:
        """
        Execute a command in Git Bash (local Linux environment).

        Uses a blacklist for security - dangerous commands are blocked,
        but most read-only and information commands are allowed.

        BLOCKED commands (use file tools instead):
        - Delete: rm, rmdir, del → use delete_file
        - Write: tee, >, >> → use write_file
        - Edit: vim, nano → use edit_file
        - Move: mv → use move_file
        - Copy: cp → use copy_file
        - System: reboot, shutdown, mkfs, chmod, etc.

        ALLOWED commands:
        - Read-only: ls, cat, head, tail, grep, find, less, more
        - Info: ps, top, df, du, free, uname, whoami, date
        - Network: ping, curl, wget, netstat, ssh
        - Dev: git, npm, node, python, docker, docker-compose
        - Compression view: tar -tzf, unzip -l, zipinfo

        Args:
            command: Command to execute in Git Bash.
            timeout: Command timeout in seconds (default 30).

        Returns:
            Command output (stdout + stderr) and exit code.
        """
        if not bash_path:
            return "Error: Git Bash not found. Please install Git for Windows from https://git-scm.com/download/win"

        # Check blacklist with improved detection
        is_blocked, reason = is_command_blocked(command)
        if is_blocked:
            return f"Error: {reason}"

        try:
            # Execute command in Git Bash
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
            output_lines.append(f"Exit Code: {result.returncode}")

            return "\n".join(output_lines)

        except subprocess.TimeoutExpired:
            return f"Error: Command timed out after {timeout} seconds"
        except Exception as e:
            return f"Error executing command: {str(e)}"

    @mcp.tool(name="MyPC-bash_blocked")
    def bash_blocked() -> str:
        """
        List all blocked command categories.

        Blocked commands have dedicated tools or are too dangerous.
        Use this to understand what commands are not allowed and why.

        Returns:
            List of blocked command categories and reasons.
        """
        return f"""Blocked Command Categories (use file tools instead):

{get_blocked_commands_str()}

For file operations, use these tools instead:
- MyPC-read_file: Read any file
- MyPC-write_file: Write to files in safe zones
- MyPC-edit_file: Search and replace in files
- MyPC-copy_file: Copy files (single or batch)
- MyPC-move_file: Move/rename files
- MyPC-delete_file: Delete files (with recycle bin)

For other operations, most commands are allowed:
- Development: git, npm, node, python, docker, docker-compose
- Information: ls, cat, grep, find, ps, top, df, du, free
- Network: ping, curl, wget, ssh, netstat
- Compression view: tar -tzf, unzip -l, less, gzip -cd"""

    @mcp.tool(name="MyPC-bash_status")
    def bash_status() -> str:
        """
        Check Git Bash installation status.

        Returns:
            Git Bash path and availability information.
        """
        if bash_path:
            return f"""Git Bash Status: Installed

Path: {bash_path}

Git Bash is ready to execute commands.

Use MyPC-bash to execute commands, or MyPC-bash_blocked to see blocked commands."""
        else:
            return """Git Bash Status: Not Found

Git Bash was not found in common locations:
- C:\\Program Files\\Git\\bin\\bash.exe
- C:\\Program Files (x86)\\Git\\bin\\bash.exe
- D:\\Program Files\\Git\\bin\\bash.exe
- E:\\Program Files\\Git\\bin\\bash.exe

Please install Git for Windows from:
https://git-scm.com/download/win

After installation, restart the MCP server."""
