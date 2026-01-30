# MYPC-MCP

[English](#english) | [中文](#中文)

---

<a name="english"></a>
## English

A powerful Model Context Protocol (MCP) server that provides AI agents with comprehensive control over a local Windows PC. MYPC-MCP acts as an AI assistant butler, giving AI vision (screenshots), hearing (system status), hands (control), and voice (automation) capabilities.

### 🚀 Key Features

- **🖥️ Screen Tools**: Full screen, active window, and webcam capture with AI analysis
- **🪟 Window Management**: List, focus, minimize, maximize, close windows
- **📁 File Operations**: Read, write, edit, copy, move, delete with safety zones
- **🔍 File Search**: Fast search using Everything integration
- **⌨️ Keyboard & Mouse**: Text input and hotkey automation
- **🧠 Smart Detection**: Auto-detect active window's file path
- **📊 Excel Automation**: Control open Excel workbooks via xlwings
- **📝 Office Automation**: Control Word and PowerPoint via pywin32
- **🌐 SSH Remote**: Execute commands on remote Linux servers
- **📊 System Control**: Volume, power, notifications, hardware status
- **📋 Clipboard**: Read and write clipboard content
- **🐧 Bash Shell**: Execute Git Bash commands with blacklist protection

---

## 🛠️ Quick Start

### One-Line Installation (Windows)

```batch
# Double-click to run - automatically checks and installs everything
start.bat
```

Or manually:

```batch
# Install
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

# Start
python main.py
```

The server will start on `http://localhost:9999` by default.

---

## 📖 Configuration

### Basic Config (config.json)

```json
{
    "server": {
        "enabled": true,
        "name": "MyPC-MCP",
        "port": 9999,
        "host": "0.0.0.0",
        "domain": "localhost"
    },
    "safe_zones": [
        "%USERPROFILE%\\Documents",
        "%USERPROFILE%\\Downloads",
        "%USERPROFILE%\\Desktop",
        "D:\\"
    ]
}
```

### Network Access Modes

**Local Only** (default):
```json
{
    "server": {
        "host": "127.0.0.1",
        "domain": "localhost"
    }
}
```

**LAN Access**:
```json
{
    "server": {
        "host": "0.0.0.0",
        "domain": "192.168.1.100"
    }
}
```

Run `setup-firewall.bat` as Administrator to allow port 9999.

---

## 🛠️ Available Tools

### Screen & Vision
- `MyPC-take_screenshot` - Capture full screen or specific monitor
- `MyPC-screenshot_active_window` - Capture active window only
- `MyPC-take_webcam_photo` - Capture from webcam
- `MyPC-list_monitors` - List all monitors

### Window Management
- `MyPC-get_active_window` - Get active window info
- `MyPC-list_windows` - List all visible windows
- `MyPC-focus_window` - Focus a window by title/process
- `MyPC-minimize_window` - Minimize a window
- `MyPC-maximize_window` - Maximize a window
- `MyPC-close_window` - Close a window
- `MyPC-screenshot_active_window` - Screenshot active window
- `MyPC-get_browser_url` - Get browser URL from address bar
- `MyPC-get_focused_control` - Get focused UI control info

### Process Management
- `MyPC-list_processes` - List processes by CPU/memory usage
- `MyPC-kill_process` - Kill process by name or PID
- `MyPC-open_app` - Open application

### File Operations
- `MyPC-list_directory` - List directory contents
- `MyPC-read_file` - Read files (txt, docx, xlsx, pptx, pdf)
- `MyPC-write_file` - Write text to file (safe zones only)
- `MyPC-edit_file` - Search and replace in file
- `MyPC-copy_file` - Copy file (into safe zones only)
- `MyPC-move_file` - Move/rename file (safe zones only)
- `MyPC-delete_file` - Delete file (moves to Recycle Bin)
- `MyPC-get_file_info` - Get detailed file information
- `MyPC-search_files` - Search files using Everything

### File Search & Detection
- `MyPC-search_files` - Fast file search via Everything
- `MyPC-detect_active_file` - Smart detection of active window's file

### Excel Automation
- `MyPC-execute_excel_code` - Execute xlwings code in active Excel
- `MyPC-list_excel_books` - List all open Excel workbooks

### Office Automation
- `MyPC-execute_word_code` - Execute pywin32 code in active Word
- `MyPC-execute_ppt_code` - Execute pywin32 code in active PowerPoint

### Input Automation
- `type_text` - Type text (supports Chinese via clipboard)
- `hotkey` - Press hotkey combinations

### Clipboard
- `MyPC-get_clipboard` - Get clipboard content
- `MyPC-set_clipboard` - Set clipboard content

### System Control
- `MyPC-get_system_status` - Get CPU, memory, battery status
- `MyPC-set_volume` / `MyPC-get_volume` - Volume control
- `MyPC-lock_screen` - Lock workstation
- `MyPC-sleep_display` - Turn off display
- `MyPC-hibernate` - Hibernate system

### Notifications
- `MyPC-show_notification` - Show Windows Toast notification

### Hardware
- `MyPC-get_hardware_status` - CPU, GPU, memory, disk info

### Bash Shell
- `MyPC-bash` - Execute Git Bash commands
- `MyPC-bash_blocked` - List blocked commands
- `MyPC-bash_status` - Check Bash installation

### SSH Remote
- `MyPC-ssh_list_hosts` - List configured SSH hosts
- `MyPC-ssh_execute` - Execute remote command
- `MyPC-ssh_test_connection` - Test SSH connection
- `MyPC-ssh_allowed_commands` - List allowed commands

### Utilities
- `MyPC-delay` - Delay execution (1-120 seconds)

---

## 📚 Usage Examples

### Example 1: Read and Edit a File

```python
# List files in Documents
MyPC-list_directory(path="%USERPROFILE%\\Documents")

# Read a file
MyPC-read_file(path="%USERPROFILE%\\Documents\\notes.txt")

# Edit the file
MyPC-edit_file(
    path="%USERPROFILE%\\Documents\\notes.txt",
    search_text="old text",
    replace_text="new text"
)
```

### Example 2: Browser Automation

```python
# Get current browser URL
MyPC-get_browser_url()

# Take a screenshot with AI analysis
MyPC-take_screenshot(display_index=1, ai_analysis=True)

# Copy URL to clipboard
MyPC-set_clipboard(text="https://example.com")
```

### Example 3: Window Management

```python
# List all windows
MyPC-list_windows()

# Focus Notepad
MyPC-focus_window(title="Notepad")

# Type something
type_text(text="Hello World", enter=True)

# Save and close
hotkey(keys=["ctrl", "s"])
MyPC-close_window(title="Notepad")
```

### Example 4: File Search

```python
# Search for Python files
MyPC-search_files(query="*.py", limit=20)

# Search in specific directory
MyPC-search_files(query="project notes ext:pdf")
```

### Example 5: Smart File Detection

```python
# Detect which file is currently open
MyPC-detect_active_file()
# Returns: {
#   "path": "D:\\Projects\\main.py",
#   "filename": "main.py",
#   "software": "Python",
#   "strategy": "标题路径提取"
# }
```

### Example 6: Remote SSH Execution

```python
# List SSH hosts
MyPC-ssh_list_hosts()

# Execute command on remote server
MyPC-ssh_execute(host="MyServer", command="docker ps")
```

---

## 🔒 Security

### Safe Zones

File operations are categorized by permission level:

- **Read Operations**: Any directory
- **Write Operations**: Only in configured safe zones
- **Copy Operations**: Can copy INTO safe zones only

### SSH Security

- Command whitelist mechanism
- Password and key file authentication
- Encoded connection support

### Data Protection

- Recycle Bin for file deletions
- No permanent data loss by default
- Path validation and normalization

---

## 🔧 Advanced Configuration

### Environment Variables

Config supports environment variable expansion:
- `%USERPROFILE%` - User profile directory
- `%APPDATA%` - Application data
- `%TEMP%` - Temporary folder
- `~` - Home directory

### Everything Search

To enable fast file search, install [Everything](https://www.voidtools.com/) and configure path in `config.json`:

```json
{
    "paths": {
        "everything": [
            "C:\\Program Files\\Everything\\es.exe",
            "D:\\APP\\Everything\\es.exe"
        ]
    }
}
```

### Git Bash

Configure Git Bash path in `config.json`:

```json
{
    "paths": {
        "git_bash": [
            "C:\\Program Files\\Git\\bin\\bash.exe"
        ]
    }
}
```

### AI Analysis (VLM)

Configure Vision Language Model for screenshot analysis:

```json
{
    "vlm": {
        "enabled": true,
        "base_url": "https://open.bigmodel.cn/api/paas/v4",
        "api_key": "your_api_key_here",
        "model": "glm-4.6v",
        "prompt": "Please identify all text in this image."
    }
}
```

---

## 📁 Project Structure

```
MYPC-MCP/
├── main.py                   # Server entry point
├── config.example.json       # Configuration template
├── config.json              # Your configuration (gitignored)
├── requirements.txt          # Python dependencies
├── start.bat               # Quick start script
├── install.bat             # Installation script
├── stop.bat                # Stop server script
├── setup-firewall.bat      # Firewall configuration
├── tools/                   # Tool modules
│   ├── screen.py            # Screen & webcam tools
│   ├── system.py            # System control tools
│   ├── files.py             # File management tools
│   ├── window.py            # Window management tools
│   ├── search.py            # File search tools
│   ├── ssh.py               # SSH remote tools
│   ├── bash.py              # Git Bash tools
│   ├── keyboard_mouse.py    # Keyboard & mouse automation
│   ├── detector.py          # Smart file detection
│   ├── detect_active_file.py # File detection implementation
│   ├── excel.py             # Excel automation (xlwings)
│   └── office.py            # Office automation (Word/PPT)
├── utils/                   # Utility modules
│   └── config.py            # Configuration loader
└── screenshots/             # Screenshot storage
```

---

## 🌐 Network Access

### Local Access

```bash
# Start server
python main.py

# Access from same machine
http://localhost:9999
```

### LAN Access

1. Configure firewall (run `setup-firewall.bat` as Administrator)
2. Edit `config.json`:
   ```json
   {
       "server": {
           "host": "0.0.0.0",
           "domain": "YOUR_LAN_IP"
       }
   }
   ```
3. Access from LAN: `http://YOUR_LAN_IP:9999`

### WAN Access

1. Configure router port forwarding (External 9999 → Your PC:9999)
2. Run `setup-firewall.bat` as Administrator
3. Configure domain or use public IP

See [NETWORK.md](NETWORK.md) for detailed guide.

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit Pull Request.

---

## 📄 License

MIT License

---

<a name="中文"></a>
## 中文

一个强大的基于 Model Context Protocol (MCP) 的 Windows 电脑控制服务器。MYPC-MCP 充当 AI 智能体的管家，赋予 AI 视觉（截图）、听觉（系统状态）、手脚（控制）和语音（自动化）能力。

### 🚀 核心功能

- **🖥️ 屏幕工具**：全屏、活动窗口、网络摄像头截图，支持 AI 分析
- **🪟 窗口管理**：列出、聚焦、最小化、最大化、关闭窗口
- **📁 文件操作**：读取、写入、编辑、复制、移动、删除文件，具备安全区保护
- **🔍 文件搜索**：通过 Everything 集成实现快速文件搜索
- **⌨️ 键鼠自动化**：文本输入和快捷键自动化
- **🧠 智能检测**：自动检测活动窗口关联的文件路径
- **📊 Excel 自动化**：通过 xlwings 控制已打开的 Excel 工作簿
- **📝 Office 自动化**：通过 pywin32 控制 Word 和 PowerPoint
- **🌐 SSH 远程**：在远程 Linux 服务器上执行命令
- **📊 系统控制**：音量、电源、通知、硬件状态
- **📋 剪贴板**：读取和写入剪贴板内容
- **🐧 Bash Shell**：执行 Git Bash 命令（带黑名单保护）

---

## 🛠️ 快速开始

### 一键安装（Windows）

```batch
# 双击运行 - 自动检查并安装所有依赖
start.bat
```

或手动安装：

```batch
# 安装
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

# 启动
python main.py
```

服务器默认在 `http://localhost:9999` 上启动。

---

## 📖 配置说明

### 基础配置 (config.json)

```json
{
    "server": {
        "enabled": true,
        "name": "MyPC-MCP",
        "port": 9999,
        "host": "0.0.0.0",
        "domain": "localhost"
    },
    "safe_zones": [
        "%USERPROFILE%\\Documents",
        "%USERPROFILE%\\Downloads",
        "%USERPROFILE%\\Desktop",
        "D:\\"
    ]
}
```

### 网络访问模式

**仅本地访问**（默认）：
```json
{
    "server": {
        "host": "127.0.0.1",
        "domain": "localhost"
    }
}
```

**局域网访问**：
```json
{
    "server": {
        "host": "0.0.0.0",
        "domain": "192.168.1.100"
    }
}
```

以管理员身份运行 `setup-firewall.bat` 允许端口 9999。

---

## 🛠️ 可用工具

### 屏幕与视觉
- `MyPC-take_screenshot` - 截取全屏或指定显示器
- `MyPC-screenshot_active_window` - 仅截取活动窗口
- `MyPC-take_webcam_photo` - 从摄像头拍照
- `MyPC-list_monitors` - 列出所有显示器

### 窗口管理
- `MyPC-get_active_window` - 获取活动窗口信息
- `MyPC-list_windows` - 列出所有可见窗口
- `MyPC-focus_window` - 按标题/进程聚焦窗口
- `MyPC-minimize_window` - 最小化窗口
- `MyPC-maximize_window` - 最大化窗口
- `MyPC-close_window` - 关闭窗口
- `MyPC-screenshot_active_window` - 截取活动窗口
- `MyPC-get_browser_url` - 获取浏览器地址栏 URL
- `MyPC-get_focused_control` - 获取焦点控件信息

### 进程管理
- `MyPC-list_processes` - 按 CPU/内存列出进程
- `MyPC-kill_process` - 按名称或 PID 终止进程
- `MyPC-open_app` - 打开应用程序

### 文件操作
- `MyPC-list_directory` - 列出目录内容
- `MyPC-read_file` - 读取文件（txt、docx、xlsx、pptx、pdf）
- `MyPC-write_file` - 写入文本文件（仅安全区）
- `MyPC-edit_file` - 文件搜索替换
- `MyPC-copy_file` - 复制文件（仅进入安全区）
- `MyPC-move_file` - 移动/重命名文件（仅安全区）
- `MyPC-delete_file` - 删除文件（移入回收站）
- `MyPC-get_file_info` - 获取文件详细信息
- `MyPC-search_files` - 使用 Everything 搜索文件

### 文件搜索与检测
- `MyPC-search_files` - 通过 Everything 快速搜索文件
- `MyPC-detect_active_file` - 智能检测活动窗口的文件

### Excel 自动化
- `MyPC-execute_excel_code` - 在已打开的 Excel 中执行 xlwings 代码
- `MyPC-list_excel_books` - 列出所有打开的 Excel 工作簿

### Office 自动化
- `MyPC-execute_word_code` - 在已打开的 Word 中执行 pywin32 代码
- `MyPC-execute_ppt_code` - 在已打开的 PowerPoint 中执行 pywin32 代码

### 输入自动化
- `type_text` - 输入文本（支持通过剪贴板输入中文）
- `hotkey` - 按快捷键组合

### 剪贴板
- `MyPC-get_clipboard` - 获取剪贴板内容
- `MyPC-set_clipboard` - 设置剪贴板内容

### 系统控制
- `MyPC-get_system_status` - 获取 CPU、内存、电池状态
- `MyPC-set_volume` / `MyPC-get_volume` - 音量控制
- `MyPC-lock_screen` - 锁定工作站
- `MyPC-sleep_display` - 关闭显示器
- `MyPC-hibernate` - 休眠系统

### 通知
- `MyPC-show_notification` - 显示 Windows Toast 通知

### 硬件
- `MyPC-get_hardware_status` - CPU、GPU、内存、磁盘信息

### Bash Shell
- `MyPC-bash` - 执行 Git Bash 命令
- `MyPC-bash_blocked` - 列出被阻止的命令
- `MyPC-bash_status` - 检查 Bash 安装

### SSH 远程
- `MyPC-ssh_list_hosts` - 列出配置的 SSH 主机
- `MyPC-ssh_execute` - 执行远程命令
- `MyPC-ssh_test_connection` - 测试 SSH 连接
- `MyPC-ssh_allowed_commands` - 列出允许的命令

### 实用工具
- `MyPC-delay` - 延迟执行（1-120 秒）

---

## 📚 使用场景示例

### 场景 1：读取和编辑文件

```python
# 列出文档目录的文件
MyPC-list_directory(path="%USERPROFILE%\\Documents")

# 读取文件
MyPC-read_file(path="%USERPROFILE%\\Documents\\笔记.txt")

# 编辑文件
MyPC-edit_file(
    path="%USERPROFILE%\\Documents\\笔记.txt",
    search_text="旧文本",
    replace_text="新文本"
)
```

### 场景 2：浏览器自动化

```python
# 获取当前浏览器 URL
MyPC-get_browser_url()

# 截图并 AI 分析
MyPC-take_screenshot(display_index=1, ai_analysis=True)

# 复制 URL 到剪贴板
MyPC-set_clipboard(text="https://example.com")
```

### 场景 3：窗口管理

```python
# 列出所有窗口
MyPC-list_windows()

# 聚焦记事本
MyPC-focus_window(title="记事本")

# 输入文本
type_text(text="你好世界", enter=True)

# 保存并关闭
hotkey(keys=["ctrl", "s"])
MyPC-close_window(title="记事本")
```

### 场景 4：文件搜索

```python
# 搜索 Python 文件
MyPC-search_files(query="*.py", limit=20)

# 在特定目录搜索
MyPC-search_files(query="项目笔记 ext:pdf")
```

### 场景 5：智能文件检测

```python
# 检测当前打开的文件
MyPC-detect_active_file()
# 返回: {
#   "path": "D:\\Projects\\main.py",
#   "filename": "main.py",
#   "software": "Python",
#   "strategy": "标题路径提取"
# }
```

### 场景 6：远程 SSH 执行

```python
# 列出 SSH 主机
MyPC-ssh_list_hosts()

# 在远程服务器执行命令
MyPC-ssh_execute(host="MyServer", command="docker ps")
```

---

## 🔒 安全性

### 安全区

文件操作按权限级别分类：

- **只读操作**：任意目录
- **写入操作**：仅在配置的安全区内
- **复制操作**：只能复制进入安全区

### SSH 安全

- 命令白名单机制
- 密码和密钥文件认证
- 编码连接支持

### 数据保护

- 文件删除移入回收站
- 默认无永久数据丢失
- 路径验证和规范化

---

## 🔧 高级配置

### 环境变量

配置支持环境变量扩展：
- `%USERPROFILE%` - 用户目录
- `%APPDATA%` - 应用数据目录
- `%TEMP%` - 临时文件夹
- `~` - 主目录

### Everything 搜索

启用快速文件搜索，安装 [Everything](https://www.voidtools.com/) 并在 `config.json` 中配置路径：

```json
{
    "paths": {
        "everything": [
            "C:\\Program Files\\Everything\\es.exe",
            "D:\\APP\\Everything\\es.exe"
        ]
    }
}
```

### Git Bash

在 `config.json` 中配置 Git Bash 路径：

```json
{
    "paths": {
        "git_bash": [
            "C:\\Program Files\\Git\\bin\\bash.exe"
        ]
    }
}
```

### AI 分析（VLM）

配置视觉语言模型用于截图分析：

```json
{
    "vlm": {
        "enabled": true,
        "base_url": "https://open.bigmodel.cn/api/paas/v4",
        "api_key": "your_api_key_here",
        "model": "glm-4.6v",
        "prompt": "请识别这张图片中的所有文字。"
    }
}
```

---

## 📁 项目结构

```
MYPC-MCP/
├── main.py                   # 服务器入口
├── config.example.json       # 配置模板
├── config.json              # 配置文件（gitignored）
├── requirements.txt          # Python 依赖
├── start.bat               # 快速启动脚本
├── install.bat             # 安装脚本
├── stop.bat                # 停止服务脚本
├── setup-firewall.bat      # 防火墙配置
├── tools/                   # 工具模块
│   ├── screen.py            # 屏幕和摄像头工具
│   ├── system.py            # 系统控制工具
│   ├── files.py             # 文件管理工具
│   ├── window.py            # 窗口管理工具
│   ├── search.py            # 文件搜索工具
│   ├── ssh.py               # SSH 远程工具
│   ├── bash.py              # Git Bash 工具
│   ├── keyboard_mouse.py    # 键鼠自动化
│   ├── detector.py          # 智能文件检测
│   ├── detect_active_file.py # 文件检测实现
│   ├── excel.py             # Excel 自动化 (xlwings)
│   └── office.py            # Office 自动化 (Word/PPT)
├── utils/                   # 工具模块
│   └── config.py            # 配置加载器
└── screenshots/             # 截图存储
```

---

## 🌐 网络访问

### 本地访问

```bash
# 启动服务器
python main.py

# 本机访问
http://localhost:9999
```

### 局域网访问

1. 配置防火墙（以管理员身份运行 `setup-firewall.bat`）
2. 编辑 `config.json`：
   ```json
   {
       "server": {
           "host": "0.0.0.0",
           "domain": "你的局域网IP"
       }
   }
   ```
3. 从局域网访问：`http://你的局域网IP:9999`

### 公网访问

1. 配置路由器端口转发（外部 9999 → 你的电脑:9999）
2. 以管理员身份运行 `setup-firewall.bat`
3. 配置域名或使用公网 IP

详细指南请参阅 [NETWORK.md](NETWORK.md)。

---

## 🤝 贡献

欢迎贡献！随时提交 Pull Request。

---

## 📄 许可证

MIT License
