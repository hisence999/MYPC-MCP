@echo off
title MyPC-MCP
setlocal enabledelayedexpansion

echo ========================================
echo MyPC-MCP Quick Start
echo ========================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.10+
    pause
    exit /b 1
)

:: Create virtual environment if not exists
if not exist venv (
    echo [1/4] Creating virtual environment...
    python -m venv venv
    echo [OK] Virtual environment created
) else (
    echo [1/4] Virtual environment exists
)

:: Activate virtual environment
call venv\Scripts\activate.bat

:: Check if dependencies are installed
echo [2/4] Checking dependencies...
set NEED_INSTALL=0

:: Check critical packages
for %%p in (mcp mss Pillow pycaw comtypes psutil opencv-python pyautogui pyperclip pywin32 httpx) do (
    python -c "import %%p" >nul 2>&1
    if errorlevel 1 (
        echo [MISSING] %%p
        set NEED_INSTALL=1
    )
)

:: Install if needed
if !NEED_INSTALL! equ 1 (
    echo.
    echo [INFO] Installing missing dependencies...
    pip install -r requirements.txt
    echo [OK] Dependencies installed
) else (
    echo [OK] All dependencies installed
)

:: Create config if not exists
echo [3/4] Checking config...
if not exist config.json (
    if exist config.example.json (
        copy config.example.json config.json >nul
        echo [OK] config.json created from example
    ) else (
        echo [WARNING] No config file found
    )
) else (
    echo [OK] config.json exists
)

:: Check firewall rule
echo [4/4] Checking firewall...
netsh advfirewall firewall show rule name="MyPC-MCP" >nul 2>&1
if errorlevel 1 (
    echo [WARNING] Firewall rule not found
    echo.
    echo To allow external access, run setup-firewall.bat as Administrator
    echo Or manually add firewall rule for port 9999
) else (
    echo [OK] Firewall rule exists
)

echo.
echo ========================================
echo Network Access Information:
echo ========================================
echo.
echo Local access:   http://localhost:9999
echo.
echo For LAN access:
echo   1. Find your LAN IP: run 'ipconfig' in Command Prompt
echo   2. Access via: http://YOUR_LAN_IP:9999
echo   3. Make sure firewall allows port 9999
echo.
echo For WAN access:
echo   1. Setup port forwarding on router: External 9999 -^> Your PC:9999
echo   2. Run setup-firewall.bat as Administrator
echo   3. Access via: http://YOUR_PUBLIC_IP:9999
echo.
echo Press Ctrl+C to stop the server
echo ========================================
echo.

python main.py

:: If server exits with error, pause
if errorlevel 1 (
    echo.
    echo [ERROR] Server exited with error code: %errorlevel%
    pause
)
