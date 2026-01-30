@echo off
echo ========================================
echo MyPC-MCP Installation
echo ========================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.10+
    pause
    exit /b 1
)

:: Create virtual environment
if not exist venv (
    echo [1/2] Creating virtual environment...
    python -m venv venv
    echo [OK] Virtual environment created
) else (
    echo [1/2] Virtual environment already exists
)

:: Activate and install
echo [2/2] Installing dependencies...
call venv\Scripts\activate.bat
pip install -r requirements.txt
echo [OK] Dependencies installed

:: Create config
if not exist config.json (
    if exist config.example.json (
        copy config.example.json config.json >nul
        echo [OK] config.json created
    )
)

echo.
echo ========================================
echo Installation complete!
echo ========================================
echo.
echo Starting server in 3 seconds...
timeout /t 3 >nul

:: Start the server
call start.bat
