@echo off
echo Stopping MyPC-MCP...

:: Try to kill python.exe process with window title containing MyPC-MCP
for /f "tokens=2" %%i in ('tasklist /fi "windowtitle eq MyPC-MCP*" ^| find "python.exe"') do (
    taskkill /f /pid %%i >nul 2>&1
)

echo Server stopped
pause
