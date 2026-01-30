@echo off
title MyPC-MCP Firewall Setup
echo ========================================
echo MyPC-MCP Firewall Setup
echo ========================================
echo.
echo This script will add a Windows Firewall rule
echo to allow incoming connections on port 9999.
echo.

:: Check for administrator privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] Administrator privileges required!
    echo Please right-click and select "Run as administrator"
    pause
    exit /b 1
)

:: Check if rule already exists
netsh advfirewall firewall show rule name="MyPC-MCP" >nul 2>&1
if %errorLevel% equ 0 (
    echo [INFO] Firewall rule "MyPC-MCP" already exists.
    echo.
    set /p CHOICE="Do you want to delete and recreate it? (y/n): "
    if /i "!CHOICE!"=="y" (
        echo [INFO] Deleting existing rule...
        netsh advfirewall firewall delete rule name="MyPC-MCP" >nul
        echo [OK] Rule deleted
    ) else (
        echo [INFO] Keeping existing rule
        pause
        exit /b 0
    )
)

:: Add firewall rule
echo [INFO] Adding firewall rule...
netsh advfirewall firewall add rule name="MyPC-MCP" dir=in action=allow protocol=TCP localport=9999 profile=any >nul 2>&1

if %errorLevel% equ 0 (
    echo [OK] Firewall rule added successfully!
    echo.
    echo ========================================
    echo Firewall Rule Details:
    echo ========================================
    echo Name: MyPC-MCP
    echo Protocol: TCP
    echo Port: 9999
    echo Profile: Any (Domain, Private, Public)
    echo Action: Allow
    echo.
    echo Your server is now accessible from:
    echo - Localhost: http://localhost:9999
    echo - LAN: http://YOUR_LAN_IP:9999
    echo - WAN: http://YOUR_PUBLIC_IP:9999 (if port forwarded)
    echo.
) else (
    echo [ERROR] Failed to add firewall rule
    echo Please run this script as Administrator
)

pause
