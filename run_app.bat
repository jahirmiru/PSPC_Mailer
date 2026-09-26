@echo off
title Plansculpt Auto Mailer Pro
cd /d "%~dp0"

echo ===================================================
echo       Plansculpt Auto Mailer Pro Launcher
echo ===================================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not in PATH!
    echo Please install Python from https://www.python.org and check "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

:: Check and install requirements if needed
echo [*] Checking dependencies...
python -c "import customtkinter, pandas, openpyxl, tkinterweb" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [!] Missing dependencies detected. Installing requirements...
    echo.
    pip install -r requirements.txt
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo [ERROR] Failed to install required packages.
        pause
        exit /b 1
    )
    echo.
    echo [+] Dependencies installed successfully!
    echo.
)

:: Run GUI Application
echo [*] Starting Plansculpt Auto Mailer Pro...
python pspc_mailer.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] An error occurred while running the application.
    pause
)
