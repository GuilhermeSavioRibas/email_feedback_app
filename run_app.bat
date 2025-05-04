@echo off
title Email Feedback App - Initializing...

REM
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not in the PATH.
    echo Please install Python 3.12+ from: https://www.python.org/downloads/
    pause
    exit /b
)

echo.
echo Checking and installing dependencies...
python -m pip install --upgrade pip
python -m pip install pandas openpyxl

echo.
echo Launching the system...
python main.py

echo.
pause
