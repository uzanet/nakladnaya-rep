@echo off
title Auto-Print Excel

where python >nul 2>&1
if errorlevel 1 (
    echo Python not found. Please install Python 3.8+ and add it to PATH.
    pause
    exit /b 1
)

python -m pip install -r requirements.txt --quiet

pythonw auto_print.py 2>nul
if errorlevel 1 (
    python auto_print.py
)
