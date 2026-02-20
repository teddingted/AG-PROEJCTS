@echo off
cd /d "%~dp0"
echo Checking dependencies...
python -c "import PySide6" >nul 2>&1
if %errorlevel% neq 0 (
    echo Dependencies not found. Installing...
    pip install -r requirements.txt
)
python -c "import cv2" >nul 2>&1
if %errorlevel% neq 0 (
    echo OpenCV not found. Installing...
    pip install -r requirements.txt
)

echo Starting AutoPlotDigitizer V2...
python main.py
if %errorlevel% neq 0 (
    echo Application crashed.
    pause
)
