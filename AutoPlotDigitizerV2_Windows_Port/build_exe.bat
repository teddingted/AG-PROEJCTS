@echo off
cd /d "%~dp0"
echo Checking for PyInstaller...
pyinstaller --version >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller not found. Installing...
    pip install pyinstaller
)

echo Building AutoPlotDigitizer V2 Executable...
pyinstaller -y AutoPlotDigitizer.spec
if %errorlevel% == 0 (
    echo.
    echo Build SUCCESS!
    echo Executable is located in the dist folder.
    echo.
    echo Running app...
    call custom_run.bat
) else (
    echo.
    echo Build FAILED.
    pause
)
