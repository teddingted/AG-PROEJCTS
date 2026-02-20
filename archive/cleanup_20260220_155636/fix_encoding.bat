@echo off
REM =====================================================
REM Python UTF-8 Encoding Fix for Windows
REM Run this once to set environment variables
REM =====================================================

echo Setting Python UTF-8 encoding...

REM Set for current session
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

REM Set permanently for user (requires running as admin for system-wide)
setx PYTHONIOENCODING utf-8
setx PYTHONUTF8 1

echo.
echo ============================================
echo UTF-8 encoding configured successfully!
echo ============================================
echo.
echo Changes applied:
echo - PYTHONIOENCODING=utf-8
echo - PYTHONUTF8=1
echo.
echo Please RESTART your terminal for changes to take effect.
echo.
pause
