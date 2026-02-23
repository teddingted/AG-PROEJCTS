@echo off
echo Starting HiWay Calculator...
cd /d "%~dp0"

:: Open the HTML file in the default browser or msedge
start "" msedge.exe --app="%cd%\HiWayCalculator.html"
if %errorlevel% neq 0 (
    start "" "HiWayCalculator.html"
)

echo Calculator launched.
