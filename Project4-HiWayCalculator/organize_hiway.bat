@echo off
echo Organizing HiWayCalculator...

set "SOURCE=C:\Users\Admin\Desktop\AG-BEGINNING\HiWayCalculator"
set "DEST=C:\Users\Admin\Desktop\AG-REPOSITORY\SIDE002 - HiWayCalculator"

if not exist "%DEST%" mkdir "%DEST%"

copy /Y "%SOURCE%\HiWayCalculator.html" "%DEST%\"
copy /Y "%SOURCE%\README.md" "%DEST%\"

echo files copied.
pause
