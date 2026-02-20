@echo off
echo Starting Project Organization...

rem Define Source and Destination Paths
set "SOURCE_ROOT=C:\Users\Admin\Desktop\AG-BEGINNING"
set "DEST_ROOT=C:\Users\Admin\Desktop\AG-REPOSITORY"
set "AUTO_PLOT_DEST=%DEST_ROOT%\MAIN001 - AUTO PLOT DIGITIZER"
set "HYSYS_DEST=%DEST_ROOT%\SIDE001 - HYSYS AUTOMATION"

echo --------------------------------------------------
echo Organizing Auto Plot Digitizer...
echo --------------------------------------------------

if not exist "%AUTO_PLOT_DEST%" mkdir "%AUTO_PLOT_DEST%"
xcopy "%SOURCE_ROOT%\AutoPlotDigitizerV2_Windows_Port" "%AUTO_PLOT_DEST%" /E /H /C /I /Y

echo --------------------------------------------------
echo Organizing Hysys Automation (Folder)...
echo --------------------------------------------------

if not exist "%HYSYS_DEST%\hysys_automation" mkdir "%HYSYS_DEST%\hysys_automation"
xcopy "%SOURCE_ROOT%\hysys_automation" "%HYSYS_DEST%\hysys_automation" /E /H /C /I /Y

echo --------------------------------------------------
echo Organizing Hysys Automation (Root Scripts)...
echo --------------------------------------------------

copy /Y "%SOURCE_ROOT%\hysys_optimizer_unified.py" "%HYSYS_DEST%\"
copy /Y "%SOURCE_ROOT%\hysys_optimizer_dispatch.py" "%HYSYS_DEST%\"
copy /Y "%SOURCE_ROOT%\hysys_optimizer_hybrid.py" "%HYSYS_DEST%\"
copy /Y "%SOURCE_ROOT%\hysys_optimizer_multidim.py" "%HYSYS_DEST%\"
copy /Y "%SOURCE_ROOT%\hysys_optimizer_acc.py" "%HYSYS_DEST%\"
copy /Y "%SOURCE_ROOT%\hysys_optimizer_2d.py" "%HYSYS_DEST%\"
copy /Y "%SOURCE_ROOT%\test_hysys_connection.py" "%HYSYS_DEST%\"

echo --------------------------------------------------
echo Operation Complete.
echo --------------------------------------------------
pause
