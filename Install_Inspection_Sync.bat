@echo off
REM Run once. Schedules Sync_Inspections.bat to run every night at 2:00 AM
REM under your Windows account, then runs it once now.
cd /d "%~dp0"
schtasks /Create /TN "MSP Inspection Sync" /TR "\"%~dp0Sync_Inspections.bat\"" /SC DAILY /ST 02:00 /F
if errorlevel 1 (
    echo.
    echo Could not create the scheduled task.
    pause
    exit /b 1
)
echo.
echo Scheduled. Running the first sync now...
call "%~dp0Sync_Inspections.bat"
pause
