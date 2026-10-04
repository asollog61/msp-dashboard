@echo off
REM Copies inspection photos and PDF reports from the Dropbox app folder into
REM each building's share folder:
REM   ACTIVE PROPERTIES\<building>\0_xxShare\Inspections\<date>_<time>\
REM It only copies. Nothing is deleted or moved, because the dashboard still
REM reads the photos from the app folder.
setlocal
set "DST=C:\Dropbox\ASRA Investments\Marion St Properties\ACTIVE PROPERTIES"
set "LOG=%TEMP%\MSP_Sync_Inspections.log"

set "SRC="
for %%A in ("C:\Dropbox\Apps\MSP Inspect" "C:\Dropbox\Apps\MSP Building Inspections") do (
    if exist "%%~A" set "SRC=%%~A"
)
if not defined SRC (
    echo Could not find the Dropbox app folder under C:\Dropbox\Apps.
    echo %date% %time% app folder not found >> "%LOG%"
    exit /b 1
)
if not exist "%DST%" (
    echo Could not find "%DST%".
    echo %date% %time% destination not found >> "%LOG%"
    exit /b 1
)

echo %date% %time% sync from "%SRC%" >> "%LOG%"
call :one "114 Central"       "114 Central Westfield\0_114Share"
call :one "15 South"          "15 South Street\0_15Share"
call :one "36 South"          "36 South Street\0_36Share"
call :one "1280 Springfield"  "1280-86 Springfield Ave\0_1280Share"
echo Done. Log: %LOG%
exit /b 0

:one
REM %~1 = building name used by the dashboard, %~2 = building folder\share folder
call :copy "%SRC%\%~2\Inspections" "%DST%\%~2\Inspections"
call :copy "%SRC%\ASRA Investments\Marion St Properties\ACTIVE PROPERTIES\%~2\Inspections" "%DST%\%~2\Inspections"
REM Layout used by the first version of the tab: <building>\<date>_<id>\
call :copy "%SRC%\%~1" "%DST%\%~2\Inspections"
call :copy "%SRC%\MSP Inspections\%~1" "%DST%\%~2\Inspections"
exit /b 0

:copy
if not exist "%~1" exit /b 0
robocopy "%~1" "%~2" /E /XO /R:2 /W:5 /NP /NDL /NJH /NJS >> "%LOG%"
exit /b 0
