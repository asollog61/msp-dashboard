@echo off
cd /d "%~dp0"
echo ========================================
echo   MSP Dashboard - Push Updates
echo ========================================
echo.

REM A crashed or interrupted git leaves this behind, and every later commit
REM then fails with "Unable to create index.lock: File exists".
if exist ".git\index.lock" (
    echo Clearing a stale git lock...
    del /f /q ".git\index.lock"
)

git add -A
git commit -m "Dashboard update %date% %time:~0,5%"
git push
if errorlevel 1 (
    echo.
    echo ========================================
    echo   PUSH FAILED - nothing will redeploy.
    echo ========================================
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Done! App will redeploy in ~1 minute.
echo.
echo   If the app errors with a missing name,
echo   the module cache is stale - run
echo   Reboot_App.bat or use Manage app - Reboot.
echo ========================================
pause
