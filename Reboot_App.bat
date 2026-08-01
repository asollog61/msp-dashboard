@echo off
setlocal
cd /d "%~dp0"

echo ========================================
echo   MSP Dashboard - Force Cold Start
echo ========================================
echo.
echo Streamlit Cloud reuses Python modules it has
echo already imported, so a normal push can leave a
echo new app.py talking to an old lease_*.py.
echo.
echo Touching requirements.txt makes Streamlit rebuild
echo the environment, which forces a genuine cold start
echo and a fresh import of every module.
echo.

if not exist "requirements.txt" (
    echo ERROR: requirements.txt not found in this folder.
    echo Run this from the msp-dashboard-src folder.
    pause
    exit /b 1
)

REM A crashed git leaves this behind and blocks every later commit.
if exist ".git\index.lock" (
    echo Clearing a stale git lock...
    del /f /q ".git\index.lock"
)

REM Replace the marker line rather than appending a new one each time,
REM so requirements.txt does not grow without limit.
set "MARK=# cold-start marker:"
findstr /v /c:"%MARK%" requirements.txt > requirements.tmp
if errorlevel 2 (
    echo ERROR: could not read requirements.txt
    del /f /q requirements.tmp 2>nul
    pause
    exit /b 1
)
echo %MARK% %date% %time:~0,8%>>requirements.tmp
move /y requirements.tmp requirements.txt >nul

git add -A
git commit -m "Force cold start %date% %time:~0,5%"
git push
if errorlevel 1 (
    echo.
    echo ERROR: the push failed. Nothing will redeploy.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Pushed. Rebuild takes ~2-3 minutes.
echo.
echo   Faster alternative: in the app, use
echo   Manage app - three dots - Reboot app.
echo ========================================
pause
