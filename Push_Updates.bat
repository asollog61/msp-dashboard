@echo off
cd /d "%~dp0"
echo ========================================
echo   MSP Dashboard - Push Updates
echo ========================================
echo.

git add -A
git commit -m "Dashboard update %date% %time:~0,5%"
git push

echo.
echo ========================================
echo   Done! App will redeploy in ~1 minute.
echo ========================================
pause
