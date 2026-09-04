@echo off
setlocal EnableExtensions
cd /d "%~dp0"
set "LOCK=.git\index.lock"

echo ========================================
echo   MSP Dashboard - Git Unlock / Clear
echo ========================================
echo.

if not exist "%LOCK%" (
    echo No index.lock is present. The repository is already clear.
    pause
    exit /b 0
)

echo Found: %CD%\%LOCK%
echo.
echo Active Git-related processes:
tasklist /fi "imagename eq git.exe" /fo table
tasklist /fi "imagename eq git-remote-https.exe" /fo table
tasklist /fi "imagename eq git-credential-manager.exe" /fo table
echo.

REM First try a normal deletion. This is safe for a stale lock.
del /f /q "%LOCK%" >nul 2>&1
if not exist "%LOCK%" (
    echo Stale lock cleared successfully.
    pause
    exit /b 0
)

echo The lock is actively held by another Windows process.
echo Close VS Code, GitHub Desktop, terminals, and File Explorer windows
 echo opened in this repository before continuing.
echo.
choice /c YN /m "Force-close Git processes and retry"
if errorlevel 2 goto :blocked

echo.
echo Closing Git processes...
taskkill /f /im git.exe >nul 2>&1
taskkill /f /im git-remote-https.exe >nul 2>&1
taskkill /f /im git-credential-manager.exe >nul 2>&1
timeout /t 2 /nobreak >nul

del /f /q "%LOCK%" >nul 2>&1
if not exist "%LOCK%" (
    echo Lock cleared. You can now run Push_Updates.bat.
    pause
    exit /b 0
)

:blocked
echo.
echo Still locked. Dropbox or another app is likely holding the file.
echo Pause Dropbox, close any app using this repo, then run this file again.
echo Lock path: %CD%\%LOCK%
pause
exit /b 1
