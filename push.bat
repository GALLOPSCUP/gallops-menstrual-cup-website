@echo off
echo =========================================
echo    GALLOPS CUP - GitHub Push Script
echo =========================================
echo.

WHERE git >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Git is not installed!
    echo Download from: https://git-scm.com/downloads
    pause
    exit /b
)

cd /d "d:\GALLOPS CUP WEBSITE"

echo [INFO] Adding all changes...
git add -A

echo.
set /p MSG="Enter commit message (or press Enter for default): "
if "%MSG%"=="" set MSG=Update website files

echo [INFO] Committing: %MSG%
git commit -m "%MSG%"

echo [INFO] Pushing to GitHub...
git push origin main

echo.
echo =========================================
echo  SUCCESS! Changes pushed to GitHub.
echo  Your live site will update in 1-2 min.
echo =========================================
pause
