@echo off
title Fiuxxed's AutoTyper — Updater
color 0D
echo.
echo  ========================================
echo   Fiuxxed's AutoTyper — Update Tool
echo  ========================================
echo.

REM Check for internet connection
ping -n 1 raw.githubusercontent.com >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] No internet connection detected.
    echo  Please connect to the internet and try again.
    pause
    exit /b 1
)

REM Read current version if it exists
set CURRENT_VER=none
if exist version.txt set /p CURRENT_VER=<version.txt

REM Fetch latest version number
echo  Checking for updates...
curl -s -o version_new.txt "https://raw.githubusercontent.com/fiuxxed/fiuxxed-autotyper/main/version.txt" >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Could not reach update server.
    if exist version_new.txt del version_new.txt
    pause
    exit /b 1
)

set /p NEW_VER=<version_new.txt
del version_new.txt

echo  Current version : %CURRENT_VER%
echo  Latest version  : %NEW_VER%
echo.

if "%CURRENT_VER%"=="%NEW_VER%" (
    echo  You already have the latest version ^(%NEW_VER%^)!
    echo  No update needed.
    echo.
    pause
    exit /b 0
)

echo  Update available: %CURRENT_VER% ^-^> %NEW_VER%
echo.
set /p CONFIRM= Install update? (Y/N): 
if /i not "%CONFIRM%"=="Y" (
    echo  Update cancelled.
    pause
    exit /b 0
)

echo.
echo  Downloading updates...
echo.

REM Create web folder if it doesn't exist
if not exist web mkdir web
if not exist web\sounds mkdir web\sounds

REM Download main files
echo  [1/6] Updating main.py...
curl -s -o main.py "https://raw.githubusercontent.com/fiuxxed/fiuxxed-autotyper/main/main.py"
if errorlevel 1 ( echo  [FAILED] main.py ) else ( echo  [OK] main.py )

echo  [2/6] Updating web/index.html...
curl -s -o web\index.html "https://raw.githubusercontent.com/fiuxxed/fiuxxed-autotyper/main/web/index.html"
if errorlevel 1 ( echo  [FAILED] web/index.html ) else ( echo  [OK] web/index.html )

echo  [3/6] Updating requirements.txt...
curl -s -o requirements.txt "https://raw.githubusercontent.com/fiuxxed/fiuxxed-autotyper/main/requirements.txt"
if errorlevel 1 ( echo  [FAILED] requirements.txt ) else ( echo  [OK] requirements.txt )

REM Download sound files (only if they don't exist or server has newer versions)
echo  [4/6] Checking sounds...
curl -s -o web\sounds\info.mp3 "https://raw.githubusercontent.com/fiuxxed/fiuxxed-autotyper/main/sounds/info.mp3" 2>nul
curl -s -o web\sounds\warning.mp3 "https://raw.githubusercontent.com/fiuxxed/fiuxxed-autotyper/main/sounds/warning.mp3" 2>nul
curl -s -o web\sounds\alert.mp3 "https://raw.githubusercontent.com/fiuxxed/fiuxxed-autotyper/main/sounds/alert.mp3" 2>nul
curl -s -o web\sounds\cheer.mp3 "https://raw.githubusercontent.com/fiuxxed/fiuxxed-autotyper/main/sounds/cheer.mp3" 2>nul
echo  [OK] Sounds

echo  [5/6] Updating install.bat...
curl -s -o install.bat "https://raw.githubusercontent.com/fiuxxed/fiuxxed-autotyper/main/install.bat" 2>nul
echo  [OK] install.bat

echo  [6/6] Saving version...
echo %NEW_VER%> version.txt
echo  [OK] Version saved as %NEW_VER%

echo.
echo  ========================================
echo   Update complete! Now on v%NEW_VER%
echo  ========================================
echo.
echo  Restart the app to apply the update.
echo.
pause
