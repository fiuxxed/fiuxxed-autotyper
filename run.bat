@echo off
cd /d "%~dp0"

set PYTHON=

py --version >nul 2>&1
if not errorlevel 1 (set PYTHON=py)

if "%PYTHON%"=="" (
    python --version >nul 2>&1
    if not errorlevel 1 (set PYTHON=python)
)

if "%PYTHON%"=="" (
    python3 --version >nul 2>&1
    if not errorlevel 1 (set PYTHON=python3)
)

if "%PYTHON%"=="" (
    echo Python not found. Run install.bat first.
    pause
    exit /b 1
)

%PYTHON% -c "import flask" >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    call install.bat
)

%PYTHON% "%~dp0launch.pyw"
