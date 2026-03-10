@echo off
cd /d "%~dp0"

set PYTHON=

py --version >nul 2>&1
if not errorlevel 1 (set PYTHON=py) & goto :found_python

python --version >nul 2>&1
if not errorlevel 1 (set PYTHON=python) & goto :found_python

python3 --version >nul 2>&1
if not errorlevel 1 (set PYTHON=python3) & goto :found_python

echo Python not found. Run install.bat first.
pause
exit /b 1

:found_python

:: Check deps — if flask missing, run installer first
%PYTHON% -c "import flask" >nul 2>&1
if errorlevel 1 (
    echo First time setup — running installer...
    call install.bat
)

:: Check pywin32 is properly registered — critical for titlebar + always-on-top
%PYTHON% -c "import win32gui" >nul 2>&1
if errorlevel 1 (
    echo Registering pywin32 DLLs...
    %PYTHON% -c "import sys,os,subprocess; s=os.path.join(sys.exec_prefix,'Scripts','pywin32_postinstall.py'); subprocess.call([sys.executable,s,'-install']) if os.path.exists(s) else None" >nul 2>&1
)

%PYTHON% "%~dp0launch.pyw"
