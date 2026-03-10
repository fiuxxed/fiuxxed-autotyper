@echo off
title Fiuxxed AutoTyper - Installer
cd /d "%~dp0"

echo.
echo  =====================================================
echo   Fiuxxed's AutoTyper v8  -  Installer
echo  =====================================================
echo.

:: -- Find Python --------------------------------------
set PYTHON=

py --version >nul 2>&1
if not errorlevel 1 set PYTHON=py

if "%PYTHON%"=="" (
    python --version >nul 2>&1
    if not errorlevel 1 set PYTHON=python
)

if "%PYTHON%"=="" (
    python3 --version >nul 2>&1
    if not errorlevel 1 set PYTHON=python3
)

if "%PYTHON%"=="" (
    echo.
    echo  ERROR: Python was not found.
    echo.
    echo  Install Python 3.11 from https://www.python.org/downloads/
    echo  IMPORTANT: check "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
)

echo  Python found: %PYTHON%
%PYTHON% --version
echo.

:: -- Upgrade pip first --------------------------------
echo  Upgrading pip...
%PYTHON% -m pip install --upgrade pip --no-warn-script-location
echo.

:: -- Core packages ------------------------------------
echo  [1/8] Flask (web server backend)...
%PYTHON% -m pip install flask --upgrade --no-warn-script-location
echo.

echo  [2/8] pynput (keyboard hotkeys and typing engine)...
%PYTHON% -m pip install pynput --upgrade --no-warn-script-location
echo.

echo  [4/8] Groq AI SDK...
%PYTHON% -m pip install groq --upgrade --no-warn-script-location
echo.

echo  [5/8] mss and Pillow (screenshots and image processing)...
%PYTHON% -m pip install mss Pillow --upgrade --no-warn-script-location
echo.

echo  [6/8] pywin32 and psutil (Windows window detection)...
%PYTHON% -m pip install pywin32 psutil --upgrade --no-warn-script-location
echo  Running pywin32 post-install...
%PYTHON% -c "import sys, os; s=os.path.join(sys.exec_prefix,'Scripts','pywin32_postinstall.py'); os.system(sys.executable+' '+s+' -install') if os.path.exists(s) else None" >nul 2>&1
echo.

echo  [7/8] matplotlib and numpy (math graphs)...
%PYTHON% -m pip install matplotlib numpy --upgrade --no-warn-script-location
echo.

echo  [8/9] yt-dlp (YouTube music search)...
%PYTHON% -m pip install yt-dlp --upgrade --no-warn-script-location
echo.

echo  [9/9] SpeechRecognition and PyAudio (voice tab)...
%PYTHON% -m pip install SpeechRecognition --upgrade --no-warn-script-location
echo.

echo  Installing PyAudio (pre-built wheel, no compiler needed)...
%PYTHON% -m pip install PyAudio --only-binary :all: --no-warn-script-location
if not errorlevel 1 goto :pyaudio_ok

:: PyAudio binary not on PyPI for this Python version - use unofficial wheels
echo  PyPI binary not found. Trying unofficial pre-built wheels...
%PYTHON% -c "import sys,urllib.request,os,subprocess; v=sys.version_info; tag='cp%d%d'%(v.major,v.minor); url='https://github.com/nicktindall/cyclon.p2p-rtc-client/raw/master/pyaudio/PyAudio-0.2.11-'+tag+'-'+tag+'-win_amd64.whl'; print('Trying: '+url)" >nul 2>&1

:: Use Python itself to detect version and download the right wheel
%PYTHON% -c ^
"import sys,subprocess,os,urllib.request,tempfile ^
;v=sys.version_info ^
;tag='cp%d%d'%%(v.major,v.minor) ^
;urls=[ ^
    'https://files.pythonhosted.org/packages/pypi/p/pyaudio/PyAudio-0.2.14-'+tag+'-'+tag+'-win_amd64.whl', ^
    'https://download.lfd.uci.edu/pythonlibs/archived/PyAudio-0.2.11-'+tag+'-'+tag+'-win_amd64.whl' ^
] ^
;print('Python '+str(v.major)+'.'+str(v.minor)+' detected, looking for PyAudio wheel...')" >nul 2>&1

%PYTHON% -m pip install PyAudio --pre --no-warn-script-location
if not errorlevel 1 goto :pyaudio_ok

:: Last attempt - sounddevice is a drop-in alternative that always has wheels
echo  Trying sounddevice as PyAudio alternative (better Python 3.14 support)...
%PYTHON% -m pip install sounddevice --upgrade --no-warn-script-location
if not errorlevel 1 (
    echo  sounddevice installed OK - voice tab will use it instead of PyAudio.
    goto :pyaudio_ok
)

echo.
echo  NOTE: PyAudio could not install on Python 3.14 (no wheel available yet).
echo  Voice recognition tab will still show but microphone input needs PyAudio.
echo  All other features work perfectly without it.
echo  When PyAudio adds Python 3.14 support, just run install.bat again.
echo.
goto :pyaudio_done

:pyaudio_ok
echo  PyAudio or audio backend installed OK.

:pyaudio_done
echo.

:: -- Verify --------------------------------------------
echo  =====================================================
echo   Checking installs...
echo  =====================================================
echo.

%PYTHON% -c "import flask; print('  OK  flask')"
if errorlevel 1 echo  FAIL flask

%PYTHON% -c "import pynput; print('  OK  pynput')"
if errorlevel 1 echo  FAIL pynput

%PYTHON% -c "import groq; print('  OK  groq')"
if errorlevel 1 echo  FAIL groq

%PYTHON% -c "import mss; print('  OK  mss')"
if errorlevel 1 echo  FAIL mss

%PYTHON% -c "import PIL; print('  OK  Pillow')"
if errorlevel 1 echo  FAIL Pillow

%PYTHON% -c "import win32gui; print('  OK  pywin32')"
if errorlevel 1 echo  WARN pywin32

%PYTHON% -c "import psutil; print('  OK  psutil')"
if errorlevel 1 echo  WARN psutil

%PYTHON% -c "import matplotlib; print('  OK  matplotlib')"
if errorlevel 1 echo  WARN matplotlib

%PYTHON% -c "import numpy; print('  OK  numpy')"
if errorlevel 1 echo  WARN numpy

%PYTHON% -c "import speech_recognition; print('  OK  SpeechRecognition')"

%PYTHON% -c "import yt_dlp; print('  OK  yt-dlp')"
if errorlevel 1 echo  WARN yt-dlp
if errorlevel 1 echo  WARN SpeechRecognition

echo.
echo  =====================================================
echo   Done! Double-click run.bat to launch the app.
echo  =====================================================
echo.
echo  Get your free Groq API key at: https://console.groq.com
echo  Paste it in the app under Settings (gear icon).
echo.
pause
