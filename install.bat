@echo off
title Fiuxxed's AutoTyper v9 - Installer
cd /d "%~dp0"
color 0A

echo.
echo  =====================================================
echo   Fiuxxed's AutoTyper v9  --  AI Edition
echo   Installer
echo  =====================================================
echo.

:: -------------------------------------------------------
:: Find Python
:: -------------------------------------------------------
set PYTHON=

py --version >nul 2>&1
if not errorlevel 1 (set PYTHON=py) & goto :found_python

python --version >nul 2>&1
if not errorlevel 1 (set PYTHON=python) & goto :found_python

python3 --version >nul 2>&1
if not errorlevel 1 (set PYTHON=python3) & goto :found_python

echo.
echo  [ERROR] Python was not found on this machine.
echo.
echo  Download Python 3.11 or newer from:
echo    https://www.python.org/downloads/
echo.
echo  IMPORTANT: Check "Add Python to PATH" during install.
echo.
pause
exit /b 1

:found_python
echo  Python found:
%PYTHON% --version
echo.

:: -------------------------------------------------------
:: Upgrade pip silently
:: -------------------------------------------------------
echo  Upgrading pip...
%PYTHON% -m pip install --upgrade pip --quiet --no-warn-script-location
echo  Done.
echo.

:: -------------------------------------------------------
:: Core packages
:: -------------------------------------------------------
echo  [1/9] Flask  (web server)...
%PYTHON% -m pip install flask --upgrade --quiet --no-warn-script-location
echo  Done.
echo.

echo  [2/9] pynput  (hotkeys + typing engine)...
%PYTHON% -m pip install pynput --upgrade --quiet --no-warn-script-location
echo  Done.
echo.

echo  [3/9] Groq AI SDK...
%PYTHON% -m pip install groq --upgrade --quiet --no-warn-script-location
echo  Done.
echo.

echo  [4/9] mss + Pillow  (screenshots)...
%PYTHON% -m pip install mss Pillow --upgrade --quiet --no-warn-script-location
echo  Done.
echo.

echo  [5/9] pywin32 + psutil  (window control - NEEDED for titlebar + always-on-top)...
%PYTHON% -m pip install pywin32 psutil --upgrade --quiet --no-warn-script-location
echo  Registering pywin32 DLLs (required for win32gui to work)...
%PYTHON% -c "import sys, os, subprocess; script = os.path.join(sys.exec_prefix, 'Scripts', 'pywin32_postinstall.py'); subprocess.call([sys.executable, script, '-install']) if os.path.exists(script) else None"
echo  Done.
echo.

echo  [6/10] pywebview  (app window fallback)...
%PYTHON% -m pip install pywebview --upgrade --quiet --no-warn-script-location
echo  Done.
echo.

echo  [6b] tkinterweb  (embedded browser in tkinter)...
%PYTHON% -m pip install tkinterweb --upgrade --quiet --no-warn-script-location
echo  Done.
echo.

echo  [7/10] matplotlib + numpy  (math graphs)...
%PYTHON% -m pip install matplotlib numpy --upgrade --quiet --no-warn-script-location
echo  Done.
echo.

echo  [8/10] yt-dlp  (YouTube music search)...
%PYTHON% -m pip install yt-dlp --upgrade --quiet --no-warn-script-location
echo  Done.
echo.

echo  [9/10] SpeechRecognition  (voice tab)...
%PYTHON% -m pip install SpeechRecognition --upgrade --quiet --no-warn-script-location
echo  Done.
echo.

echo  [10/10] PyAudio  (microphone input)...
%PYTHON% -m pip install PyAudio --only-binary :all: --quiet --no-warn-script-location
if not errorlevel 1 goto :audio_ok

echo  PyAudio wheel not available, trying sounddevice...
%PYTHON% -m pip install sounddevice --upgrade --quiet --no-warn-script-location
if not errorlevel 1 goto :audio_ok

echo  NOTE: No audio backend installed. Voice tab needs PyAudio or sounddevice.
goto :audio_done

:audio_ok
echo  Done.

:audio_done
echo.

:: -------------------------------------------------------
:: Verify
:: -------------------------------------------------------
echo  =====================================================
echo   Verifying installs...
echo  =====================================================
echo.

%PYTHON% -c "import flask; print('  [OK]  flask')" 2>nul
if errorlevel 1 echo  [FAIL] flask - run install.bat again

%PYTHON% -c "import pynput; print('  [OK]  pynput')" 2>nul
if errorlevel 1 echo  [FAIL] pynput - run install.bat again

%PYTHON% -c "import groq; print('  [OK]  groq')" 2>nul
if errorlevel 1 echo  [FAIL] groq - run install.bat again

%PYTHON% -c "import mss; print('  [OK]  mss')" 2>nul
if errorlevel 1 echo  [FAIL] mss - run install.bat again

%PYTHON% -c "import PIL; print('  [OK]  Pillow')" 2>nul
if errorlevel 1 echo  [FAIL] Pillow - run install.bat again

%PYTHON% -c "import win32gui; print('  [OK]  pywin32  ^(titlebar + always-on-top active^)')" 2>nul
if errorlevel 1 echo  [FAIL] pywin32 - titlebar hiding + always-on-top will NOT work - rerun install.bat

%PYTHON% -c "import psutil; print('  [OK]  psutil')" 2>nul
if errorlevel 1 echo  [WARN] psutil missing

%PYTHON% -c "import matplotlib; print('  [OK]  matplotlib')" 2>nul
if errorlevel 1 echo  [WARN] matplotlib missing - math graphs disabled

%PYTHON% -c "import numpy; print('  [OK]  numpy')" 2>nul
if errorlevel 1 echo  [WARN] numpy missing - math graphs disabled

%PYTHON% -c "import yt_dlp; print('  [OK]  yt-dlp')" 2>nul
if errorlevel 1 echo  [WARN] yt-dlp missing - YouTube search less reliable

%PYTHON% -c "import webview; print('  [OK]  pywebview  (AOT + frameless window)')" 2>nul
if errorlevel 1 echo  [FAIL] pywebview - always-on-top and frameless will NOT work

%PYTHON% -c "import speech_recognition; print('  [OK]  SpeechRecognition')" 2>nul
if errorlevel 1 echo  [WARN] SpeechRecognition missing - voice tab disabled

%PYTHON% -c "import pyaudio; print('  [OK]  PyAudio')" 2>nul
if errorlevel 1 (
    %PYTHON% -c "import sounddevice; print('  [OK]  sounddevice')" 2>nul
    if errorlevel 1 echo  [WARN] no audio backend - mic input disabled
)

echo.
echo  =====================================================
echo   Done! Run run.bat to launch.
echo  =====================================================
echo.
echo  Get your free Groq API key at: https://console.groq.com
echo  Paste it in the app under Settings ^(gear icon^).
echo.
pause
