@echo off
title Fiuxxed's AutoTyper - Install / Update Dependencies
color 0D
echo.
echo  ========================================
echo   Fiuxxed's AutoTyper - Installer
echo  ========================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    python3 --version >nul 2>&1
    if errorlevel 1 (
        echo  [ERROR] Python not found. Install Python 3.10+ from python.org
        pause
        exit /b 1
    )
    set PYTHON=python3
) else (
    set PYTHON=python
)

echo  Python found. Installing packages...
echo.

REM Core
%PYTHON% -m pip install --upgrade pip --quiet
%PYTHON% -m pip install flask --quiet
%PYTHON% -m pip install groq --quiet

REM NEW: Gemini 2.0 Flash — primary image scanner
echo  Installing Google Generative AI (Gemini 2.0 Flash)...
%PYTHON% -m pip install google-generativeai --quiet

REM Vision / screenshot
%PYTHON% -m pip install mss Pillow --quiet

REM Windows API
%PYTHON% -m pip install pywin32 psutil --quiet

REM Voice
%PYTHON% -m pip install SpeechRecognition sounddevice --quiet

REM Hotkeys
%PYTHON% -m pip install pynput --quiet

REM OCR (optional — improves math scanning, needs Tesseract installed separately)
%PYTHON% -m pip install pytesseract --quiet

REM YouTube transcripts
%PYTHON% -m pip install youtube-transcript-api --quiet

echo.
echo  ========================================
echo   All done! You can now run the app.
echo  ========================================
echo.
echo  IMPORTANT: Add your API keys in Settings after launch:
echo   - Groq key:   console.groq.com  (free)
echo   - Gemini key: aistudio.google.com  (free, 1500 req/day)
echo.
pause
