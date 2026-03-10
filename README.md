A Windows desktop app that auto-types text into any field with human-like rhythm, typos, and pacing. Built with Python + Flask backend and a custom frameless UI running in Edge app-mode.

<img width="604" height="933" alt="image" src="https://github.com/user-attachments/assets/59beae53-41be-41e8-a69e-85e75095780f" />

Features:

⌨️ AutoType — types any text with configurable WPM, rhythm variance, stutters, typos, thinking pauses, and hotkey triggers
🔍 AI Scanner — screenshot any window, AI detects every question, checks answers, flags wrong ones with corrections
📐 AI Math Solver — scans math problems, returns step-by-step solutions with explanations, vertical method, and graphs
💬 Q&A — ask the AI anything, one-shot or follow-up conversation mode
🎤 Voice Control — wake-word activated commands to start/pause/stop typing hands-free, with live mic monitor
🎵 Music Player — search YouTube or play local files (mp3/flac/wav/etc) without leaving the app
📋 History — auto-saves all past scans and math solutions, formula library for reuse
🪟 Custom UI — fully frameless window with custom purple titlebar, drag, minimize, and close — no default Windows chrome

Stack: Python · Flask · Groq API (Llama 4 Scout + Llama 3.3) · yt-dlp · pywin32 · pynput · Web Audio API
Setup: Run install.bat once, then run.bat to launch. Requires a free Groq API key.
