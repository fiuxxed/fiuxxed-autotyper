A Windows desktop app that auto-types text into any field with human-like rhythm, typos, and pacing. Built with Python + Flask backend and a custom frameless UI running in Edge app-mode.

<img width="597" height="570" alt="image" src="https://github.com/user-attachments/assets/03e9b241-aade-4b41-861d-91efc150c5ae" />


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

SETUP: Run install.bat once, then run.bat to launch. Requires a free Groq API key.
