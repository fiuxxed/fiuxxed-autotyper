═══════════════════════════════════════════════════════
  Fiuxxed's AutoTyper v9.0  —  README
═══════════════════════════════════════════════════════

QUICK START (3 steps)
─────────────────────
1. Double-click  install.bat   (first time only)
2. Double-click  run.bat       (launches the app)
3. Add your Groq API key:
     → Go to https://console.groq.com  (free, no expiry)
     → Sign up, create an API key, copy it
     → In the app: Settings ⚙ → AI & App tab → paste key → 💾 Save Key

FEATURES
────────
⌨  AutoType  — Types text humanly (WPM, stutter, typos, thinking pauses)
🔍  Scanner  — Screenshots any window, AI answers ALL questions
               Detects answered questions, flags wrong ones, shows correction boxes
📐  Math     — Per-question tabs: Explanation + Vertical Method + Graph for EACH problem
💬  Q&A      — Ask anything, follow up infinitely with full context
🎤  Voice    — Continuous Hey-Siri style wake word detection
📋  History  — All past scans & math results saved automatically

HOTKEYS (customizable in ⚙ Settings → Typing tab)
──────────────────────────────────────────────────
F8  = Start AutoType
F9  = Pause / Resume
F10 = Stop

VOICE COMMANDS
──────────────
"hey fiuxxed start"   → trigger AutoType
"hey fiuxxed pause"   → pause/resume
"hey fiuxxed stop"    → stop typing
(Change wake word in Settings → Voice tab)

SCANNER — NEW IN V9
─────────────────────
- Color-coded question borders: Green=correct, Red=wrong, Purple=unanswered
- Correction boxes appear for wrong answers showing exact format fix
- Detects numbered/lettered questions (1., 2., Q1, a), b)) reliably
- "Recheck" button re-evaluates any answer with higher accuracy

MATH SOLVER — NEW IN V9
─────────────────────────
- Click any problem card to expand it
- Three tabs per problem: Explanation | Vertical Method | Graph
- Every plottable function gets its own graph
- Solutions auto-saved to Formula Library (Settings → Scanner tab to toggle)

SETTINGS TABS
─────────────
⌨ Typing   — WPM, stutter, typos, hotkeys, timing
🔍 Scanner — Wrong answer strictness, highlight mode, math graphs, formula library
🎤 Voice   — Sensitivity, language, wake word
🤖 AI & App — Groq API key (with save button), always on top, opacity

FILES
─────
main.py              Python backend
web/index.html       Frontend UI
settings.json        Your settings (auto-created)
formula_library.json Saved math solutions (auto-created)
scan_history.json    Past scan results (auto-created)
install.bat          Installs all Python packages
run.bat              Launches the app

═══════════════════════════════════════════════════════
