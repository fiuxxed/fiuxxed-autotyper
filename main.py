# ══════════════════════════════════════════════════════════════════════
#  Fiuxxed's AutoTyper v9.1  —  main.py
# ══════════════════════════════════════════════════════════════════════
import threading, time, random, json, os, sys, re, io, base64, traceback, ctypes

BASE      = os.path.dirname(os.path.abspath(__file__))
SAVE_FILE = os.path.join(BASE, "settings.json")
WEB_DIR   = os.path.join(BASE, "web")

# ── Admin system ─────────────────────────────────────────────────────
# IP is written to a hidden file in the user's home folder on startup.
# Nobody else has this file. Admin routes check both the file AND password.
_ADMIN_TOKEN_FILE = os.path.join(os.path.expanduser("~"), ".fiuxxed_admin")
_ADMIN_PASSWORD   = "fiuxxedADMIN783"
_ADMIN_SESSION    = set()

# Supabase constants for server-side operations
_SUPA_URL         = "https://gmuiprcobibatlayyemz.supabase.co"
_SUPA_KEY         = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImdtdWlwcmNvYmliYXRsYXl5ZW16Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzM2OTMwNDgsImV4cCI6MjA4OTI2OTA0OH0._eD7ZflslMY0cHdAyzM1bVwiUaiilEuoKD042k_3ibI"
_SUPA_SERVICE_KEY = _SUPA_KEY

def _write_admin_ip():
    """Called once at startup — writes current machine's loopback marker."""
    try:
        import socket
        # Get the local LAN ip too in case Flask sees that instead of 127.0.0.1
        local_ip = socket.gethostbyname(socket.gethostname())
        with open(_ADMIN_TOKEN_FILE, "w") as f:
            f.write(json.dumps({"ips": ["127.0.0.1", "::1", "::ffff:127.0.0.1", local_ip]}))
        _log(f"Admin IP file written to {_ADMIN_TOKEN_FILE}")
    except Exception as e:
        _log(f"Admin IP write failed: {e}")

def _is_admin_ip(request_obj):
    """Returns True only if the request comes from the admin's machine."""
    try:
        with open(_ADMIN_TOKEN_FILE) as f:
            data = json.load(f)
        allowed = data.get("ips", [])
        remote = request_obj.remote_addr or ""
        return remote in allowed
    except Exception:
        return False

def _is_admin_session(request_obj):
    """Returns True if request has a valid admin session token."""
    token = request_obj.headers.get("X-Admin-Token","") or request_obj.json.get("admin_token","") if request_obj.is_json else request_obj.headers.get("X-Admin-Token","")
    return token in _ADMIN_SESSION

def _require_admin(request_obj):
    """Returns (ok, error_response). Call at top of every admin route."""
    if not _is_admin_ip(request_obj):
        return False, jsonify({"error":"forbidden"})
    return True, None

# ── Global error log ─────────────────────────────────────────────────
LOG_FILE = os.path.join(BASE, "error.log")

def _log(msg):
    """Write msg to terminal and error.log with a timestamp."""
    line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def _log_exc(label, exc=None):
    """Log a labelled exception with full traceback."""
    _log(f"ERROR — {label}: {exc if exc else ''}")
    tb = traceback.format_exc()
    if tb.strip() != "NoneType: None":
        for line in tb.splitlines():
            _log("  " + line)

# Catch any unhandled exception on any thread
def _unhandled_thread_exc(args):
    _log_exc(f"Unhandled exception in thread '{args.thread.name}'", args.exc_value)

threading.excepthook = _unhandled_thread_exc

# Catch any unhandled exception on the main thread
def _unhandled_exc(exc_type, exc_value, exc_tb):
    _log(f"FATAL unhandled exception: {exc_type.__name__}: {exc_value}")
    for line in traceback.format_tb(exc_tb):
        _log("  " + line.rstrip())
    sys.__excepthook__(exc_type, exc_value, exc_tb)

sys.excepthook = _unhandled_exc

_log("AutoTyper starting up")
_write_admin_ip()

DEFAULTS = {
    "text": "", "wpm": 80, "countdown": 5,
    "repeat_count": 1, "repeat_delay": 2.0,
    "stop_after_chars": 0, "stop_after_words": 0, "stop_after_lines": 0,
    "typo_chance": 0, "stutter_chance": 40, "stutter_duration": 2.0,
    "thinking_pause_chance": 3, "thinking_pause_min": 300, "thinking_pause_max": 800,
    "synonym_swap_chance": 15, "synonym_swap_pause": 1.2,
    "symbol_pause_min": 2.0, "symbol_pause_max": 6.0,
    "punct_delay_mult": 2.2, "newline_delay_mult": 3.0, "rhythm_variance": 35,
    "hotkey_start": "F8", "hotkey_pause": "F9", "hotkey_stop": "F10",
    "always_on_top": True, "opacity": 100,
    "start_delay": 0.0, "end_delay": 0.0,
    "char_blacklist": "", "line_by_line": False, "line_pause": 1.0,
    "groq_api_key": "",
    "gemini_api_key": "",
    "voice_wake_word": "hey fiuxxed",
    "voice_sensitivity": 50,
    "voice_language": "en-US",
    "scanner_confidence_threshold": 70,
    "scanner_auto_retype": True,
    "scanner_highlight_mode": True,
    "scanner_wrong_answer_strictness": "flag_all",
    "examine_examples": False,  # if True, AI will include worked examples in explanations
    "math_show_graphs": True,
    "math_formula_library": True,
}

def load_cfg():
    try:
        with open(SAVE_FILE) as f: d = json.load(f)
        for k, v in DEFAULTS.items(): d.setdefault(k, v)
        return d
    except Exception: return dict(DEFAULTS)

def save_cfg(d):
    try:
        with open(SAVE_FILE, "w") as f: json.dump(d, f, indent=2)
    except Exception: pass

try:
    from flask import Flask, request, jsonify, send_from_directory
    HAS_FLASK = True
except ImportError:
    print("Flask not found. Run install.bat first."); sys.exit(1)

try:
    import webview
    HAS_WEBVIEW = True
except ImportError: HAS_WEBVIEW = False

try:
    import win32gui, win32process
    import psutil
    HAS_WIN32 = True
except ImportError: HAS_WIN32 = False

try:
    import mss
    from PIL import Image, ImageEnhance
    HAS_MSS = True
except ImportError: HAS_MSS = False

try:
    from groq import Groq
    HAS_GROQ = True
except ImportError: HAS_GROQ = False

try:
    import google.generativeai as genai
    HAS_GEMINI = True
except ImportError: HAS_GEMINI = False

try:
    from pynput import keyboard as kb
    HAS_PYNPUT = True
except ImportError: HAS_PYNPUT = False

HAS_VOICE = False; HAS_PYAUDIO = False; HAS_SOUNDDEVICE = False; HAS_SR = False

try:
    import speech_recognition as sr
    HAS_SR = True
except ImportError: pass

try:
    import pyaudio as _pyaudio_mod
    HAS_PYAUDIO = True; HAS_VOICE = True
except ImportError: pass

try:
    import sounddevice as sd
    import numpy as np_sd
    HAS_SOUNDDEVICE = True
    if not HAS_PYAUDIO: HAS_VOICE = True
except ImportError: pass

if not HAS_PYAUDIO and not HAS_SOUNDDEVICE: HAS_VOICE = False

try:
    import matplotlib; matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np
    import numpy as _np
    HAS_PLOT = True
except ImportError: HAS_PLOT = False

try:
    import fitz as _fitz   # PyMuPDF — pip install pymupdf
    HAS_FITZ = True
except ImportError: HAS_FITZ = False

BROWSERS = {"chrome","firefox","msedge","opera","brave","vivaldi","chromium","iexplore","waterfox","librewolf"}

# ── Formula Library ──────────────────────────────────────────────────
FORMULA_LIB_FILE = os.path.join(BASE, "formula_library.json")

def load_formula_lib():
    try:
        with open(FORMULA_LIB_FILE) as f: return json.load(f)
    except Exception: return []

def save_formula_lib(lib):
    try:
        with open(FORMULA_LIB_FILE, "w") as f: json.dump(lib, f, indent=2)
    except Exception: pass

def add_to_formula_lib(problem, answer, steps):
    lib = load_formula_lib()
    for e in lib:
        if e.get("problem","").strip().lower() == problem.strip().lower(): return
    lib.append({"problem": problem, "answer": answer, "steps": steps, "saved_at": time.strftime("%Y-%m-%d %H:%M")})
    if len(lib) > 200: lib = lib[-200:]
    save_formula_lib(lib)

# ── Scan History ─────────────────────────────────────────────────────
HISTORY_FILE = os.path.join(BASE, "scan_history.json")

def load_history():
    try:
        with open(HISTORY_FILE) as f: return json.load(f)
    except Exception: return []

def save_history_file(h):
    try:
        with open(HISTORY_FILE, "w") as f: json.dump(h, f, indent=2)
    except Exception: pass

def add_to_history(mode, result):
    h = load_history()
    h.append({"mode": mode, "result": result, "time": time.strftime("%Y-%m-%d %H:%M:%S")})
    if len(h) > 100: h = h[-100:]
    save_history_file(h)

# ══════════════════════════════════════════════════════════════════════
#  SYNONYM LOOKUP  — simple built-in map, no external deps
# ══════════════════════════════════════════════════════════════════════
_SYNONYMS = {
    "good":["great","solid","nice","fine"],"bad":["poor","weak","rough","lousy"],
    "big":["large","huge","major","wide"],"small":["tiny","little","minor","slim"],
    "fast":["quick","rapid","swift","speedy"],"slow":["gradual","steady","gentle","lazy"],
    "hard":["tough","firm","rigid","rough"],"easy":["simple","basic","smooth","light"],
    "smart":["bright","sharp","clever","wise"],"dumb":["slow","dense","dim","thick"],
    "happy":["glad","pleased","content","joyful"],"sad":["upset","down","low","blue"],
    "angry":["mad","upset","furious","cross"],"scared":["afraid","nervous","uneasy","tense"],
    "beautiful":["pretty","lovely","nice","gorgeous"],"ugly":["rough","plain","harsh","gross"],
    "important":["key","major","serious","big"],"boring":["dull","flat","dry","bland"],
    "interesting":["cool","neat","fun","wild"],"funny":["silly","goofy","wild","odd"],
    "old":["aged","dated","worn","prior"],"new":["fresh","recent","latest","modern"],
    "many":["lots","several","various","plenty"],"few":["some","a couple","barely any","limited"],
    "very":["really","super","quite","pretty"],"also":["too","as well","plus","and"],
    "however":["but","though","yet","still"],"therefore":["so","thus","hence","then"],
    "because":["since","as","given that","seeing as"],"although":["even though","while","though","despite"],
    "usually":["often","mostly","generally","typically"],"sometimes":["at times","now and then","occasionally","here and there"],
    "always":["every time","all the time","constantly","forever"],"never":["not once","at no point","not ever","zero times"],
    "shows":["proves","tells","reveals","makes clear"],"helps":["aids","supports","assists","makes easier"],
    "uses":["applies","works with","relies on","takes advantage of"],"makes":["creates","builds","forms","produces"],
    "gets":["receives","gains","picks up","ends up with"],"gives":["provides","offers","hands","passes"],
    "said":["stated","noted","mentioned","pointed out"],"found":["discovered","noticed","saw","came across"],
    "used":["applied","worked with","employed","took"],"changed":["shifted","moved","switched","altered"],
    "increase":["grow","rise","go up","climb"],"decrease":["drop","fall","go down","shrink"],
    "improve":["get better","boost","strengthen","upgrade"],"affect":["impact","influence","shape","touch"],
    "allow":["let","permit","enable","give room for"],"prevent":["stop","block","keep from","avoid"],
    "require":["need","call for","demand","take"],"include":["have","cover","contain","involve"],
    "develop":["build","grow","create","work on"],"provide":["give","offer","supply","bring"],
    "believe":["think","feel","figure","reckon"],"suggest":["hint","imply","point to","indicate"],
    "explain":["describe","lay out","break down","go over"],"compare":["look at","weigh","contrast","measure against"],
    "consider":["think about","look at","weigh","factor in"],"understand":["get","grasp","follow","see"],
    "result":["outcome","effect","end result","what happens"],"reason":["cause","point","why","factor"],
    "example":["case","instance","sample","like"],"idea":["thought","concept","point","notion"],
    "problem":["issue","trouble","challenge","situation"],"solution":["answer","fix","way out","approach"],
    "people":["folks","others","individuals","everyone"],"thing":["item","part","piece","aspect"],
    "time":["period","moment","point","stretch"],"way":["method","approach","manner","means"],
    "place":["area","spot","location","region"],"group":["set","bunch","collection","cluster"],
    "part":["section","piece","bit","portion"],"point":["detail","factor","aspect","element"],
    "work":["effort","task","job","activity"],"life":["living","existence","daily routine","experience"],
    "society":["community","world","culture","people"],"government":["state","authorities","leadership","officials"],
    "history":["past","background","record","story"],"science":["research","study","field","knowledge"],
    "technology":["tech","tools","systems","advances"],"environment":["surroundings","ecosystem","world","nature"],
    "education":["learning","schooling","training","studies"],"economy":["market","finances","trade","business"],
}

def get_synonym(word):
    """Return a synonym for word if available, else None. Preserves capitalization."""
    key = word.lower().rstrip(".,!?;:")
    syns = _SYNONYMS.get(key)
    if not syns: return None
    syn = random.choice(syns)
    if word[0].isupper(): syn = syn[0].upper() + syn[1:]
    return syn


# ══════════════════════════════════════════════════════════════════════
#  TYPING ENGINE
# ══════════════════════════════════════════════════════════════════════
class TypingEngine:
    def __init__(self):
        self.ctrl   = kb.Controller() if HAS_PYNPUT else None
        self._stop  = threading.Event()
        self._pause = threading.Event()

    def start(self, text, cfg, on_prog, on_done, on_err):
        self._stop.clear(); self._pause.clear()
        threading.Thread(target=self._run, args=(text, cfg, on_prog, on_done, on_err), daemon=True).start()

    def stop(self):   self._stop.set()
    def pause(self):  self._pause.set()
    def resume(self): self._pause.clear()

    def _run(self, text, cfg, on_prog, on_done, on_err):
        try:
            base = 60.0 / (max(cfg["wpm"], 1) * 5)
            bl   = set(cfg.get("char_blacklist", ""))
            lbl  = cfg.get("line_by_line", False)
            lp   = cfg.get("line_pause", 1.0)
            sc   = cfg.get("stop_after_chars", 0)
            sw   = cfg.get("stop_after_words", 0)
            sl_  = cfg.get("stop_after_lines", 0)

            if cfg.get("start_delay", 0) > 0:
                self._isleep(cfg["start_delay"])
                if self._stop.is_set(): return

            lines = text.split("\n"); total = len(text)
            tc = tw = tl = wcount = 0
            inw = False; wchars = []; stutter = False

            for li, line in enumerate(lines):
                if self._stop.is_set(): return
                for i, ch in enumerate(list(line)):
                    if self._stop.is_set(): return
                    while self._pause.is_set():
                        if self._stop.is_set(): return
                        time.sleep(0.05)
                    if ch in bl: tc += 1; continue
                    if sc and tc >= sc: self._stop.set(); return
                    if sw and tw >= sw: self._stop.set(); return
                    if sl_ and tl >= sl_: self._stop.set(); return

                    iswc = ch.strip() != ""
                    if iswc and not inw:
                        inw = True; wchars = []; wcount += 1; tw += 1
                        stutter = (wcount % 4 == 0) and (random.random() < cfg.get("stutter_chance", 40) / 100)
                    elif not iswc and inw:
                        # Word just finished — maybe do synonym swap
                        finished_word = "".join(wchars)
                        sw_chance = cfg.get("synonym_swap_chance", 15) / 100
                        if finished_word and random.random() < sw_chance:
                            syn = get_synonym(finished_word)
                            if syn:
                                # Pause like thinking, then backspace the word, type synonym
                                swap_pause = cfg.get("synonym_swap_pause", 1.2) * (0.85 + random.random() * 0.3)
                                self._isleep(swap_pause)
                                if self._stop.is_set(): return
                                # Backspace original word
                                for _ in range(len(finished_word)):
                                    self._key('', cfg)
                                    time.sleep(base * (0.7 + random.random() * 0.3))
                                if self._stop.is_set(): return
                                # Type synonym
                                for sc2 in syn:
                                    self._key(sc2, cfg)
                                    time.sleep(self._delay(base, sc2, cfg))
                                if self._stop.is_set(): return
                        inw = False; stutter = False
                    if inw: wchars.append(ch)

                    if stutter and inw:
                        ei = i + 1; lc = list(line)
                        while ei < len(lc) and lc[ei].strip() != "": ei += 1
                        wlen = ei - (i - len(wchars) + 1)
                        mid  = int(wlen * (0.4 + random.random() * 0.2))
                        if len(wchars) == mid:
                            stutter = False
                            self._isleep(cfg.get("stutter_duration", 2.0) * (0.9 + random.random() * 0.2))
                            if self._stop.is_set(): return

                    tp = cfg.get("typo_chance", 0) / 100
                    if tp > 0 and ch.isalpha() and random.random() < tp:
                        nk = self._nearby(ch)
                        if nk:
                            self._key(nk, cfg); time.sleep(base * (0.8 + random.random() * 0.4))
                            self._key('\x08', cfg); time.sleep(base * (0.5 + random.random() * 0.3))

                    self._key(ch, cfg); tc += 1
                    if tc % 5 == 0 or tc == total:
                        on_prog(tc, total, int(tc / max(total, 1) * 100))
                    time.sleep(self._delay(base, ch, cfg))

                tl += 1
                if lbl and li < len(lines) - 1:
                    self._isleep(lp)
                    if self._stop.is_set(): return
                if li < len(lines) - 1:
                    if sl_ and tl >= sl_: self._stop.set(); return
                    self._key('\n', cfg)
                    time.sleep(self._delay(base, '\n', cfg))

            if cfg.get("end_delay", 0) > 0: self._isleep(cfg["end_delay"])
            if not self._stop.is_set(): on_done()
        except Exception as e: on_err(str(e))

    def _isleep(self, s):
        end = time.time() + s
        while time.time() < end:
            if self._stop.is_set(): return
            time.sleep(0.05)

    def _key(self, ch, cfg=None):
        if not self.ctrl: return
        is_sym = (cfg and ord(ch) > 127 and ch not in "àáâãäåæçèéêëìíîïðñòóôõöùúûüýþÿÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝÞŸ")
        if is_sym and cfg:
            mn = cfg.get("symbol_pause_min", 2.0); mx = cfg.get("symbol_pause_max", 6.0)
            time.sleep(mn + random.random() * (mx - mn))
        try:
            if   ch == '\x08': self.ctrl.press(kb.Key.backspace); self.ctrl.release(kb.Key.backspace)
            elif ch == '\n':   self.ctrl.press(kb.Key.enter);     self.ctrl.release(kb.Key.enter)
            else:              self.ctrl.type(ch)
        except Exception:
            try: self.ctrl.press(ch); self.ctrl.release(ch)
            except Exception: pass

    def _delay(self, base, ch, cfg):
        rv = cfg.get("rhythm_variance", 35) / 100
        d  = base * (max(0.01, 1 - rv) + random.random() * rv * 2)
        if ch in ".!?,;:": d *= cfg.get("punct_delay_mult", 2.2)
        if ch == " ":       d *= 1.1
        if ch == "\n":      d *= cfg.get("newline_delay_mult", 3.0)
        tp = cfg.get("thinking_pause_chance", 3) / 100
        if random.random() < tp:
            mn = cfg.get("thinking_pause_min", 300) / 1000
            mx = cfg.get("thinking_pause_max", 800) / 1000
            d += mn + random.random() * (mx - mn)
        return max(0.02, d)

    _NB = {
        'a':'sq','b':'vghn','c':'xdfv','d':'serfcx','e':'wrsdf','f':'drtgvc',
        'g':'ftyhbv','h':'gyujnb','i':'uojk','j':'huikm','k':'jiolm','l':'kop',
        'm':'njk','n':'bhjm','o':'ipkl','p':'ol','q':'wa','r':'edft',
        's':'aqwedxz','t':'ryfg','u':'yhji','v':'cfgb','w':'qase',
        'x':'zsdc','y':'tugh','z':'asx'
    }
    def _nearby(self, ch):
        opts = self._NB.get(ch.lower(), "")
        return random.choice(opts) if opts else None


# ══════════════════════════════════════════════════════════════════════
#  VOICE LISTENER  — Continuous Hey Siri style
# ══════════════════════════════════════════════════════════════════════
class VoiceListener:
    def __init__(self, cfg_fn, autotype_fn, pause_fn, stop_fn, device_index=None):
        self.cfg_fn      = cfg_fn
        self.autotype_fn = autotype_fn
        self.pause_fn    = pause_fn
        self.stop_fn     = stop_fn
        self.device_index= device_index
        self._running    = False
        self.log         = []
        self.status      = "idle"

    def start(self):
        if not HAS_VOICE: return
        self._running = True
        self.status = "listening"
        threading.Thread(target=self._loop, daemon=True).start()

    def stop(self):
        self._running = False
        self.status = "idle"

    def _loop(self):
        if HAS_SOUNDDEVICE and not HAS_PYAUDIO:
            self._loop_sd()
        elif HAS_PYAUDIO and HAS_SR:
            self._loop_pa()
        else:
            self._running = False

    def _loop_sd(self):
        import io, wave
        RATE    = 16000
        CHUNK   = int(RATE * 2.5)
        SILENCE = 0.010

        while self._running:
            try:
                frames = sd.rec(CHUNK, samplerate=RATE, channels=1, dtype='int16',
                                device=self.device_index, blocking=True)
                if not self._running: break
                rms = float((frames.astype('float32')**2).mean()**0.5) / 32768
                if rms < SILENCE: continue
                self.status = "processing"
                buf = io.BytesIO()
                with wave.open(buf, 'wb') as wf:
                    wf.setnchannels(1); wf.setsampwidth(2); wf.setframerate(RATE)
                    wf.writeframes(frames.tobytes())
                buf.seek(0)
                rec = sr.Recognizer()
                with sr.AudioFile(buf) as src: audio = rec.record(src)
                try:
                    lang = self.cfg_fn().get("voice_language", "en-US")
                    text = rec.recognize_google(audio, language=lang).lower().strip()
                    self.status = "heard"
                    self._handle(text)
                except sr.UnknownValueError: pass
                self.status = "listening"
            except Exception:
                time.sleep(0.2)
                self.status = "listening"

    def _loop_pa(self):
        rec = sr.Recognizer()
        rec.dynamic_energy_threshold = True
        rec.pause_threshold = 0.6
        rec.non_speaking_duration = 0.4
        try:
            kwargs = {} if self.device_index is None else {"device_index": self.device_index}
            mic = sr.Microphone(**kwargs)
            with mic as source: rec.adjust_for_ambient_noise(source, duration=0.8)
        except Exception:
            self._running = False; return

        while self._running:
            try:
                with mic as source:
                    self.status = "listening"
                    audio = rec.listen(source, timeout=2, phrase_time_limit=6)
                self.status = "processing"
                lang = self.cfg_fn().get("voice_language", "en-US")
                text = rec.recognize_google(audio, language=lang).lower().strip()
                self.status = "heard"
                self._handle(text)
            except sr.WaitTimeoutError: self.status = "listening"
            except sr.UnknownValueError: self.status = "listening"
            except Exception: time.sleep(0.3); self.status = "listening"

    def _handle(self, text):
        wake = self.cfg_fn().get("voice_wake_word", "hey fiuxxed").lower()
        wake_words = wake.split(); text_words = text.split()
        match = False; cmd = ""
        for i in range(len(text_words) - len(wake_words) + 1):
            if text_words[i:i+len(wake_words)] == wake_words:
                match = True; cmd = " ".join(text_words[i+len(wake_words):]).strip(); break
        if not match and wake in text:
            match = True; cmd = text.replace(wake,"").strip()
        if match:
            entry = {"time": time.strftime("%H:%M:%S"), "text": text, "cmd": cmd or "(wake only)"}
            self.log.append(entry)
            if len(self.log) > 50: self.log.pop(0)
            if cmd: self._dispatch(cmd)

    def _dispatch(self, cmd):
        if any(w in cmd for w in ["start","type","autotype","go"]): self.autotype_fn()
        elif any(w in cmd for w in ["pause","resume","hold"]): self.pause_fn()
        elif any(w in cmd for w in ["stop","cancel","quit","end"]): self.stop_fn()


# ══════════════════════════════════════════════════════════════════════
#  SCREENSHOT + IMAGE
# ══════════════════════════════════════════════════════════════════════
def _hide_app_window():
    """Hide the AutoTyper window so it doesn't appear in screenshots."""
    if not HAS_WIN32: return None
    try:
        hwnd = _AOT_HWND or _find_app_hwnd()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE
            time.sleep(0.12)  # give compositor time to repaint
        return hwnd
    except Exception: return None

def _show_app_window(hwnd=None):
    """Restore the AutoTyper window after screenshot."""
    if not HAS_WIN32: return
    try:
        h = hwnd or _AOT_HWND or _find_app_hwnd()
        if h:
            ctypes.windll.user32.ShowWindow(h, 9)   # SW_RESTORE
            ctypes.windll.user32.SetForegroundWindow(h)
    except Exception: pass


def capture_window(hwnd=None, bring_to_front=True):
    if not HAS_MSS: raise RuntimeError("mss/pillow not installed")
    if hwnd and HAS_WIN32:
        if bring_to_front:
            try: win32gui.SetForegroundWindow(hwnd); time.sleep(0.45)
            except Exception: pass
        rect = win32gui.GetWindowRect(hwnd)
        x1, y1, x2, y2 = rect; w = x2-x1; h = y2-y1
        if w < 10 or h < 10: raise RuntimeError("Window too small")

        PW_RENDERFULLCONTENT = 3
        img = None
        try:
            import win32ui, win32con
            hdc_src  = ctypes.windll.user32.GetDC(hwnd)
            hdc_dst  = win32ui.CreateDCFromHandle(hdc_src)
            bmp_dc   = hdc_dst.CreateCompatibleDC()
            bmp      = win32ui.CreateBitmap()
            bmp.CreateCompatibleBitmap(hdc_dst, w, h)
            bmp_dc.SelectObject(bmp)
            result = ctypes.windll.user32.PrintWindow(hwnd, bmp_dc.GetSafeHdc(), PW_RENDERFULLCONTENT)
            bmp_info = bmp.GetInfo()
            bmp_bits = bmp.GetBitmapBits(True)
            bmp_dc.DeleteDC(); hdc_dst.DeleteDC()
            ctypes.windll.user32.ReleaseDC(hwnd, hdc_src)
            win32ui.DeleteObject(bmp.GetHandle())
            if result and any(b != 0 for b in bmp_bits[:512]):
                img = Image.frombuffer("RGB", (bmp_info["bmWidth"], bmp_info["bmHeight"]),
                                       bmp_bits, "raw", "BGRX", 0, 1)
        except Exception:
            pass

        # Fallback: mss screen-grab — hide our window first so it doesn't appear
        if img is None:
            our_hwnd = _hide_app_window()
            try:
                region = {"top": y1, "left": x1, "width": w, "height": h}
                with mss.mss() as sct:
                    raw = sct.grab(region)
                    img = Image.frombytes("RGB", raw.size, raw.bgra, "raw", "BGRX")
            finally:
                _show_app_window(our_hwnd)
    else:
        # Full-screen grab — always hide our window
        our_hwnd = _hide_app_window()
        try:
            with mss.mss() as sct:
                region = sct.monitors[1]
                raw = sct.grab(region)
                img = Image.frombytes("RGB", raw.size, raw.bgra, "raw", "BGRX")
        finally:
            _show_app_window(our_hwnd)

    img = img.convert("RGB")
    if img.width < 1600:
        scale = 1600 / img.width
        img = img.resize((int(img.width * scale), int(img.height * scale)), Image.LANCZOS)
    img = ImageEnhance.Contrast(img).enhance(1.15)
    img = ImageEnhance.Sharpness(img).enhance(1.3)
    return img

def img_to_b64(img):
    buf = io.BytesIO(); img.save(buf, format="PNG", optimize=False)
    return base64.b64encode(buf.getvalue()).decode("utf-8")


# ══════════════════════════════════════════════════════════════════════
#  GRAPH GENERATOR
# ══════════════════════════════════════════════════════════════════════
def make_diagram(diagram_data):
    """Draw a geometry diagram faithfully copying what the AI extracted from the image."""
    if not HAS_PLOT: return None
    try:
        import math as _math
        import re as _re

        def _clean_label(s):
            """Strip LaTeX and convert to plain text for matplotlib."""
            if not s: return ""
            s = str(s)
            # Convert fractions: \frac{1}{8} -> 1/8, 2\frac{1}{8} -> 2 1/8
            s = _re.sub(r'(\d*)\s*\\frac\{(\d+)\}\{(\d+)\}',
                        lambda m: (m.group(1)+" " if m.group(1) else "") + m.group(2)+"/"+m.group(3), s)
            # Convert other LaTeX
            s = s.replace(r'\cong','≅').replace(r'\angle','∠').replace(r'\triangle','△')
            s = s.replace(r'\overline','').replace(r'\circ','°')
            # Remove remaining backslash commands
            s = _re.sub(r'\\[a-zA-Z]+\s*','', s)
            # Remove $, {, }
            s = s.replace('$','').replace('{','').replace('}','').strip()
            return s

        points      = diagram_data.get("points", {})
        edges       = diagram_data.get("edges", [])
        angles      = diagram_data.get("angles", {})
        labels      = {k: _clean_label(v) for k,v in diagram_data.get("labels", {}).items()}
        title       = _clean_label(diagram_data.get("title", "Diagram"))
        tick_marks  = diagram_data.get("tick_marks", {})
        arc_marks   = diagram_data.get("arc_marks", {})
        right_angles= diagram_data.get("right_angles", [])
        par_marks   = diagram_data.get("parallel_marks", [])

        if not points: return None

        fig, ax = plt.subplots(figsize=(6.0, 5.0), facecolor="#07070e")
        ax.set_facecolor("#0d0c18")
        ax.set_aspect("equal")

        xs = [v[0] for v in points.values()]
        ys = [v[1] for v in points.values()]
        cx_all, cy_all = sum(xs)/len(xs), sum(ys)/len(ys)

        def _unit(ax_, ay_, bx, by):
            dx, dy = bx-ax_, by-ay_
            l = max(_math.hypot(dx, dy), 1e-9)
            return dx/l, dy/l

        def _perp(ux, uy): return -uy, ux

        # ── Draw edges ──
        for edge in edges:
            if len(edge) < 2: continue
            a, b = edge[0], edge[1]
            if a not in points or b not in points: continue
            x0,y0 = points[a]; x1,y1 = points[b]
            ax.plot([x0,x1],[y0,y1], color="#a855f7", linewidth=2.0, zorder=3, solid_capstyle="round")

            key1, key2 = a+b, b+a
            lbl = labels.get(key1) or labels.get(key2)
            if lbl:
                mx,my = (x0+x1)/2,(y0+y1)/2
                ux,uy = _unit(x0,y0,x1,y1)
                px,py = _perp(ux,uy)
                ax.text(mx+px*0.35, my+py*0.35, lbl, color="#f59e0b", fontsize=9,
                        ha="center", va="center",
                        bbox=dict(boxstyle="round,pad=0.15", fc="#07070e", ec="none", alpha=0.85))

            n_ticks = tick_marks.get(key1) or tick_marks.get(key2) or 0
            if n_ticks:
                mx,my = (x0+x1)/2,(y0+y1)/2
                ux,uy = _unit(x0,y0,x1,y1)
                px,py = _perp(ux,uy)
                tick_len = 0.18
                spacing = 0.13
                offsets = [0] if n_ticks == 1 else [-spacing/2, spacing/2] if n_ticks == 2 else [-spacing, 0, spacing]
                for off in offsets:
                    tx_ = mx + ux*off; ty_ = my + uy*off
                    ax.plot([tx_-px*tick_len, tx_+px*tick_len],
                            [ty_-py*tick_len, ty_+py*tick_len],
                            color="#10b981", linewidth=2.0, zorder=6)

            par_count = 0
            for pi, pair in enumerate(par_marks):
                if key1 in pair or key2 in pair:
                    par_count = pi + 1; break
            if par_count:
                mx,my = (x0+x1)/2,(y0+y1)/2
                ux,uy = _unit(x0,y0,x1,y1)
                arr_size = 0.15
                for rep in range(par_count):
                    off = (rep - (par_count-1)/2) * 0.2
                    ax_ = mx + ux*off; ay_ = my + uy*off
                    ax.annotate("", xy=(ax_+ux*arr_size, ay_+uy*arr_size),
                                xytext=(ax_-ux*arr_size, ay_-uy*arr_size),
                                arrowprops=dict(arrowstyle="-|>", color="#60a5fa", lw=1.5))

        # ── Arc marks ──
        for pt_name, n_arcs in arc_marks.items():
            if pt_name not in points: continue
            px_,py_ = points[pt_name]
            connected_pts = []
            for edge in edges:
                if pt_name == edge[0] and edge[1] in points: connected_pts.append(edge[1])
                elif pt_name == edge[1] and edge[0] in points: connected_pts.append(edge[0])
            if len(connected_pts) < 2: continue
            a_pt, b_pt = connected_pts[0], connected_pts[1]
            ax_c,ay_c = points[a_pt]; bx,by = points[b_pt]
            ang1 = _math.degrees(_math.atan2(ay_c-py_, ax_c-px_))
            ang2 = _math.degrees(_math.atan2(by-py_, bx-px_))
            if ang1 > ang2: ang1, ang2 = ang2, ang1
            n_arcs = int(n_arcs) if isinstance(n_arcs, (int,float)) else 1
            for arc_i in range(n_arcs):
                r = 0.28 + arc_i * 0.12
                theta = _np.linspace(_math.radians(ang1), _math.radians(ang2), 40)
                ax.plot(px_ + r*_np.cos(theta), py_ + r*_np.sin(theta),
                        color="#f59e0b", linewidth=1.8, zorder=6)

        # ── Right angle boxes ──
        right_angle_pts = list(right_angles or [])
        for pt_name, ang_label in angles.items():
            if "90" in str(ang_label) or "right" in str(ang_label).lower():
                if pt_name not in right_angle_pts: right_angle_pts.append(pt_name)
        for pt_name in right_angle_pts:
            if pt_name not in points: continue
            px_,py_ = points[pt_name]
            connected = []
            for edge in edges:
                if pt_name == edge[0] and edge[1] in points: connected.append(edge[1])
                elif pt_name == edge[1] and edge[0] in points: connected.append(edge[0])
            if len(connected) >= 2:
                s = 0.22
                u1x,u1y = _unit(px_,py_,*points[connected[0]])
                u2x,u2y = _unit(px_,py_,*points[connected[1]])
                sq = plt.Polygon([
                    [px_+s*u1x, py_+s*u1y],
                    [px_+s*u1x+s*u2x, py_+s*u1y+s*u2y],
                    [px_+s*u2x, py_+s*u2y],
                    [px_, py_],
                ], fill=False, edgecolor="#10b981", linewidth=1.5, zorder=7)
                ax.add_patch(sq)
            else:
                ax.text(px_+0.1, py_+0.1, "90°", color="#10b981", fontsize=7.5)

        # ── Angle labels ──
        for pt_name, ang_label in angles.items():
            lbl_clean = _clean_label(ang_label)
            if "90" in lbl_clean or "right" in lbl_clean.lower(): continue
            if pt_name not in points: continue
            px_,py_ = points[pt_name]
            ax.text(px_, py_-0.38, lbl_clean, color="#f59e0b", fontsize=8,
                    ha="center", va="top",
                    bbox=dict(boxstyle="round,pad=0.1", fc="#07070e", ec="none", alpha=0.7))

        # ── Vertices ──
        ax.scatter(xs, ys, color="#a855f7", s=35, zorder=8)
        for name,(px_,py_) in points.items():
            ox = 0.32 * (1 if px_ >= cx_all else -1)
            oy = 0.32 * (1 if py_ >= cy_all else -1)
            ax.text(px_+ox, py_+oy, name, color="#e9d5ff", fontsize=11,
                    fontweight="bold", ha="center", va="center", zorder=9)

        pad = 1.0
        ax.set_xlim(min(xs)-pad, max(xs)+pad)
        ax.set_ylim(min(ys)-pad, max(ys)+pad)
        ax.set_title(title, color="#e9d5ff", fontsize=10, pad=8)
        ax.tick_params(colors="#6b5b9e", labelsize=7)
        ax.grid(True, color="#28204a", alpha=0.35, linewidth=0.5)
        for sp in ax.spines.values(): sp.set_color("#28204a")

        plt.tight_layout()
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=120, bbox_inches="tight", facecolor="#07070e")
        plt.close(fig)
        return base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception as e:
        _log(f"make_diagram error: {e}")
        return None
    """Draw a geometry diagram faithfully copying what the AI extracted from the image.
    Renders tick marks, arc marks, right angle boxes, parallel arrows, angle labels."""
    if not HAS_PLOT: return None
    try:
        import math as _math
        points      = diagram_data.get("points", {})
        edges       = diagram_data.get("edges", [])
        angles      = diagram_data.get("angles", {})
        labels      = diagram_data.get("labels", {})
        title       = diagram_data.get("title", "Diagram")
        tick_marks  = diagram_data.get("tick_marks", {})   # {"AB":1,"CD":1,"EF":2}
        arc_marks   = diagram_data.get("arc_marks", {})    # {"A":1,"B":2}
        right_angles= diagram_data.get("right_angles", []) # ["C"]
        par_marks   = diagram_data.get("parallel_marks", []) # [["AB","CD"]]

        if not points: return None

        fig, ax = plt.subplots(figsize=(6.0, 5.0), facecolor="#07070e")
        ax.set_facecolor("#0d0c18")
        ax.set_aspect("equal")

        xs = [v[0] for v in points.values()]
        ys = [v[1] for v in points.values()]
        cx_all, cy_all = sum(xs)/len(xs), sum(ys)/len(ys)

        def _unit(ax_, ay_, bx, by):
            dx, dy = bx-ax_, by-ay_
            l = max(_math.hypot(dx, dy), 1e-9)
            return dx/l, dy/l

        def _perp(ux, uy): return -uy, ux

        # ── Draw edges ──
        for edge in edges:
            if len(edge) < 2: continue
            a, b = edge[0], edge[1]
            if a not in points or b not in points: continue
            x0,y0 = points[a]; x1,y1 = points[b]
            ax.plot([x0,x1],[y0,y1], color="#a855f7", linewidth=2.0, zorder=3, solid_capstyle="round")

            # Edge label at midpoint
            key1, key2 = a+b, b+a
            lbl = labels.get(key1) or labels.get(key2)
            if lbl:
                mx,my = (x0+x1)/2,(y0+y1)/2
                ux,uy = _unit(x0,y0,x1,y1)
                px,py = _perp(ux,uy)
                ax.text(mx+px*0.35, my+py*0.35, lbl, color="#f59e0b", fontsize=8.5,
                        ha="center", va="center",
                        bbox=dict(boxstyle="round,pad=0.15", fc="#07070e", ec="none", alpha=0.85))

            # ── Tick marks ──
            n_ticks = tick_marks.get(key1) or tick_marks.get(key2) or 0
            if n_ticks:
                mx,my = (x0+x1)/2,(y0+y1)/2
                ux,uy = _unit(x0,y0,x1,y1)
                px,py = _perp(ux,uy)
                tick_len = 0.18
                spacing = 0.13
                offsets = [0] if n_ticks == 1 else [-spacing/2, spacing/2] if n_ticks == 2 else [-spacing, 0, spacing]
                for off in offsets:
                    tx = mx + ux*off; ty = my + uy*off
                    ax.plot([tx-px*tick_len, tx+px*tick_len],
                            [ty-py*tick_len, ty+py*tick_len],
                            color="#10b981", linewidth=2.0, zorder=6)

            # ── Parallel arrow marks ──
            par_count = 0
            for pi, pair in enumerate(par_marks):
                if key1 in pair or key2 in pair:
                    par_count = pi + 1
                    break
            if par_count:
                mx,my = (x0+x1)/2,(y0+y1)/2
                ux,uy = _unit(x0,y0,x1,y1)
                arr_size = 0.15
                for rep in range(par_count):
                    off = (rep - (par_count-1)/2) * 0.2
                    ax_ = mx + ux*off; ay_ = my + uy*off
                    ax.annotate("", xy=(ax_+ux*arr_size, ay_+uy*arr_size),
                                xytext=(ax_-ux*arr_size, ay_-uy*arr_size),
                                arrowprops=dict(arrowstyle="-|>", color="#60a5fa", lw=1.5))

        # ── Arc marks on angles ──
        for pt_name, n_arcs in arc_marks.items():
            if pt_name not in points: continue
            px_,py_ = points[pt_name]
            connected_pts = []
            for edge in edges:
                if pt_name == edge[0] and edge[1] in points:
                    connected_pts.append(edge[1])
                elif pt_name == edge[1] and edge[0] in points:
                    connected_pts.append(edge[0])
            if len(connected_pts) < 2: continue
            a_pt, b_pt = connected_pts[0], connected_pts[1]
            ax_,ay_ = points[a_pt]; bx,by = points[b_pt]
            ang1 = _math.degrees(_math.atan2(ay_-py_, ax_-px_))
            ang2 = _math.degrees(_math.atan2(by-py_, bx-px_))
            if ang1 > ang2: ang1, ang2 = ang2, ang1
            n_arcs = int(n_arcs) if isinstance(n_arcs, (int,float)) else 1
            arc_colors = {1:"#f59e0b", 2:"#f59e0b"}
            for arc_i in range(n_arcs):
                r = 0.28 + arc_i * 0.12
                theta = _np.linspace(_math.radians(ang1), _math.radians(ang2), 40)
                ax.plot(px_ + r*_np.cos(theta), py_ + r*_np.sin(theta),
                        color=arc_colors.get(n_arcs,"#f59e0b"), linewidth=1.8, zorder=6)

        # ── Right angle boxes ──
        right_angle_pts = list(right_angles or [])
        # Also check angles dict
        for pt_name, ang_label in angles.items():
            if "90" in str(ang_label) or "right" in str(ang_label).lower():
                if pt_name not in right_angle_pts:
                    right_angle_pts.append(pt_name)
        for pt_name in right_angle_pts:
            if pt_name not in points: continue
            px_,py_ = points[pt_name]
            connected = []
            for edge in edges:
                if pt_name == edge[0] and edge[1] in points: connected.append(edge[1])
                elif pt_name == edge[1] and edge[0] in points: connected.append(edge[0])
            if len(connected) >= 2:
                s = 0.22
                u1x,u1y = _unit(px_,py_,*points[connected[0]])
                u2x,u2y = _unit(px_,py_,*points[connected[1]])
                sq = plt.Polygon([
                    [px_+s*u1x,          py_+s*u1y],
                    [px_+s*u1x+s*u2x,    py_+s*u1y+s*u2y],
                    [px_+s*u2x,          py_+s*u2y],
                    [px_,                py_],
                ], fill=False, edgecolor="#10b981", linewidth=1.5, zorder=7)
                ax.add_patch(sq)
            else:
                # Just label it
                ax.text(px_+0.1, py_+0.1, "90°", color="#10b981", fontsize=7.5)

        # ── Non-right angle text labels ──
        for pt_name, ang_label in angles.items():
            if "90" in str(ang_label) or "right" in str(ang_label).lower(): continue
            if pt_name not in points: continue
            px_,py_ = points[pt_name]
            ax.text(px_, py_-0.38, str(ang_label), color="#f59e0b", fontsize=8,
                    ha="center", va="top",
                    bbox=dict(boxstyle="round,pad=0.1", fc="#07070e", ec="none", alpha=0.7))

        # ── Points and vertex labels ──
        ax.scatter(xs, ys, color="#a855f7", s=35, zorder=8)
        for name,(px_,py_) in points.items():
            ox = 0.32 * (1 if px_ >= cx_all else -1)
            oy = 0.32 * (1 if py_ >= cy_all else -1)
            ax.text(px_+ox, py_+oy, name, color="#e9d5ff", fontsize=11,
                    fontweight="bold", ha="center", va="center", zorder=9)

        # ── Padding and style ──
        pad = 1.1
        ax.set_xlim(min(xs)-pad, max(xs)+pad)
        ax.set_ylim(min(ys)-pad, max(ys)+pad)
        ax.set_title(title, color="#e9d5ff", fontsize=10, pad=8)
        ax.tick_params(colors="#6b5b9e", labelsize=7)
        ax.grid(True, color="#28204a", alpha=0.35, linewidth=0.5)
        for sp in ax.spines.values(): sp.set_color("#28204a")

        plt.tight_layout()
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=120, bbox_inches="tight", facecolor="#07070e")
        plt.close(fig)
        return base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception as e:
        _log(f"make_diagram error: {e}")
        return None

def make_graph(eq_str):
    if not HAS_PLOT: return None
    eq = eq_str.strip()
    if eq.startswith("y="): expr = eq[2:]
    elif eq.startswith("f(x)="): expr = eq[5:]
    else: expr = eq
    expr_py = expr.replace("^","**").replace("sin","np.sin").replace("cos","np.cos").replace("tan","np.tan").replace("sqrt","np.sqrt").replace("abs","np.abs").replace("log","np.log").replace("ln","np.log").replace("exp","np.exp").replace("pi","np.pi").replace("e","np.e")
    x = np.linspace(-10, 10, 600)
    ns = {"x": x, "np": np, "__builtins__": {}}
    try: y = eval(expr_py, ns)
    except Exception: return None
    y = np.where(np.isfinite(y), y, np.nan)
    fig, ax = plt.subplots(figsize=(5.5, 3.8), facecolor="#07070e")
    ax.set_facecolor("#0d0c18")
    ax.plot(x, y, color="#a855f7", linewidth=2.2, zorder=3)
    ax.axhline(0, color="#6b5b9e", linewidth=0.8, linestyle="--", alpha=0.7)
    ax.axvline(0, color="#6b5b9e", linewidth=0.8, linestyle="--", alpha=0.7)
    ax.grid(True, color="#28204a", alpha=0.6, linewidth=0.8)
    ax.tick_params(colors="#c4b5fd", labelsize=8)
    for sp in ax.spines.values(): sp.set_color("#28204a")
    ax.set_title(f"y = {expr}", color="#e9d5ff", fontsize=10, pad=8, fontfamily="monospace")
    ax.set_xlabel("x", color="#6b5b9e", fontsize=9); ax.set_ylabel("y", color="#6b5b9e", fontsize=9, rotation=0)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=110, bbox_inches="tight", facecolor="#07070e")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode("utf-8")


# ══════════════════════════════════════════════════════════════════════
#  AI CALLS
# ══════════════════════════════════════════════════════════════════════
def get_client():
    api_key = cfg.get("groq_api_key", "").strip()
    if not api_key: raise RuntimeError("No Groq API key — add it in ⚙ Settings")
    if not HAS_GROQ: raise RuntimeError("groq not installed — run install.bat")
    return Groq(api_key=api_key)

# ── Model constants — one place to update ──
MODEL_TEXT   = "llama-3.3-70b-versatile"                        # text + reasoning (fast, reliable)
MODEL_VISION = "meta-llama/llama-4-scout-17b-16e-instruct"      # vision (multimodal, available on Groq)
MODEL_GEMINI = "gemini-2.0-flash"                                # primary image scanner (use if key set)

def _gemini_vision(b64: str, prompt: str) -> str:
    """
    Send image + prompt to Gemini 2.0 Flash.
    Returns the raw text response, or raises on failure.
    """
    if not HAS_GEMINI:
        raise RuntimeError("google-generativeai not installed — run: pip install google-generativeai")
    key = cfg.get("gemini_api_key", "").strip()
    if not key:
        raise RuntimeError("No Gemini API key")
    genai.configure(api_key=key)
    model = genai.GenerativeModel(MODEL_GEMINI)
    # Decode b64 to bytes for Gemini
    img_bytes = base64.b64decode(b64)
    import PIL.Image
    img = PIL.Image.open(io.BytesIO(img_bytes))
    response = model.generate_content([prompt, img])
    return response.text.strip()

def _safe_json_parse(raw: str) -> list:
    """
    Robustly parse a JSON array from AI output.
    Handles invalid backslash escapes (LaTeX), markdown fences, trailing commas.
    """
    import re as _re, json as _json

    text = _re.sub(r'```(?:json)?\s*', '', raw).strip().strip('`').strip()
    m = _re.search(r'\[.*\]', text, _re.DOTALL)
    if not m:
        return []
    text = m.group(0)

    def fix_escapes(s):
        result = []
        i = 0
        while i < len(s):
            if s[i] == '\\' and i + 1 < len(s):
                nxt = s[i+1]
                if nxt in ('"', '\\', '/', 'b', 'f', 'n', 'r', 't'):
                    result.append(s[i]); result.append(nxt); i += 2
                elif nxt == 'u' and i+5 < len(s) and all(c in '0123456789abcdefABCDEF' for c in s[i+2:i+6]):
                    result.append(s[i:i+6]); i += 6
                else:
                    # Read full LaTeX command name
                    j = i + 1
                    while j < len(s) and s[j].isalpha():
                        j += 1
                    cmd = s[i+1:j]
                    latex_map = {
                        'angle':'∠','cong':'≅','triangle':'△','overline':'',
                        'sim':'~','perp':'⊥','parallel':'∥','neq':'≠',
                        'leq':'≤','geq':'≥','cdot':'·','times':'×',
                        'div':'÷','sqrt':'√','pi':'π','theta':'θ',
                        'alpha':'α','beta':'β','circ':'°','frac':'',
                        'left':'','right':'','Delta':'△','Sigma':'Σ',
                        'vec':'','hat':'','bar':'','mathbf':'','text':'',
                    }
                    result.append(latex_map.get(cmd, ''))
                    i = j
            else:
                result.append(s[i]); i += 1
        return ''.join(result)

    text = fix_escapes(text)
    # Remove trailing commas before ] or }
    text = _re.sub(r',\s*([}\]])', r'\1', text)

    try:
        return _json.loads(text)
    except Exception:
        try:
            # Nuclear fallback — strip ALL remaining lone backslashes
            text2 = _re.sub(r'\\(?!["\\/bfnrtu])', '', text)
            return _json.loads(text2)
        except Exception as e:
            _log(f"_safe_json_parse failed: {e} | raw[:200]={raw[:200]}")
            return []


def _strip_latex(text: str) -> str:
    """Convert LaTeX math notation to plain readable text."""
    if not text or '$' not in text and '\\' not in text:
        return text
    import re
    # Remove $...$ and $$...$$ wrappers
    text = re.sub(r'\$\$(.+?)\$\$', r'\1', text, flags=re.DOTALL)
    text = re.sub(r'\$(.+?)\$', r'\1', text)
    # Convert common LaTeX commands to Unicode
    text = text.replace(r'\cong', '≅')
    text = text.replace(r'\angle', '∠')
    text = text.replace(r'\triangle', '△')
    text = text.replace(r'\overline', '')
    text = text.replace(r'\vec', '')
    text = text.replace(r'\sim', '~')
    text = text.replace(r'\perp', '⊥')
    text = text.replace(r'\parallel', '∥')
    text = text.replace(r'\neq', '≠')
    text = text.replace(r'\leq', '≤')
    text = text.replace(r'\geq', '≥')
    text = text.replace(r'\cdot', '·')
    text = text.replace(r'\times', '×')
    text = text.replace(r'\div', '÷')
    text = text.replace(r'\sqrt', '√')
    text = text.replace(r'\pi', 'π')
    text = text.replace(r'\theta', 'θ')
    text = text.replace(r'\alpha', 'α')
    text = text.replace(r'\beta', 'β')
    text = text.replace(r'\circ', '°')
    text = text.replace(r'\frac', '')
    text = text.replace(r'\left', '').replace(r'\right', '')
    text = text.replace(r'\{', '{').replace(r'\}', '}')
    # Remove remaining backslash commands
    text = re.sub(r'\\[a-zA-Z]+\s*', '', text)
    # Remove leftover braces
    text = text.replace('{', '').replace('}', '')
    # Clean up extra spaces
    text = re.sub(r' +', ' ', text).strip()
    return text


def _strip_latex_from_problem(p: dict) -> dict:
    """Recursively strip LaTeX from all string fields in a problem dict."""
    for key, val in p.items():
        if isinstance(val, str):
            p[key] = _strip_latex(val)
        elif isinstance(val, list):
            p[key] = [_strip_latex(v) if isinstance(v, str) else v for v in val]
    return p


def ai_scan(b64, strictness="flag_all", examine_examples=False, extra_context=None, extra_images=None):
    client = get_client()
    strict_note = ("Flag ANY answer that looks even slightly off, unclear, or potentially wrong."
        if strictness == "flag_all" else "Only flag answers that are clearly and definitively wrong.")

    # Build the full prompt text
    ctx_prefix = ""
    if extra_context and any(c.strip() for c in extra_context if c):
        ctx_prefix = "EXTRA CONTEXT (use this to help answer questions):\n" + "\n\n".join(c.strip() for c in extra_context if c and c.strip()) + "\n\n"

    scan_prompt = (
        ctx_prefix +
        "You are scanning a student's assignment screenshot. Read the ENTIRE screen carefully.\n\n"
        "══ PLATFORM DETECTION — identify which platform this is first ══\n\n"
        "── GOOGLE DOCS / WORD / WORKSHEET ──\n"
        "Layout: white document page centered on gray background. Toolbar at top (File/Edit/View/Insert etc) is NOT content.\n"
        "TWO-COLUMN split is very common: SOURCE SIDE (reading passage/article) | ANSWER SIDE (labeled fields).\n"
        "ANSWER FIELDS to detect:\n"
        "  • 'Answer: |' or 'Answer: ___' = unanswered (cursor or blank line)\n"
        "  • 'Main Idea:', 'Summary:', 'Key Details:', 'Response:', 'Evidence:', 'Explanation:' followed by blank/cursor\n"
        "  • Underlined blank spaces '________' = fill-in-the-blank\n"
        "  • Empty table cells next to a question label\n"
        "  • Highlighted text boxes (yellow/green highlight on the field label) = where to write\n"
        "ANSWER LOGIC:\n"
        "  → Find the CLOSEST source text (heading/paragraph) the field relates to\n"
        "  → 'Main Idea' = 1-2 sentence summary of that passage's central point\n"
        "  → 'Summary' = condense passage into 2-3 sentences\n"
        "  → 'Key Details' = bullet the most important specific facts\n"
        "  → 'Evidence' = direct quote or paraphrase supporting the claim\n"
        "  → Fill-in-the-blank = one word or short phrase that fits grammatically\n\n"
        "── GOOGLE FORMS ──\n"
        "Layout: white cards with purple/blue accents, question title at top of each card.\n"
        "Question types:\n"
        "  • RADIO BUTTONS (○ filled ● = selected, ○ empty = not selected) → MULTIPLE_CHOICE\n"
        "  • CHECKBOXES (☐ empty, ☑ checked) → MULTIPLE_CHOICE (select all that apply)\n"
        "  • DROPDOWN (shows selected value or 'Choose') → MULTIPLE_CHOICE\n"
        "  • SHORT ANSWER / PARAGRAPH (text input box, empty or filled) → WRITTEN\n"
        "  • LINEAR SCALE (1-5 or 1-10 scale) → MULTIPLE_CHOICE\n"
        "  • A filled radio/checkbox = answered. An empty text box = unanswered.\n"
        "  • Required questions marked with * (red asterisk)\n\n"
        "── GOOGLE CLASSROOM / ASSIGNMENT SHEETS ──\n"
        "May show: assignment title, instructions block, then numbered questions below.\n"
        "Student responses appear in text boxes or typed inline after the question.\n"
        "Blank after question number = unanswered even if there's a cursor.\n\n"
        "── MATCHING ACTIVITIES ──\n"
        "Two columns: Column A (terms) and Column B (definitions/answers).\n"
        "Lines or letters/numbers connect them. Unmatched items = unanswered.\n"
        "Represent each match as a separate question: 'Match: [term]' → answer is the matching definition.\n\n"
        "── TABLES / GRIDS ──\n"
        "Each empty table cell adjacent to a row/column header = a question.\n"
        "question_label = the row header. question = what's being asked per column header.\n\n"
        "── HOW TO FIND ALL QUESTIONS ──\n"
        "- Numbered: '1.', '2.', 'Q1', 'a)', 'b)', '①②③' — all questions\n"
        "- Labeled fields: 'Answer:', 'Main Idea:', 'Summary:', 'Fill in:', 'Response:'\n"
        "- Empty boxes, blank lines, underlines, empty table cells\n"
        "- Yellow/green highlighted labels = answer goes here\n"
        "- Geometric diagrams next to numbers = geometry/math questions\n"
        "- Include ALL questions even if already answered\n\n"
        "── FOR EACH QUESTION ──\n"
        "1. question_label: label/number if visible (e.g. 'Main Idea', '1.', 'Q2') — null if none\n"
        "2. question: what is being asked.\n"
        "3. type: MULTIPLE_CHOICE, TRUE_FALSE, or WRITTEN\n"
        "4. answered: true ONLY if a real answer is already written\n"
        "5. user_answer: what they wrote (null if unanswered)\n"
        "6. correct_answer: YOUR complete answer\n"
        "7. is_correct: true/false/null\n"
        f"8. {strict_note}\n"
        "9. correction: if wrong, the corrected answer\n"
        "10. confident: false only if you genuinely cannot read the content\n"
        + ("11. Where helpful, include a brief worked example.\n\n" if examine_examples else "11. Do NOT include worked examples.\n\n") +
        "Return ONLY valid JSON array (no markdown):\n"
        '[{"question_label":"1.","question":"...","type":"WRITTEN","answered":false,"user_answer":null,'
        '"correct_answer":"...","is_correct":null,"correction":null,"confident":true}]\n'
        "If no questions found: []"
    )

    # Gemini first, Maverick fallback
    raw = ""
    try:
        raw = re.sub(r"```(?:json)?|```", "", _gemini_vision(b64, scan_prompt)).strip()
        _log("Scan: Gemini 2.0 Flash")
    except Exception as gem_err:
        _log(f"Gemini scan failed ({gem_err}), falling back to Maverick")
        img_parts = [{"type":"image_url","image_url":{"url":f"data:image/png;base64,{b64}"}}]
        for ei in (extra_images or []):
            if ei: img_parts.append({"type":"image_url","image_url":{"url":f"data:image/jpeg;base64,{ei}" if not ei.startswith("data:") else ei}})
        resp = client.chat.completions.create(
            model=MODEL_VISION,
            messages=[{"role":"user","content":img_parts + [{"type":"text","text":scan_prompt}]}],
            max_tokens=3500, temperature=0.1
        )
        raw = re.sub(r"```(?:json)?|```", "", resp.choices[0].message.content.strip()).strip()

    results = _safe_json_parse(raw)
    # Strip any LaTeX notation Gemini might have used
    for item in results:
        for k, v in item.items():
            if isinstance(v, str): item[k] = _strip_latex(v)
    return results

def _preprocess_math_image(b64: str) -> str:
    """
    Pre-processing pipeline — boost contrast, sharpen, upscale small images.
    Returns new b64. Helps massively with low-res screenshots and handwriting.
    """
    try:
        raw = base64.b64decode(b64)
        img = Image.open(io.BytesIO(raw)).convert("RGB")
        w, h = img.size
        # Upscale if tiny — AI misses small tick marks on small images
        if w < 800 or h < 600:
            scale = max(800/w, 600/h, 1.5)
            img = img.resize((int(w*scale), int(h*scale)), Image.LANCZOS)
        # Contrast boost — makes tick marks, angle arcs, labels pop
        img = ImageEnhance.Contrast(img).enhance(1.4)
        # Sharpness — crisps up printed text and diagram lines
        img = ImageEnhance.Sharpness(img).enhance(2.0)
        # Brightness — slightly brighten so light pencil marks show
        img = ImageEnhance.Brightness(img).enhance(1.05)
        buf = io.BytesIO(); img.save(buf, format="PNG", optimize=False); buf.seek(0)
        return base64.b64encode(buf.read()).decode()
    except Exception as e:
        _log(f"_preprocess_math_image failed: {e}")
        return b64


def _ocr_hint(b64: str) -> str:
    """
    OCR pre-pass using pytesseract — extracts raw text from image.
    AI uses this as a hint so it doesn't misread printed text.
    Returns empty string if tesseract not installed.
    """
    try:
        import pytesseract
        raw = base64.b64decode(b64)
        img = Image.open(io.BytesIO(raw))
        text = pytesseract.image_to_string(img, config="--psm 6").strip()
        if text:
            return f"\n\n[OCR PRE-SCAN — raw text extracted from image, use as reading hint]:\n{text}\n"
        return ""
    except Exception:
        return ""  # tesseract not installed — silently skip


def _formula_lib_context() -> str:
    """Pull relevant past solutions from formula library as context hint."""
    try:
        lib = cfg.get("formula_library", [])
        if not lib: return ""
        recent = lib[-8:]  # last 8 formulas
        lines = [f"• {f.get('problem','')} → {f.get('answer','')}" for f in recent if f.get("answer")]
        if lines:
            return "\n\n[FORMULA LIBRARY — similar problems you've solved before, use as reference]:\n" + "\n".join(lines) + "\n"
        return ""
    except Exception:
        return ""


def _detect_format_pass(client, b64: str) -> str:
    """Fast format detection — Gemini first for better accuracy."""
    prompt = (
        "Look at this math worksheet image. In ONE short phrase, describe its format.\n"
        "Choose from: MULTIPLE_CHOICE_GEOMETRY, MULTIPLE_CHOICE_ALGEBRA, FREE_RESPONSE_EQUATIONS, "
        "WORD_PROBLEMS, TABLE_OF_VALUES, GRAPH_PROBLEMS, MIXED_WORKSHEET, ARITHMETIC_DRILL, PROOF_PROBLEMS.\n"
        "Also note: does it have hand-drawn diagrams? (yes/no). Does it have pre-filled answers? (yes/no).\n"
        "Reply in this exact format: FORMAT|has_diagrams:yes|has_prefilled:no"
    )
    try:
        return _gemini_vision(b64, prompt)
    except Exception:
        pass
    try:
        resp = client.chat.completions.create(
            model=MODEL_VISION,
            messages=[{"role":"user","content":[
                {"type":"image_url","image_url":{"url":f"data:image/png;base64,{b64}"}},
                {"type":"text","text":prompt}
            ]}],
            max_tokens=60, temperature=0.0
        )
        return resp.choices[0].message.content.strip()
    except Exception:
        return "MULTIPLE_CHOICE_GEOMETRY|has_diagrams:yes|has_prefilled:no"


def _verify_geometry_answers(b64: str, problems: list) -> list:
    """
    Self-verification pass for geometry MCQ problems.
    Re-checks each answer against the image using the explicit reasoning chain.
    Only runs when there are geometry MCQ problems.
    """
    geo_problems = [p for p in problems if
        "CHOICE" in (p.get("problem_type","")).upper() and
        any(kw in (p.get("solving_method","") + p.get("problem","")).upper()
            for kw in ["ASA","AAS","NEITHER","CONGRUENCE","TRIANGLE"])]

    if not geo_problems:
        return problems

    # Build verification prompt
    answers_summary = "\n".join(
        f"Problem {p.get('problem_label','?')}: I answered '{p.get('answer','')}'"
        for p in geo_problems
    )

    verify_prompt = (
        "I just solved these geometry congruence problems from this worksheet image:\n\n"
        + answers_summary +
        "\n\nFor EACH problem, verify my answer by doing this EXACT check:\n"
        "1. Look at the specific diagram for that problem number\n"
        "2. Find AUTO-GIVEN parts: shared side (Reflexive), vertical angles (X-shape), right angles (box symbol)\n"
        "3. Read EXPLICIT marks: tick marks on sides, arc marks on angles\n"
        "4. Count total: ___ angle pairs, ___ side pairs\n"
        "5. Is any marked side BETWEEN the two marked angles? YES=ASA, NO=AAS\n"
        "6. Is the arrangement SSA (2 sides + non-included angle)? → NEITHER\n"
        "7. Is my answer correct?\n\n"
        "USE PLAIN TEXT ONLY — no LaTeX, no $, no backslashes.\n"
        "Return ONLY valid JSON array:\n"
        '[{"problem_label":"1.","my_answer":"B) AAS","verdict":"correct","correct_answer":"B) AAS","reasoning":"Arc at A and B, tick on AC (non-included side), AAS confirmed"}]\n'
        "If my answer is correct, set verdict=correct and correct_answer=same as my_answer.\n"
        "If wrong, set verdict=wrong and provide correct_answer."
    )

    try:
        raw = ""
        try:
            raw = _gemini_vision(b64, verify_prompt)
            _log("Geometry verify pass: Gemini 2.0 Flash")
        except Exception as e:
            _log(f"Gemini verify failed ({e}), skipping verification")
            return problems

        corrections = _safe_json_parse(raw)
        if not corrections:
            return problems

        # Apply corrections
        corr_map = {c.get("problem_label"): c for c in corrections if c.get("verdict") == "wrong"}
        for p in problems:
            lbl = p.get("problem_label")
            if lbl in corr_map:
                c = corr_map[lbl]
                old = p.get("answer","")
                new = c.get("correct_answer","")
                if new and new != old:
                    _log(f"Verify corrected {lbl}: {old} → {new}")
                    p["answer"] = _strip_latex(new)
                    p["solving_method"] = p.get("solving_method","") + " [verified+corrected]"
                    p["confidence"] = min(p.get("confidence",8), 8)  # cap at 8 since it needed correction

        return problems
    except Exception as e:
        _log(f"_verify_geometry_answers failed: {e}")
        return problems


def _format_aware_extra_instructions(fmt: str) -> str:
    """Return extra prompt instructions based on detected format."""
    f = fmt.upper()
    extra = ""
    if "MULTIPLE_CHOICE" in f or "MCQ" in f:
        extra += (
            "FORMAT: MULTIPLE CHOICE WORKSHEET\n"
            "- Read ALL options for each question CAREFULLY\n"
            "- Check if any answer is already filled/bubbled — still give YOUR OWN correct answer\n"
            "- For 'ASA, AAS, or Neither' problems use this EXACT reasoning chain:\n"
            "  1. AUTO-GIVEN FIRST (before looking at marks):\n"
            "     → SHARED SIDE: does a single line/segment form part of BOTH triangles? If yes → Reflexive Property, it's equal\n"
            "     → VERTICAL ANGLES: do the triangles share a vertex where lines cross (X/bowtie shape)? If yes → vertical angles equal\n"
            "     → RIGHT ANGLES: is there a right angle box or perpendicular symbol? Both angles at that line = 90°\n"
            "  2. READ EXPLICIT MARKS: name the exact segments/angles with tick/arc marks\n"
            "  3. COMPILE: list all congruent pairs (auto-given + explicit). Count angle pairs and side pairs.\n"
            "  4. THEOREM TEST:\n"
            "     → ASA: 2 angle pairs + 1 side pair where side is BETWEEN the two angles\n"
            "     → AAS: 2 angle pairs + 1 side pair where side is OUTSIDE the two angles\n"
            "     → SSS: 3 side pairs\n"
            "     → SAS: 2 sides + angle BETWEEN them\n"
            "     → SSA or ASS (2 sides + non-included angle) = NEITHER, not valid\n"
            "     → Not enough pairs for any theorem = NEITHER\n"
            "  5. INCLUDED SIDE TEST for ASA vs AAS:\n"
            "     Look at the two marked angles. Is the marked side SANDWICHED between them? YES=ASA, NO=AAS\n"
            "- For 'what ADDITIONAL information is needed' problems:\n"
            "  1. Read triangle names for correspondence (△ABC≅△DEF → A↔D B↔E C↔F)\n"
            "  2. Catalog GIVEN: ticks, arcs, shared sides, vertical angles\n"
            "  3. Find MISSING piece from correspondence\n"
            "  4. Match to answer — REJECT any choice already in GIVEN catalog\n"
            "- Answer format: 'X) [full option text]'\n\n"
        )
    if "GEOMETRY" in f or "has_diagrams:yes" in fmt:
        extra += (
            "DIAGRAM READING — THIS IS CRITICAL:\n"
            "- Single tick mark (|) on a side = that side is congruent to other single-tick sides\n"
            "- Double tick mark (||) = congruent to other double-tick sides\n"
            "- Arc mark on an angle = that angle is congruent to other arc-marked angles\n"
            "- Double arc = congruent to other double-arc angles\n"
            "- Right angle box = 90 degrees\n"
            "- Arrow on a line = the line extends to infinity (parallel lines use matching arrows)\n"
            "- For congruence proofs (ASA/AAS/SSS/SAS/HL): state EXACTLY which parts are marked congruent\n"
            "  and which theorem applies\n"
            "- COPY THE DIAGRAM FAITHFULLY in diagram_data — use the actual letter labels visible\n"
            "  Place points at proportional coordinates matching their visual position in the image\n\n"
        )
    if "has_prefilled:yes" in fmt:
        extra += (
            "PRE-FILLED ANSWERS DETECTED:\n"
            "- Some answers may already be written/bubbled in by the student\n"
            "- Note pre-filled answers as 'Student answered: X' in your steps\n"
            "- Always give the CORRECT answer regardless — mark if student was right or wrong\n\n"
        )
    if "PROOF" in f:
        extra += (
            "FORMAT: PROOF PROBLEMS\n"
            "- Build a two-column proof: Statement | Reason\n"
            "- Each step in vertical_method = 'Statement | Reason'\n"
            "- Common reasons: Given, Definition of congruent angles/segments, Reflexive Property,\n"
            "  Vertical Angles Theorem, AAS/ASA/SSS/SAS/HL Theorem, CPCTC\n\n"
        )
    if "TABLE" in f:
        extra += (
            "FORMAT: TABLE OF VALUES\n"
            "- Extract the table headers and all rows as text\n"
            "- If asked to complete the table, fill in each missing cell\n"
            "- If asked to graph from a table, include graph_eq if a pattern/function is clear\n\n"
        )
    return extra


def ai_math(b64, examine_examples=False, extra_context=None, extra_images=None):
    client = get_client()

    # ── FEATURE 2: Pre-process image ──
    b64_processed = _preprocess_math_image(b64)

    # ── FEATURE 7: OCR pre-pass hint ──
    ocr_hint = _ocr_hint(b64_processed)

    # ── FEATURE 8: Formula library context ──
    formula_ctx = _formula_lib_context()

    # ── EXAMPLES: Inject any saved example methods ──
    examples_ctx = _examples_context()

    # ── FEATURE 3: Format detection pass ──
    fmt = _detect_format_pass(client, b64_processed)
    format_instructions = _format_aware_extra_instructions(fmt)

    # Extra images
    extra_img_parts = []
    for ei in (extra_images or []):
        if ei:
            url = ei if ei.startswith("data:") else f"data:image/jpeg;base64,{ei}"
            extra_img_parts.append({"type":"image_url","image_url":{"url":url}})

    # Extra context block
    ctx_block = ""
    if extra_context and any(c.strip() for c in extra_context if c):
        ctx_block = "EXTRA CONTEXT (use this to help solve problems):\n" + "\n\n".join(c.strip() for c in extra_context if c and c.strip()) + "\n\n"

    # ── Build the master prompt ──
    master_prompt = (
        ctx_block + ocr_hint + formula_ctx + examples_ctx +
        "You are an expert math tutor and diagram reader. Study the image CAREFULLY and solve EVERY problem you see.\n\n"

        "READING THE IMAGE — CRITICAL:\n"
        "- Look at the ENTIRE image, top to bottom, left to right\n"
        "- READ every shape, diagram, label, number, angle mark, tick mark, and arc mark you can see\n"
        "- For diagrams: describe what you literally SEE (triangles, lines, tick marks, arc marks, boxes)\n"
        "- Do NOT assume — only report what is visually present\n"
        "- Tick marks on sides = congruent sides. Arc marks on angles = congruent angles.\n"
        "- Right angle box = 90 degrees. Shared side = Reflexive Property. X-crossing = Vertical Angles.\n\n"

        "GEOMETRY THEOREM QUICK GUIDE:\n"
        "- ASA: 2 angles + side BETWEEN them\n"
        "- AAS: 2 angles + side NOT between them\n"
        "- SAS: 2 sides + angle BETWEEN them\n"
        "- SSS: 3 pairs of equal sides\n"
        "- HL: right triangle, hypotenuse + leg\n"
        "- Neither: if marks form SSA or are insufficient\n\n"

        "FOR EACH PROBLEM return a JSON object with these fields:\n"
        "- problem_label: the number/letter (e.g. '1.', 'Q2')\n"
        "- problem_type: GEOMETRY | EQUATION_SOLVE | ARITHMETIC | WORD_PROBLEM | MULTIPLE_CHOICE | FILL_BLANK | GRAPH_PROBLEM\n"
        "- problem: full description — for diagrams describe what you see AND what is asked\n"
        "- answer: final answer. For MCQ include letter + full text. For geometry include theorem used.\n"
        "- confidence: 1-10\n"
        "- steps: array of step strings\n"
        "- explanations: array of why strings (SAME length as steps)\n"
        "- solving_method: short method name\n"
        "- vertical_method: column layout string\n"
        "- horizontal_steps: array of step chains\n"
        "- graph_eq: 'y=expr' if plottable, else null\n"
        "- has_graph: true if graph_eq set\n"
        "- diagram_data: for geometry problems with visible diagrams:\n"
        '  {"title":"...","points":{"A":[x,y],...},"edges":[["A","B"],...],'
        '"tick_marks":{"AB":1},"arc_marks":{"A":1},"right_angles":["C"],'
        '"angles":{"A":"45"},"labels":{"AB":"6"},"description":"..."} or null\n'
        "- diagram_description: plain English description of the diagram, or null\n\n"

        + ("After solving, include one brief worked example.\n\n" if examine_examples else "") +

        "Return ONLY a valid JSON array, no markdown:\n"
        '[{"problem_type":"ARITHMETIC","solving_method":"Addition","problem_label":"1.",'
        '"problem":"What is 15 + 27?","answer":"42","confidence":10,'
        '"steps":["Add: 15+27=42"],"explanations":["Direct addition"],'
        '"horizontal_steps":["15+27 → 42"],"vertical_method":"  15\\n+ 27\\n----\\n  42",'
        '"graph_eq":null,"has_graph":false,"diagram_data":null,"diagram_description":null}]\n'
        "No math found: []"
    )

    # ── FEATURE 1: Pass 1 — main solve (Gemini first, Maverick fallback) ──
    raw1 = ""
    try:
        gemini_raw = _gemini_vision(b64_processed, master_prompt)
        raw1 = re.sub(r"```(?:json)?|```","", gemini_raw).strip()
        _log("Math pass 1: Gemini 2.0 Flash")
    except Exception as gem_err:
        _log(f"Gemini math pass 1 failed ({gem_err}), falling back to Maverick")
        pass1_result = client.chat.completions.create(
            model=MODEL_VISION,
            messages=[{"role":"user","content":
                [{"type":"image_url","image_url":{"url":f"data:image/png;base64,{b64_processed}"}}]
                + extra_img_parts
                + [{"type":"text","text":master_prompt}]
            }],
            max_tokens=6000, temperature=0.05
        )
        raw1 = re.sub(r"```(?:json)?|```","", pass1_result.choices[0].message.content.strip()).strip()

    problems = _safe_json_parse(raw1)

    # ── FEATURE 1: Pass 2 — catch anything missed (Gemini first, Maverick fallback) ──
    if len(problems) < 6:
        try:
            pass2_prompt = (
                f"I already found {len(problems)} problem(s) in this image: "
                + ", ".join(p.get('problem_label','?') for p in problems if p.get('problem_label'))
                + ".\n\nLook at the image again VERY carefully. Are there ANY problems I missed? "
                "Look especially for: problems in the bottom half, problems hidden near diagrams, "
                "fill-in-blank boxes, problems numbered differently.\n"
                "Return ONLY a JSON array of the MISSED problems in the same format. "
                "If nothing was missed, return []."
            )
            raw2 = ""
            try:
                raw2 = re.sub(r"```(?:json)?|```","", _gemini_vision(b64_processed, pass2_prompt)).strip()
                _log("Math pass 2: Gemini 2.0 Flash")
            except Exception as gem2_err:
                _log(f"Gemini math pass 2 failed ({gem2_err}), falling back to Maverick")
                pass2_result = client.chat.completions.create(
                    model=MODEL_VISION,
                    messages=[{"role":"user","content":
                        [{"type":"image_url","image_url":{"url":f"data:image/png;base64,{b64_processed}"}}]
                        + [{"type":"text","text":pass2_prompt}]
                    }],
                    max_tokens=3000, temperature=0.05
                )
                raw2 = re.sub(r"```(?:json)?|```","", pass2_result.choices[0].message.content.strip()).strip()

            missed = _safe_json_parse(raw2)
            if missed:
                existing_labels = {p.get('problem_label') for p in problems}
                for mp in missed:
                    if mp.get('problem_label') not in existing_labels:
                        problems.append(mp)
                        existing_labels.add(mp.get('problem_label'))
        except Exception as e:
            _log(f"Pass 2 failed (non-critical): {e}")

    # ── FEATURE 6: Flag low-confidence problems + strip LaTeX ──
    for p in problems:
        _strip_latex_from_problem(p)
        conf = p.get("confidence", 8)
        if isinstance(conf, (int, float)) and conf < 6:
            p["low_confidence"] = True
        else:
            p["low_confidence"] = False

    # ── VERIFICATION: Self-check geometry answers ──
    problems = _verify_geometry_answers(b64_processed, problems)

    return problems


def ai_double_check(question):
    client = get_client()
    resp = client.chat.completions.create(
        model=MODEL_TEXT,
        messages=[{"role":"user","content":(
            f"Reconsider this question very carefully:\n\n\"{question}\"\n\n"
            f"Think step by step, then give the single most accurate answer.\n"
            f"MCQ = choose the correct letter + option text.\nTrue/False = answer True or False.\n"
            f"Written = concise, factual answer like a top student.\n"
            f"Reply with ONLY the answer — no explanations, no preamble."
        )}],
        max_tokens=300, temperature=0.15
    )
    return resp.choices[0].message.content.strip()

def ai_qa(question, history, is_followup, extra_context=None, extra_images=None, personality="default"):
    client = get_client()
    has_images = bool(extra_images and any(ei for ei in extra_images))
    model = MODEL_VISION if has_images else MODEL_TEXT

    # Personality system prompts
    _personalities = {
        "default": "You are a sharp, direct assistant. Answer in ONE sentence, max 20 words. No intros. Just the answer.",
        "tutor":   "You are a patient, encouraging tutor. Explain clearly step by step. Use simple language. 2-4 sentences.",
        "funny":   "You are a funny, witty assistant. Answer accurately but keep it casual and entertaining. Throw in a joke if it fits. 2-3 sentences.",
        "expert":  "You are a deep domain expert. Give a thorough, precise answer with nuance and detail. 3-5 sentences. Use technical terms when appropriate.",
        "eli5":    "Explain like the person asking is 5 years old. Use super simple words, short sentences, and fun analogies. 2-3 sentences max.",
        "brutal":  "You are brutally honest. No sugarcoating, no fluff. Say exactly what you think, even if it's harsh. 1-2 sentences.",
        "hype":    "You are an enthusiastic hype man. Answer the question but make it sound AMAZING and exciting. Use caps, energy, and enthusiasm. 2-3 sentences.",
    }

    ctx_block = ""
    if extra_context:
        joined = "\n\n".join(c.strip() for c in extra_context if c.strip())
        if joined: ctx_block = f"\n\n[EXTRA CONTEXT]:\n{joined}"

    def _uc(text):
        if has_images:
            parts = []
            for ei in (extra_images or []):
                if not ei: continue
                url = ei if ei.startswith("data:") else f"data:image/jpeg;base64,{ei}"
                parts.append({"type":"image_url","image_url":{"url":url}})
            parts.append({"type":"text","text":text})
            return parts
        return text

    if is_followup and history:
        messages = list(history) + [{"role":"user","content":_uc(f"{question}{ctx_block}\n\n[Follow-up. Be more detailed. 2-4 sentences max.]")}]
        max_tok = 400
    else:
        if has_images:
            sys_msg = (
                "You are an expert at analyzing images in detail. Look at the entire image carefully and use high-resolution detail. "
                "When identifying living things (animals, plants, fungi, etc.): use specific, precise names — including rare, uncommon, or lesser-known species. "
                "Prefer scientific names when the creature is unusual. For objects, text, or scenes: describe precisely. "
                "Answer directly and concisely but with enough detail to be accurate."
            )
            max_tok = 400
        else:
            sys_msg = _personalities.get(personality, _personalities["default"])
            max_tok = 300 if personality in ("expert","tutor","funny","hype","brutal") else 120
        messages = [{"role":"system","content":sys_msg}, {"role":"user","content":_uc(question+ctx_block)}]

    resp = client.chat.completions.create(model=model, messages=messages, max_tokens=max_tok, temperature=0.5)
    answer = resp.choices[0].message.content.strip()

    personality_sys = _personalities.get(personality, _personalities["default"])
    if not is_followup:
        new_history = [{"role":"system","content":personality_sys},
                       {"role":"user","content":question},{"role":"assistant","content":answer}]
    else:
        new_history = list(history) + [{"role":"user","content":question},{"role":"assistant","content":answer}]
    return answer, new_history


def ai_write(mode, text, extra="", doc_context=""):
    """Handles all Write tab modes: analyze, rewrite, outline, prompt_decode, tone, citation, argument, summary, vocab."""
    client = get_client()
    ctx_block = f"\n\n[DOCUMENT CONTEXT]:\n{doc_context.strip()}" if doc_context and doc_context.strip() else ""

    prompts = {
        "analyze": (
            f"You are an expert writing coach. Analyze this essay or writing:{ctx_block}\n\n"
            f"\"{text}\"\n\n"
            "Return a structured analysis with these exact sections:\n"
            "**THESIS** — Is there a clear thesis? Quote it or note it's missing.\n"
            "**ARGUMENT FLOW** — Does the argument progress logically? Any gaps?\n"
            "**EVIDENCE** — Is evidence used well? Specific or vague?\n"
            "**TONE & VOICE** — Academic, casual, inconsistent? Any passive voice issues?\n"
            "**GRAMMAR & STYLE** — Top 3 grammar or style issues found.\n"
            "**STRENGTHS** — What genuinely works.\n"
            "**GRADE** — Letter grade (A-F) with one sentence justification.\n"
            "**PRIORITY FIXES** — Top 3 most impactful improvements to make.\n"
            "Be specific, reference the actual text. Don't be vague."
        ),
        "rewrite": (
            f"Rewrite the following text. Target style: {extra or 'clear student writing'}.\n"
            f"Keep the same meaning and length. Sound natural, not AI-generated.{ctx_block}\n\n"
            f"Text: {text}\n\nRewritten (return only the rewritten text):"
        ),
        "outline": (
            f"Create a detailed essay outline for this topic or draft:{ctx_block}\n\n\"{text}\"\n\n"
            f"Format: {'Argument style: '+extra if extra else 'standard 5-paragraph essay'}\n"
            "Include:\n"
            "- HOOK idea for intro\n"
            "- THESIS statement\n"
            "- 3 BODY PARAGRAPHS each with: topic sentence, 2 supporting points, transition\n"
            "- CONCLUSION approach\n"
            "Make it specific enough that someone could write from it immediately."
        ),
        "prompt_decode": (
            f"A student has this essay prompt:\n\n\"{text}\"\n\n"
            "Break it down completely:\n"
            "**WHAT THEY'RE REALLY ASKING** — Plain English explanation\n"
            "**KEY REQUIREMENTS** — Every specific thing that must be included\n"
            "**COMMON MISTAKES** — What students usually get wrong on this type of prompt\n"
            "**SUGGESTED APPROACH** — Step-by-step game plan to tackle it\n"
            "**STRONG THESIS STARTER** — Give one example thesis that would score well\n"
            "Be practical. A student should be able to start writing immediately after reading this."
        ),
        "tone": (
            f"Analyze the tone and style of this writing:{ctx_block}\n\n\"{text}\"\n\n"
            "Return:\n"
            "**OVERALL TONE** — One word, then a sentence explanation\n"
            "**FORMALITY LEVEL** — 1-10 scale (1=texting, 10=academic paper), with why\n"
            "**PASSIVE VOICE** — List every passive voice sentence found\n"
            "**REPEATED WORDS** — Words used too often (3+ times)\n"
            "**WEAK PHRASES** — Vague or filler phrases to cut\n"
            "**SENTENCE VARIETY** — Are sentences too similar in structure?\n"
            "**FIXES** — Rewrite the 2 weakest sentences to show improvement\n"
            "Be specific — quote actual lines from the text."
        ),
        "citation": (
            f"Format this source as a proper citation. Style requested: {extra or 'MLA'}.\n\n"
            f"Source info: {text}\n\n"
            "Return:\n"
            "**MLA** format\n"
            "**APA** format\n"
            "**Chicago** format\n"
            "Then give a one-line in-text citation example for MLA and APA."
        ),
        "argument": (
            f"Analyze the argument in this text:{ctx_block}\n\n\"{text}\"\n\n"
            "Return:\n"
            "**MAIN CLAIM** — What is being argued\n"
            "**LOGIC CHECK** — Is the reasoning valid? Any logical fallacies?\n"
            "**WEAKNESSES** — Specific holes in the argument\n"
            "**MISSING EVIDENCE** — What would strengthen it\n"
            "**COUNTERARGUMENT** — The strongest opposing argument someone could make\n"
            "**HOW TO STRENGTHEN** — 3 concrete ways to make this argument more convincing\n"
            "Quote specific lines from the text."
        ),
        "summary": (
            f"Summarize this text:{ctx_block}\n\n\"{text}\"\n\n"
            "Return THREE versions:\n"
            "**ONE SENTENCE** — The absolute core idea in one sentence\n"
            "**ONE PARAGRAPH** — Key points in 3-5 sentences\n"
            "**BULLET BREAKDOWN** — 5-8 bullet points of the most important facts/ideas\n"
            "Be accurate to the source. Don't add interpretation."
        ),
        "vocab": (
            f"The word or phrase to analyze: \"{text}\"\n"
            f"Context it appears in: \"{extra}\"\n\n"
            "Return:\n"
            "**DEFINITION** — Clear definition in context\n"
            "**SYNONYMS** — 5 synonyms, ordered from casual to formal\n"
            "**STRONGER ALTERNATIVES** — 3 more precise or impactful word choices\n"
            "**EXAMPLE SENTENCE** — One example using the word well\n"
            "**AVOID IF** — When NOT to use this word"
        ),
    }

    prompt = prompts.get(mode, prompts["analyze"])
    max_tok = 900 if mode in ("analyze","outline","argument","tone") else 600

    resp = client.chat.completions.create(
        model=MODEL_TEXT,
        messages=[{"role":"system","content":"You are a sharp, expert writing coach and academic tutor. Be specific, practical, and direct. Reference the actual text provided."},
                  {"role":"user","content":prompt}],
        max_tokens=max_tok, temperature=0.4
    )
    return resp.choices[0].message.content.strip()

# Persistent document context store (in-memory per session)
_doc_context_store = {}


app_flask   = Flask(__name__, static_folder=WEB_DIR)
cfg         = load_cfg()
engine      = TypingEngine()
_hk_listener    = None
_voice_listener = None
_type_state     = {"phase": "idle", "progress": 0, "status": "Ready"}
_webview_window = None
_windows_cache = []
_edge_pid    = None   # PID of the Edge process we launched
_AOT_HWND    = None   # cached window handle
_aot_thread  = None   # background watcher thread
_is_dragging = False  # pause AOT watcher during window drag to prevent crash
# Throttle apply_opacity / apply_always_on_top so rapid slider drags don't hammer win32 and crash
_last_apply_time = 0.0
_apply_throttle_sec = 0.12
# Keep ctypes callback + hook alive at module level — local vars get GC'd and silently kill the hook
_winevent_proc = None
_winevent_hook = None

# Must be ctypes.c_void_p(-1) / c_void_p(-2) — NOT raw -1/-2 integers.
# On 64-bit Python, ctypes passes raw -1 as 0x00000000FFFFFFFF (32-bit truncated),
# which Windows rejects as an invalid handle and silently fails. c_void_p forces
# the full 64-bit pointer value 0xFFFFFFFFFFFFFFFF that Windows actually expects.
HWND_TOPMOST    = ctypes.c_void_p(-1)
HWND_NOTOPMOST  = ctypes.c_void_p(-2)
SWP_NOMOVE      = 0x0002
SWP_NOSIZE      = 0x0001
SWP_NOACTIVATE  = 0x0010
SWP_FLAGS       = SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE

def _find_app_hwnd():
    """
    Find our app window. Priority order:
    1. Exact PID match (_edge_pid set when we launched Edge)
    2. Title match for pywebview window
    3. Edge/Chrome process with our port in cmdline
    """
    if not HAS_WIN32:
        return None

    result = [None]

    # 1. PID-exact match — most reliable for Edge app mode
    if _edge_pid:
        def pid_cb(hwnd, _):
            if result[0]: return
            if not ctypes.windll.user32.IsWindowVisible(hwnd): return
            try:
                _, wpid = win32process.GetWindowThreadProcessId(hwnd)
                if wpid == _edge_pid:
                    title = win32gui.GetWindowText(hwnd)
                    if not title: return
                    try:
                        rect = win32gui.GetWindowRect(hwnd)
                        if (rect[2]-rect[0]) > 200 and (rect[3]-rect[1]) > 200:
                            result[0] = hwnd
                    except Exception:
                        result[0] = hwnd
            except Exception: pass
        try: win32gui.EnumWindows(pid_cb, None)
        except Exception: pass
        if result[0]: return result[0]

    # 2. Title scan — works for pywebview (title = "Fiuxxed's AutoTyper v9.1")
    TITLE_HINTS = ("AutoTyper", "Fiuxxed")
    SKIP_TITLES = ("cmd", "command prompt", "powershell", "administrator", "conhost")
    best_score  = [0]
    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd): return
        title = win32gui.GetWindowText(hwnd)
        if not title: return
        tl = title.lower()
        if any(s in tl for s in SKIP_TITLES): return
        score = sum(2 for h in TITLE_HINTS if h in title)
        if score == 0: return
        try:
            rect = win32gui.GetWindowRect(hwnd)
            if 300 < (rect[2]-rect[0]) < 900: score += 1
        except Exception: pass
        if score > best_score[0]:
            best_score[0] = score; result[0] = hwnd
    try: win32gui.EnumWindows(cb, None)
    except Exception: pass
    if result[0]: return result[0]

    # 3. Process scan — Edge/Chrome with our port in cmdline
    try:
        for proc in psutil.process_iter(["pid","name","cmdline"]):
            name = (proc.info.get("name") or "").lower()
            if "edge" not in name and "chrome" not in name: continue
            cmdline = " ".join(proc.info.get("cmdline") or [])
            if "7890" not in cmdline: continue
            pid = proc.info["pid"]
            def _pcb(hwnd, _pid):
                if result[0]: return
                if not ctypes.windll.user32.IsWindowVisible(hwnd): return
                try:
                    _, wpid = win32process.GetWindowThreadProcessId(hwnd)
                    if wpid == _pid and win32gui.GetWindowText(hwnd): result[0] = hwnd
                except Exception: pass
            win32gui.EnumWindows(_pcb, pid)
            if result[0]: break
    except Exception: pass

    return result[0]

def _set_hwnd_topmost(hwnd, on_top):
    """Use raw win32 SetWindowPos to pin/unpin a window."""
    try:
        insert = HWND_TOPMOST if on_top else HWND_NOTOPMOST
        ctypes.windll.user32.SetWindowPos(hwnd, insert, 0, 0, 0, 0, SWP_FLAGS)
    except Exception:
        pass

def apply_always_on_top(val):
    """Called whenever the setting changes — immediately applies it via native webview property."""
    global _last_apply_time
    now = time.time()
    if now - _last_apply_time < _apply_throttle_sec:
        return
    _last_apply_time = now
    on_top = bool(val)
    if _webview_window:
        try:
            _webview_window.on_top = on_top
        except Exception: pass

def apply_opacity(val):
    """Set window opacity (0-100) using win32 layered window.
    Always keeps WS_EX_LAYERED applied to prevent Chromium dropdown rendering bug.
    Throttled so rapid slider drags don't hammer win32 and crash."""
    global _last_apply_time
    if not HAS_WIN32: return
    now = time.time()
    if now - _last_apply_time < _apply_throttle_sec:
        return
    _last_apply_time = now
    hwnd = _AOT_HWND or _find_app_hwnd()
    if not hwnd: return
    try:
        GWL_EXSTYLE   = -20
        WS_EX_LAYERED = 0x00080000
        LWA_ALPHA     = 0x00000002
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style | WS_EX_LAYERED)
        alpha = int(max(0, min(100, val)) / 100 * 255)
        ctypes.windll.user32.SetLayeredWindowAttributes(hwnd, 0, alpha, LWA_ALPHA)
    except Exception: pass

def _trigger_start():
    if _type_state["phase"] in ("idle","done"):
        _type_state["phase"] = "hotkey_trigger"

def _toggle_pause():
    if _type_state["phase"] == "typing":
        engine.pause(); _type_state["phase"] = "paused"; _type_state["status"] = "Paused"
    elif _type_state["phase"] == "paused":
        engine.resume(); _type_state["phase"] = "typing"; _type_state["status"] = "Typing..."

def _do_stop():
    engine.stop(); _type_state.update({"phase":"idle","progress":0,"status":"Stopped."})

def _trigger_clipboard():
    """F11 — read clipboard and autotype it immediately."""
    try:
        import subprocess
        # Use PowerShell to get clipboard on Windows
        result = subprocess.run(
            ["powershell", "-command", "Get-Clipboard"],
            capture_output=True, text=True, timeout=3
        )
        text = result.stdout.strip()
        if text:
            _type_state["queued_text"] = text
            _type_state["phase"] = "clipboard_trigger"
    except Exception as e:
        _log(f"_trigger_clipboard failed: {e}")

def _do_autotype():
    """Triggered by voice command — same as pressing the hotkey start."""
    _type_state["phase"] = "hotkey_trigger"

def start_hotkeys():
    global _hk_listener
    if not HAS_PYNPUT: return
    try:
        if _hk_listener: _hk_listener.stop()
    except Exception: pass
    km = {f"f{i}": getattr(kb.Key, f"f{i}") for i in range(1, 13)}
    sk  = km.get(cfg.get("hotkey_start","F8").lower())
    pk  = km.get(cfg.get("hotkey_pause","F9").lower())
    stk = km.get(cfg.get("hotkey_stop","F10").lower())
    cbk = km.get("f11")  # F11 = clipboard autotype, always fixed
    def _press(key):
        if sk  and key == sk:  _trigger_start()
        if pk  and key == pk:  _toggle_pause()
        if stk and key == stk: _do_stop()
        if cbk and key == cbk: _trigger_clipboard()
    _hk_listener = kb.Listener(on_press=_press)
    _hk_listener.daemon = True
    _hk_listener.start()

@app_flask.route("/")
def index(): return send_from_directory(WEB_DIR, "index.html")

@app_flask.route("/<path:filename>")
def static_files(filename): return send_from_directory(WEB_DIR, filename)

@app_flask.route("/api/settings", methods=["GET","POST"])
def api_settings():
    global cfg
    if request.method == "POST":
        data = request.json or {}
        cfg.update(data); save_cfg(cfg); start_hotkeys()
        apply_always_on_top(cfg.get("always_on_top", True))
        apply_opacity(cfg.get("opacity", 100))
        return jsonify({"ok": True})
    return jsonify(cfg)

@app_flask.route("/api/settings/save_key", methods=["POST"])
def api_save_key():
    global cfg
    data = request.json or {}
    if "groq_api_key" in data:
        cfg["groq_api_key"] = data["groq_api_key"].strip()
    if "gemini_api_key" in data:
        cfg["gemini_api_key"] = data["gemini_api_key"].strip()
    save_cfg(cfg)
    return jsonify({"ok": True})

@app_flask.route("/api/autotype/start", methods=["POST"])
def api_autotype_start():
    data = request.json or {}
    text = data.get("text","").strip()
    if not text: return jsonify({"error":"No text provided"})
    cfg.update({k: v for k, v in data.items() if k in DEFAULTS})
    total_reps  = max(1, int(cfg.get("repeat_count", 1)))
    repeat_delay = float(cfg.get("repeat_delay", 2.0))
    rep_state = {"current": 1}  # mutable box

    _type_state.update({"phase":"typing","progress":0,"status":"Typing..."})

    def on_prog(idx, total, pct):
        rep_tag = f"  [{rep_state['current']}/{total_reps}]" if total_reps > 1 else ""
        _type_state.update({"progress":pct,"status":f"Typing… {pct}%  ({idx}/{total} chars){rep_tag}"})

    def on_done():
        if rep_state["current"] >= total_reps or engine._stop.is_set():
            _type_state.update({"phase":"done","progress":100,"status":"✓ All done!"})
            return
        # More repeats to go — wait then re-fire
        rep_state["current"] += 1
        def _delayed():
            for _ in range(int(repeat_delay * 20)):   # 50ms ticks
                if engine._stop.is_set(): return
                time.sleep(0.05)
            if engine._stop.is_set(): return
            _type_state.update({"phase":"typing","progress":0,
                "status":f"Repeating… ({rep_state['current']}/{total_reps})"})
            engine.start(text, cfg, on_prog, on_done, on_err)
        threading.Thread(target=_delayed, daemon=True).start()

    def on_err(e): _type_state.update({"phase":"idle","status":f"Error: {e}"})
    engine.start(text, cfg, on_prog, on_done, on_err)
    return jsonify({"ok": True})

@app_flask.route("/api/autotype/pause", methods=["POST"])
def api_pause(): _toggle_pause(); return jsonify({"phase": _type_state["phase"]})

@app_flask.route("/api/autotype/stop", methods=["POST"])
def api_stop(): _do_stop(); return jsonify({"ok": True})

@app_flask.route("/api/autotype/status")
def api_status(): return jsonify(_type_state)

@app_flask.route("/api/windows")
def api_windows():
    global _windows_cache
    if not HAS_WIN32: return jsonify({"error":"pywin32 not installed","windows":[]})

    try:
        our_pid = os.getpid()
        our_pids = {our_pid}

        # Safely get child PIDs
        try:
            import psutil as _psutil
            our_proc = _psutil.Process(our_pid)
            for child in our_proc.children(recursive=True):
                try: our_pids.add(child.pid)
                except Exception: pass
        except Exception: pass

        # Exclude edge pid if known
        if _edge_pid:
            our_pids.add(_edge_pid)

        # Build pid→name map
        pid_name = {}
        try:
            import psutil as _psutil
            for proc in _psutil.process_iter(["pid", "name"]):
                try: pid_name[proc.info["pid"]] = proc.info["name"].lower().replace(".exe","")
                except Exception: pass
        except Exception: pass

        wins = []
        def cb(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd): return
                if win32gui.IsIconic(hwnd): return
                title = win32gui.GetWindowText(hwnd)
                if not title or len(title) < 2: return
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    if pid in our_pids: return
                    name = pid_name.get(pid, "unknown")
                except Exception:
                    name = "unknown"
                is_browser = any(b in name for b in BROWSERS)
                wins.append({"hwnd":hwnd,"title":title,"name":name,"is_browser":is_browser})
            except Exception: pass

        win32gui.EnumWindows(cb, None)
        wins.sort(key=lambda x: (1 if x["is_browser"] else 0, x["title"].lower()))
        if wins: wins[0]["suggested"] = True
        _windows_cache = wins
        return jsonify({"windows": wins})
    except Exception as e:
        _log_exc("api_windows", e)
        return jsonify({"error": str(e), "windows": []})

@app_flask.route("/api/screenshot", methods=["POST"])
def api_screenshot():
    data = request.json or {}
    hwnd = data.get("hwnd"); mode = data.get("mode","scan")
    strictness = data.get("strictness", cfg.get("scanner_wrong_answer_strictness","flag_all"))
    extra_context = data.get("extra_context", [])
    extra_images = data.get("extra_images", [])   # list of base64 strings from user uploads
    region = data.get("region")  # {x,y,w,h} as fraction 0-1 of screenshot, for region crop
    try:
        img = capture_window(hwnd)
        # Apply region crop if specified (fractions of full image)
        if region:
            iw, ih = img.size
            x1 = int(region["x"] * iw); y1 = int(region["y"] * ih)
            x2 = int((region["x"]+region["w"]) * iw); y2 = int((region["y"]+region["h"]) * ih)
            x1,y1 = max(0,x1), max(0,y1); x2,y2 = min(iw,x2), min(ih,y2)
            if x2-x1 > 10 and y2-y1 > 10:
                img = img.crop((x1, y1, x2, y2))
        b64 = img_to_b64(img)
        if mode == "math":
            problems = ai_math(b64, examine_examples=cfg.get("examine_examples", False), extra_context=extra_context, extra_images=extra_images)
            for p in problems:
                eq = p.get("graph_eq")
                diag = p.get("diagram_data")
                if eq and cfg.get("math_show_graphs", True):
                    p["graph_b64"] = make_graph(eq)
                elif diag and cfg.get("math_show_graphs", True):
                    p["graph_b64"] = make_diagram(diag)
                else:
                    p["graph_b64"] = None
                if "diagram_description" not in p:
                    p["diagram_description"] = None
                if cfg.get("math_formula_library", True) and p.get("answer"):
                    add_to_formula_lib(p.get("problem",""), p.get("answer",""), p.get("steps",[]))
            add_to_history("math", problems)
            return jsonify({"ok":True,"result":{"problems":problems,"_b64":b64}})
        else:
            questions = ai_scan(b64, strictness, examine_examples=cfg.get("examine_examples", False), extra_context=extra_context, extra_images=extra_images)
            add_to_history("scan", questions)
            return jsonify({"ok":True,"result":{"questions":questions}})
    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()})

def _file_b64_to_images(b64_str, filename=""):
    """Convert a base64 file into a list of base64 PNG images (one per page for PDFs, one for images)."""
    raw = base64.b64decode(b64_str)
    fname = (filename or "").lower()
    is_pdf = fname.endswith(".pdf") or raw[:4] == b"%PDF"
    if is_pdf and HAS_FITZ:
        pages = []
        doc = _fitz.open(stream=raw, filetype="pdf")
        for page in doc:
            mat = _fitz.Matrix(2.0, 2.0)  # 2× zoom for readability
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            buf = io.BytesIO(); img.save(buf, format="PNG"); buf.seek(0)
            pages.append(base64.b64encode(buf.getvalue()).decode("utf-8"))
        doc.close()
        return pages, True  # True = is_pdf
    elif is_pdf and not HAS_FITZ:
        # Fallback: send raw PDF b64 as a single image and hope vision model handles it
        return [b64_str], True
    else:
        # It's an image — just return as-is
        return [b64_str], False

@app_flask.route("/api/scan_files", methods=["POST"])
def api_scan_files():
    """Scan attached files only (no screenshot). Handles images and PDFs page by page."""
    data = request.json or {}
    mode = data.get("mode", "scan")
    extra_context = data.get("extra_context", [])
    attached = data.get("attached_files", [])  # list of {b64, name, kind}
    strictness = data.get("strictness", cfg.get("scanner_wrong_answer_strictness","flag_all"))
    try:
        all_results = []
        for f in attached:
            b64 = f.get("b64") or f if isinstance(f, str) else None
            if not b64: continue
            name = f.get("name", "") if isinstance(f, dict) else ""
            pages, is_pdf = _file_b64_to_images(b64, name)
            for pi, page_b64 in enumerate(pages):
                page_label = f" (page {pi+1}/{len(pages)})" if is_pdf and len(pages) > 1 else ""
                ctx = list(extra_context)
                if page_label:
                    ctx = [f"[Scanning: {name}{page_label}]"] + ctx
                if mode == "math":
                    results = ai_math(page_b64, examine_examples=cfg.get("examine_examples", False), extra_context=ctx)
                    for p in results:
                        p["_source"] = name + page_label
                        eq = p.get("graph_eq"); diag = p.get("diagram_data")
                        if eq and cfg.get("math_show_graphs", True): p["graph_b64"] = make_graph(eq)
                        elif diag and cfg.get("math_show_graphs", True): p["graph_b64"] = make_diagram(diag)
                        else: p["graph_b64"] = None
                        if "diagram_description" not in p: p["diagram_description"] = None
                        if cfg.get("math_formula_library", True) and p.get("answer"):
                            add_to_formula_lib(p.get("problem",""), p.get("answer",""), p.get("steps",[]))
                    all_results.extend(results)
                else:
                    results = ai_scan(page_b64, strictness, examine_examples=cfg.get("examine_examples", False), extra_context=ctx)
                    for q2 in results:
                        q2["_source"] = name + page_label
                    all_results.extend(results)
        if mode == "math":
            add_to_history("math", all_results)
            # Include last page b64 for recheck feature
            last_b64 = pages[-1] if pages else ""
            return jsonify({"ok": True, "result": {"problems": all_results, "_b64": last_b64}})
        else:
            add_to_history("scan", all_results)
            return jsonify({"ok": True, "result": {"questions": all_results}})
    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()})

@app_flask.route("/api/screenshot/preview", methods=["POST"])
def api_screenshot_preview():
    """Take a screenshot and return full-res b64 for region selector."""
    data = request.json or {}
    hwnd = data.get("hwnd")
    result = [None]; error = [None]
    def _capture():
        try:
            our_hwnd = _hide_app_window()
            try:
                img = capture_window(hwnd, bring_to_front=False)
                result[0] = img
            finally:
                _show_app_window(our_hwnd)
        except Exception as e:
            error[0] = str(e)
    t = threading.Thread(target=_capture, daemon=True)
    t.start(); t.join(timeout=10)  # 10s hard cap — never hangs the UI
    if t.is_alive():
        return jsonify({"error": "Screenshot timed out — make sure the window is not minimized."})
    if error[0]:
        return jsonify({"error": error[0]})
    try:
        ow, oh = result[0].size
        b64 = img_to_b64(result[0])
        return jsonify({"ok":True,"b64":b64,"w":ow,"h":oh})
    except Exception as e:
        return jsonify({"error": str(e)})

@app_flask.route("/api/double_check", methods=["POST"])
def api_double_check():
    data = request.json or {}
    q = data.get("question","").strip()
    if not q: return jsonify({"error":"No question"})
    try: return jsonify({"answer": ai_double_check(q)})
    except Exception as e: return jsonify({"error": str(e)})

@app_flask.route("/api/qa", methods=["POST"])
def api_qa_route():
    data = request.json or {}
    question = data.get("question","").strip()
    history = data.get("history",[]); is_followup = bool(data.get("followup", False))
    extra_context = data.get("extra_context", [])
    personality = data.get("personality", "default")
    if not question: return jsonify({"error":"No question"})
    extra_images = data.get("extra_images", [])
    try:
        answer, new_history = ai_qa(question, history, is_followup, extra_context, extra_images, personality)
        return jsonify({"answer":answer,"history":new_history})
    except Exception as e: return jsonify({"error":str(e)})

@app_flask.route("/api/humanize", methods=["POST"])
def api_humanize():
    data   = request.json or {}
    text   = data.get("text","").strip()
    action = data.get("action","humanize")
    step   = max(1, min(int(data.get("step", 1)), 10))
    hlevel = max(1, min(int(data.get("humanize_level", 1)), 10))
    if not text: return jsonify({"error":"No text"})
    client = get_client()

    word_count = len(text.split())

    if action == "longer":
        add    = step * 12
        tokens = word_count + add + 30   # tight ceiling: current + add + small buffer
        p = (
            f"Add exactly about {add} words to this text — not more. "
            "Keep the same casual student tone. Simple words, short sentences. "
            "Just add a little more, do NOT write a full essay. "
            "Return only the expanded text, nothing else.\n\n"
            f"Text: {text}\n\nExpanded:"
        )

    elif action == "shorter":
        remove = step * 12
        target = max(6, word_count - remove)
        tokens = target + 20
        p = (
            f"Shorten this to about {target} words (it is currently {word_count} words). "
            "Cut the least important parts. Do NOT add anything new. "
            "Return only the shortened text, nothing else.\n\n"
            f"Text: {text}\n\nShortened:"
        )

    else:
        tokens = word_count + 40
        if hlevel <= 2:
            tone = "like a regular student — casual, simple, slightly imperfect."
        elif hlevel <= 4:
            tone = "like a teenager, very casual, use words like kinda, basically, like, tbh."
        elif hlevel <= 6:
            tone = "like a kid who barely paid attention. Super casual, filler words, imperfect grammar."
        elif hlevel <= 8:
            tone = "like a little kid. Very simple words, short sentences, sounds confused but right."
        else:
            tone = "like a toddler. Extremely simple. Very short. Maybe repeat words. Cute and dumb but the idea is still there."
        banned = "furthermore, notably, demonstrates, pivotal, crucial, significant, utilize, facilitate, encompasses, nuanced, delve, realm, leverage, foster, underscores, illuminates, exemplifies"
        p = (
            f"Rewrite this so it sounds {tone} "
            f"NEVER use: {banned}. Same length roughly. "
            "Return only the rewritten text, nothing else.\n\n"
            f"Text: {text}\n\nRewritten:"
        )

    try:
        resp = client.chat.completions.create(
            model=MODEL_TEXT,
            messages=[{"role":"user","content":p}],
            max_tokens=min(tokens, 600), temperature=0.75
        )
        return jsonify({"result": resp.choices[0].message.content.strip()})
    except Exception as e:
        return jsonify({"error": str(e)})

@app_flask.route("/api/write", methods=["POST"])
def api_write():
    data   = request.json or {}
    mode   = data.get("mode","analyze")
    text   = data.get("text","").strip()
    extra  = data.get("extra","").strip()
    doc_ctx= data.get("doc_context","").strip()
    if not text: return jsonify({"error":"No text provided"})
    try:
        result = ai_write(mode, text, extra, doc_ctx)
        return jsonify({"ok":True,"result":result})
    except Exception as e:
        return jsonify({"error":str(e)})

@app_flask.route("/api/doc_context", methods=["POST"])
def api_doc_context():
    data = request.json or {}
    action = data.get("action","set")
    session_id = data.get("session","default")
    if action == "set":
        _doc_context_store[session_id] = data.get("text","").strip()
        return jsonify({"ok":True,"chars":len(_doc_context_store[session_id])})
    elif action == "get":
        return jsonify({"ok":True,"text":_doc_context_store.get(session_id,"")})
    elif action == "clear":
        _doc_context_store.pop(session_id, None)
        return jsonify({"ok":True})
    return jsonify({"error":"Unknown action"})

@app_flask.route("/api/formula_library", methods=["GET"])
def api_formula_library(): return jsonify({"formulas": load_formula_lib()})

@app_flask.route("/api/formula_library/clear", methods=["POST"])
def api_formula_library_clear(): save_formula_lib([]); return jsonify({"ok":True})

# ── Math Examples Library ─────────────────────────────────────────────
EXAMPLES_LIB_FILE = os.path.join(BASE, "math_examples.json")

def load_examples_lib():
    try:
        with open(EXAMPLES_LIB_FILE) as f: return json.load(f)
    except Exception: return []

def save_examples_lib(lib):
    try:
        with open(EXAMPLES_LIB_FILE, "w") as f: json.dump(lib, f, indent=2)
    except Exception: pass

def _examples_context() -> str:
    """Inject saved examples as context into math solving prompts."""
    try:
        lib = load_examples_lib()
        if not lib: return ""
        lines = []
        for ex in lib[-10:]:  # last 10 examples
            lines.append(
                f"• TYPE: {ex.get('topic','Unknown')}\n"
                f"  PROBLEM: {ex.get('problem_example','')}\n"
                f"  METHOD: {ex.get('method','')}\n"
                f"  ANSWER FORMAT: {ex.get('answer_format','')}"
            )
        if lines:
            return (
                "\n\n[EXAMPLE PROBLEMS — you've seen these types before, use the same method]:\n"
                + "\n\n".join(lines) + "\n"
            )
        return ""
    except Exception: return ""

@app_flask.route("/api/math_recheck", methods=["POST"])
def api_math_recheck():
    """
    Recheck a single math problem answer with fresh eyes.
    Uses the explicit geometry reasoning chain to verify or correct.
    """
    data = request.json or {}
    problem_label = data.get("problem_label", "?")
    problem_text  = data.get("problem_text", "")
    current_answer = data.get("current_answer", "")
    problem_type  = data.get("problem_type", "")
    b64           = data.get("b64", "")

    if not b64:
        return jsonify({"error": "No image — rescan the worksheet first"})

    recheck_prompt = (
        f"I solved Problem {problem_label} from this worksheet and got: '{current_answer}'\n\n"
        f"Problem text: {problem_text}\n\n"
        "Please VERIFY this answer by doing EXACTLY this reasoning:\n\n"
        f"STEP 1 — Find problem {problem_label} in the image\n\n"
        "STEP 2 — Identify AUTO-GIVEN parts (before looking at any drawn marks):\n"
        "  • SHARED SIDE: is there a line segment that forms part of BOTH triangles? → Reflexive Property, equal\n"
        "  • VERTICAL ANGLES: do triangles share a vertex where lines cross (X or bowtie shape)? → vertical angles equal\n"
        "  • RIGHT ANGLES: is there a right angle box (□) or perpendicular mark? → both sides of that point = 90°\n\n"
        "STEP 3 — Read EXPLICIT marks:\n"
        "  • Tick marks on sides (single | = one pair, double || = another pair)\n"
        "  • Arc marks on angles (single arc = one pair, double arc = another pair)\n"
        "  Write: 'I see tick on [segment], arc at [angle]...'\n\n"
        "STEP 4 — Count total: ___ angle pairs, ___ side pairs\n\n"
        "STEP 5 — Apply theorem test:\n"
        "  • Is the marked side BETWEEN the two marked angles? → ASA\n"
        "  • Is the marked side OUTSIDE the two marked angles? → AAS\n"
        "  • Only sides, no angles completing a theorem? → check SSS/SAS\n"
        "  • SSA (2 sides + non-included angle) = NEITHER\n"
        "  • Not enough for any theorem = NEITHER\n\n"
        "STEP 6 — Is my answer correct?\n\n"
        "USE PLAIN TEXT ONLY — no LaTeX, no $, no backslashes. Use ∠ ≅ △ directly.\n\n"
        "Return ONLY valid JSON (one object, not array):\n"
        '{"verdict":"correct","correct_answer":"B) AAS","reasoning":"Arc at A and B (2 angles), tick on AC which is non-included side, so AAS"}\n'
        "verdict must be 'correct', 'wrong', or 'unclear'\n"
        "If correct, correct_answer = same as my answer\n"
        "If wrong, correct_answer = the right answer\n"
        "reasoning = 1-2 sentences explaining exactly what marks prove"
    )

    try:
        raw = ""
        try:
            raw = _gemini_vision(b64, recheck_prompt)
            _log(f"Recheck {problem_label}: Gemini 2.0 Flash")
        except Exception as gem_err:
            _log(f"Gemini recheck failed ({gem_err}), trying Groq")
            client = get_client()
            resp = client.chat.completions.create(
                model=MODEL_VISION,
                messages=[{"role":"user","content":[
                    {"type":"image_url","image_url":{"url":f"data:image/png;base64,{b64}"}},
                    {"type":"text","text":recheck_prompt}
                ]}],
                max_tokens=400, temperature=0.05
            )
            raw = resp.choices[0].message.content.strip()

        raw_clean = re.sub(r'```(?:json)?\s*','',raw).strip().strip('`')
        m = re.search(r'\{.*\}', raw_clean, re.DOTALL)
        if not m:
            return jsonify({"verdict":"unclear","reasoning":"Could not parse recheck response","correct_answer":current_answer})

        result = json.loads(m.group(0))
        # Strip any LaTeX from the result
        for k,v in result.items():
            if isinstance(v, str): result[k] = _strip_latex(v)

        _log(f"Recheck result for {problem_label}: {result.get('verdict')} — {result.get('correct_answer')}")
        return jsonify(result)

    except Exception as e:
        _log_exc("api_math_recheck", e)
        return jsonify({"error": str(e)})
def api_math_examples():
    return jsonify({"examples": load_examples_lib()})

@app_flask.route("/api/math_examples/delete", methods=["POST"])
def api_math_examples_delete():
    data = request.json or {}
    idx = data.get("index")
    lib = load_examples_lib()
    if idx is not None and 0 <= idx < len(lib):
        lib.pop(idx)
        save_examples_lib(lib)
    return jsonify({"ok": True, "count": len(lib)})

@app_flask.route("/api/math_examples/clear", methods=["POST"])
def api_math_examples_clear():
    save_examples_lib([])
    return jsonify({"ok": True})

@app_flask.route("/api/math_examples/scan", methods=["POST"])
def api_math_examples_scan():
    """
    Screenshot a window/region, extract the math METHOD (not just the answer),
    and save it to the examples library so it gets injected into future solves.
    """
    data = request.json or {}
    hwnd = data.get("hwnd")
    b64 = data.get("b64")  # can pass raw b64 too (region selector)

    if not b64:
        # Take screenshot of the window
        if not HAS_MSS or not HAS_WIN32:
            return jsonify({"error": "Screenshot not available — install mss and pywin32"})
        try:
            rect = win32gui.GetWindowRect(hwnd)
            x, y, x2, y2 = rect
            w, h = x2-x, y2-y
            if w < 10 or h < 10:
                return jsonify({"error": "Window too small or minimized"})
            with mss.mss() as sct:
                mon = {"left": x, "top": y, "width": w, "height": h}
                shot = sct.grab(mon)
                img = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                b64 = base64.b64encode(buf.getvalue()).decode()
        except Exception as e:
            return jsonify({"error": f"Screenshot failed: {e}"})

    # Pre-process image
    b64_processed = _preprocess_math_image(b64)

    # Ask the AI to extract the METHOD — not solve, but teach
    extract_prompt = (
        "You are looking at a math worksheet or textbook page with worked examples.\n\n"
        "Your job is to EXTRACT THE TEACHING METHOD from what you see — not to solve, but to understand:\n"
        "1. What TYPE of math problem is this? (e.g. 'Triangle Congruence AAS', 'Quadratic equations', 'Systems of equations')\n"
        "2. What is a representative EXAMPLE PROBLEM shown? (copy it verbatim)\n"
        "3. What is the STEP-BY-STEP METHOD used to solve it? (general process, not just this one answer)\n"
        "4. What does the ANSWER look like? (format, units, notation used)\n"
        "5. What KEY RULES or THEOREMS are shown? (e.g. 'AAS: 2 angles + non-included side = congruent triangles')\n\n"
        "USE PLAIN TEXT ONLY — no LaTeX, no $...$, no \\angle. Use ∠ ≅ △ directly.\n\n"
        "Return ONLY valid JSON (no markdown):\n"
        '{"topic":"Triangle Congruence","problem_example":"What additional info proves △ABC ≅ △DEF by AAS?","method":"1. Read triangle names to get correspondence A↔D B↔E C↔F. 2. Catalog what is already marked in diagram (tick marks = given sides, arc marks = given angles). 3. Identify what AAS needs: 2 angle pairs + 1 non-included side. 4. Find what is missing. 5. Match missing piece to answer choice. 6. Reject any choice already shown in diagram.","answer_format":"Letter) description e.g. C) ∠U ≅ ∠L","key_rules":["AAS: 2 angles + non-included side","Already-marked parts cannot be the additional info","Vertex correspondence from triangle name tells you the pairs"]}'
    )

    try:
        raw = ""
        # Try Gemini first
        try:
            raw = _gemini_vision(b64_processed, extract_prompt)
            _log("Example scan: Gemini 2.0 Flash")
        except Exception as gem_err:
            _log(f"Gemini example scan failed ({gem_err}), falling back")
            client = get_client()
            resp = client.chat.completions.create(
                model=MODEL_VISION,
                messages=[{"role":"user","content":[
                    {"type":"image_url","image_url":{"url":f"data:image/png;base64,{b64_processed}"}},
                    {"type":"text","text":extract_prompt}
                ]}],
                max_tokens=1000, temperature=0.1
            )
            raw = resp.choices[0].message.content.strip()

        # Parse the extracted example
        raw_clean = re.sub(r'```(?:json)?\s*','',raw).strip().strip('`')
        m = re.search(r'\{.*\}', raw_clean, re.DOTALL)
        if not m:
            return jsonify({"error": "Could not extract example structure from image"})

        example = json.loads(m.group(0))
        example["saved_at"] = time.strftime("%Y-%m-%d %H:%M")

        # Save to examples library
        lib = load_examples_lib()
        # Don't save duplicate topics
        if not any(e.get("topic","").lower() == example.get("topic","").lower() for e in lib):
            lib.append(example)
            save_examples_lib(lib)
            return jsonify({"ok": True, "example": example, "total": len(lib), "duplicate": False})
        else:
            # Update existing
            for i, e in enumerate(lib):
                if e.get("topic","").lower() == example.get("topic","").lower():
                    lib[i] = example
                    break
            save_examples_lib(lib)
            return jsonify({"ok": True, "example": example, "total": len(lib), "duplicate": True})

    except Exception as e:
        _log_exc("api_math_examples_scan", e)
        return jsonify({"error": str(e)})

@app_flask.route("/api/history", methods=["GET"])
def api_history(): return jsonify({"history": load_history()})

@app_flask.route("/api/history/clear", methods=["POST"])
def api_history_clear(): save_history_file([]); return jsonify({"ok":True})

# Local Music
AUDIO_EXTS = {".mp3",".flac",".wav",".ogg",".m4a",".aac",".wma",".opus",".aiff",".ape",".alac"}
VIDEO_EXTS = {".mp4",".mkv",".avi",".mov",".wmv",".webm",".flv",".m4v",".ts",".3gp",".mpg",".mpeg"}
IMAGE_EXTS = {".jpg",".jpeg",".png",".gif",".bmp",".webp",".tiff",".tif",".svg",".ico",".heic",".avif"}
DOC_EXTS   = {".pdf",".txt",".doc",".docx",".xls",".xlsx",".ppt",".pptx",".csv",".md",".rtf",".odt",".epub"}
ALL_MEDIA_EXTS = AUDIO_EXTS | VIDEO_EXTS | IMAGE_EXTS | DOC_EXTS

def _file_kind(ext):
    if ext in AUDIO_EXTS: return "audio"
    if ext in VIDEO_EXTS: return "video"
    if ext in IMAGE_EXTS: return "image"
    if ext in DOC_EXTS:   return "doc"
    return "other"

def _get_music_roots():
    roots = []
    home = os.path.expanduser("~")
    for c in [os.path.join(home,"Music"),os.path.join(home,"Videos"),os.path.join(home,"Pictures"),
              os.path.join(home,"Documents"),os.path.join(home,"Downloads"),os.path.join(home,"Desktop"),
              r"C:\Users\Public\Music",r"C:\Users\Public\Videos"]:
        if os.path.isdir(c): roots.append(c)
    return roots

def _scan_music_dir(directory, max_files=2000):
    results = []
    try:
        for root, dirs, files in os.walk(directory):
            dirs[:] = [d for d in dirs if not d.startswith(".")]
            for fname in files:
                ext = os.path.splitext(fname)[1].lower()
                if ext in ALL_MEDIA_EXTS:
                    full = os.path.join(root, fname)
                    try: size = os.path.getsize(full)
                    except: size = 0
                    try: mtime = os.path.getmtime(full)
                    except: mtime = 0
                    results.append({
                        "path":  full,
                        "name":  os.path.splitext(fname)[0],
                        "ext":   ext.lstrip("."),
                        "kind":  _file_kind(ext),
                        "size":  size,
                        "mtime": mtime,
                    })
                    if len(results) >= max_files: return results
    except Exception: pass
    return results

@app_flask.route("/api/music/roots")
def api_music_roots():
    return jsonify({"roots": _get_music_roots()})

HAS_YTDLP = False
try:
    import yt_dlp
    HAS_YTDLP = True
except ImportError:
    pass

# Cache for YouTube search pagination: one query's full result set, keyed by query
_music_search_cache_query = None
_music_search_cache_results = None

@app_flask.route("/api/music/search")
def api_music_search():
    """Search YouTube via yt-dlp (reliable) or Invidious fallback. Supports ?page=1,2,3... (15 per page)."""
    global _music_search_cache_query, _music_search_cache_results
    query = request.args.get("q", "").strip()
    page = max(1, int(request.args.get("page", 1)))
    if not query:
        return jsonify({"error": "No query", "results": [], "page": 1, "has_prev": False, "has_next": False})

    PAGE_SIZE = 15
    FETCH_COUNT = 500  # fetch up to 500 results (~33 pages)
    start = (page - 1) * PAGE_SIZE

    def return_paginated(full_list):
        nonlocal page, start
        total = len(full_list)
        chunk = full_list[start:start + PAGE_SIZE]
        return jsonify({
            "results": chunk,
            "page": page,
            "has_prev": page > 1,
            "has_next": start + PAGE_SIZE < total,
            "total_cached": total,
        })

    # ── Primary: yt-dlp ──────────────────────────────────────
    if HAS_YTDLP:
        try:
            # Use cache if same query and we have cached results
            if query == _music_search_cache_query and _music_search_cache_results is not None:
                return return_paginated(_music_search_cache_results)

            # Cache miss — fetch up to FETCH_COUNT results
            ydl_opts = {
                "quiet": True,
                "no_warnings": True,
                "extract_flat": "in_playlist",
                "default_search": f"ytsearch{FETCH_COUNT}",
                "skip_download": True,
                "ignoreerrors": True,
                "playlistend": FETCH_COUNT,
            }
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(f"ytsearch{FETCH_COUNT}:{query}", download=False)
            entries = (info or {}).get("entries") or []
            all_results = []
            for v in entries:
                if not v:
                    continue
                vid = v.get("id") or v.get("url","")
                thumb = v.get("thumbnail","")
                if not thumb and vid:
                    thumb = f"https://i.ytimg.com/vi/{vid}/mqdefault.jpg"
                all_results.append({
                    "videoId":  vid,
                    "title":    v.get("title","Untitled"),
                    "author":   v.get("uploader") or v.get("channel",""),
                    "duration": v.get("duration") or 0,
                    "thumb":    thumb,
                })
            _music_search_cache_query = query
            _music_search_cache_results = all_results
            return return_paginated(all_results)
        except Exception:
            pass  # fall through to Invidious

    # ── Fallback: Invidious ──────────────────────────────────
    import urllib.request, urllib.parse, ssl
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    instances = [
        "https://inv.nadeko.net", "https://invidious.privacyredirect.com",
        "https://yt.artemislena.eu", "https://invidious.nerdvpn.de",
        "https://iv.melmac.space", "https://invidious.fdn.fr",
    ]
    for base in instances:
        try:
            # Invidious may support page param; try it
            url = (f"{base}/api/v1/search?q={urllib.parse.quote(query)}"
                   f"&type=video&fields=videoId,title,author,lengthSeconds,videoThumbnails&page={page}")
            req = urllib.request.Request(url, headers={"User-Agent":"Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=5, context=ctx) as resp:
                data = json.loads(resp.read().decode())
            if not isinstance(data, list) or not data:
                continue
            results = []
            for v in data[:PAGE_SIZE]:
                thumbs = v.get("videoThumbnails",[])
                thumb = next((t["url"] for t in thumbs if t.get("quality")=="medium"),
                             thumbs[0].get("url","") if thumbs else "")
                if thumb and thumb.startswith("/"):
                    thumb = base + thumb
                results.append({
                    "videoId": v.get("videoId",""),
                    "title":   v.get("title","Untitled"),
                    "author":  v.get("author",""),
                    "duration":v.get("lengthSeconds",0),
                    "thumb":   thumb,
                })
            if results:
                return jsonify({
                    "results": results,
                    "page": page,
                    "has_prev": page > 1,
                    "has_next": len(data) >= PAGE_SIZE,
                })
        except Exception:
            continue

    return jsonify({"error": "Search unavailable — install yt-dlp (run install.bat) or check internet.", "results": []})

@app_flask.route("/api/music/browse", methods=["POST"])
def api_music_browse():
    data = request.json or {}
    directory = data.get("directory","").strip()
    if not directory:
        return jsonify({"roots":_get_music_roots(),"files":[],"directory":""})
    if not os.path.isdir(directory):
        return jsonify({"error":f"Folder not found: {directory}","files":[]})
    files = _scan_music_dir(directory)
    files.sort(key=lambda f: f["name"].lower())
    return jsonify({"files":files,"directory":directory,"count":len(files)})

@app_flask.route("/api/music/stream")
def api_music_stream():
    from flask import send_file, Response
    import re as _re
    path = request.args.get("path","")
    if not path or not os.path.isfile(path):
        return jsonify({"error":"File not found"}), 404
    ext = os.path.splitext(path)[1].lower()
    mime_map = {
        # audio
        ".mp3":"audio/mpeg",".flac":"audio/flac",".wav":"audio/wav",
        ".ogg":"audio/ogg",".m4a":"audio/mp4",".aac":"audio/aac",
        ".wma":"audio/x-ms-wma",".opus":"audio/opus",".aiff":"audio/aiff",
        ".ape":"audio/ape",".alac":"audio/mp4",
        # video
        ".mp4":"video/mp4",".mkv":"video/x-matroska",".avi":"video/x-msvideo",
        ".mov":"video/quicktime",".wmv":"video/x-ms-wmv",".webm":"video/webm",
        ".flv":"video/x-flv",".m4v":"video/mp4",".ts":"video/mp2t",
        ".3gp":"video/3gpp",".mpg":"video/mpeg",".mpeg":"video/mpeg",
        # image
        ".jpg":"image/jpeg",".jpeg":"image/jpeg",".png":"image/png",
        ".gif":"image/gif",".bmp":"image/bmp",".webp":"image/webp",
        ".tiff":"image/tiff",".tif":"image/tiff",".svg":"image/svg+xml",
        ".ico":"image/x-icon",".heic":"image/heic",".avif":"image/avif",
        # documents
        ".pdf":"application/pdf",".txt":"text/plain",".md":"text/plain",
        ".csv":"text/csv",".rtf":"application/rtf",
        ".doc":"application/msword",
        ".docx":"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xls":"application/vnd.ms-excel",
        ".xlsx":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".ppt":"application/vnd.ms-powerpoint",
        ".pptx":"application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".epub":"application/epub+zip",".odt":"application/vnd.oasis.opendocument.text",
    }
    mime = mime_map.get(ext, "application/octet-stream")
    # Images and docs — just serve directly, no range needed
    if ext in IMAGE_EXTS or ext in DOC_EXTS:
        return send_file(path, mimetype=mime, conditional=True)
    # Audio and video — support range requests for seeking
    file_size = os.path.getsize(path)
    rh = request.headers.get("Range")
    if rh:
        m = _re.search(r"bytes=(\d+)-(\d*)", rh)
        start = int(m.group(1)) if m else 0
        end   = int(m.group(2)) if m and m.group(2) else file_size - 1
        end   = min(end, file_size - 1)
        length = end - start + 1
        def stream_range():
            with open(path,"rb") as f:
                f.seek(start); rem = length
                while rem > 0:
                    chunk = f.read(min(65536, rem))
                    if not chunk: break
                    rem -= len(chunk); yield chunk
        resp = Response(stream_range(), 206, mimetype=mime, direct_passthrough=True)
        resp.headers["Content-Range"]  = f"bytes {start}-{end}/{file_size}"
        resp.headers["Accept-Ranges"]  = "bytes"
        resp.headers["Content-Length"] = str(length)
        return resp
    return send_file(path, mimetype=mime, conditional=True)

@app_flask.route("/api/voice/devices")
def api_voice_devices():
    devices = []; default_idx = None
    if HAS_SOUNDDEVICE:
        try:
            devs = sd.query_devices(); default_in = sd.default.device[0]
            for i, d in enumerate(devs):
                if d["max_input_channels"] > 0:
                    devices.append({"index":i,"name":d["name"],"is_default":(i==default_in)})
                    if i == default_in: default_idx = i
        except Exception as e: return jsonify({"error":str(e),"devices":[]})
    elif HAS_PYAUDIO and HAS_SR:
        try:
            for i, name in enumerate(sr.Microphone.list_microphone_names()):
                devices.append({"index":i,"name":name,"is_default":(i==0)})
            default_idx = 0
        except Exception as e: return jsonify({"error":str(e),"devices":[]})
    else: return jsonify({"error":"No audio backend available","devices":[]})
    return jsonify({"devices":devices,"default":default_idx})

@app_flask.route("/api/voice/status")
def api_voice_status():
    listening = bool(_voice_listener and _voice_listener._running)
    vstat = _voice_listener.status if _voice_listener else "idle"
    backend = "sounddevice" if (HAS_SOUNDDEVICE and not HAS_PYAUDIO) else ("pyaudio" if HAS_PYAUDIO else "none")
    return jsonify({"available":HAS_VOICE,"listening":listening,"status":vstat,"backend":backend,"log":(_voice_listener.log[-30:] if _voice_listener else [])})

@app_flask.route("/api/voice/toggle", methods=["POST"])
def api_voice_toggle():
    global _voice_listener
    if not HAS_VOICE: return jsonify({"error":"No audio backend found."})
    data = request.json or {}; device_index = data.get("device_index", None)
    if _voice_listener and _voice_listener._running:
        _voice_listener.stop(); _voice_listener = None
        return jsonify({"listening":False})
    _voice_listener = VoiceListener(cfg_fn=lambda:cfg, autotype_fn=_trigger_start,
        pause_fn=_toggle_pause, stop_fn=_do_stop, device_index=device_index)
    _voice_listener.start()
    return jsonify({"listening":True,"device_index":device_index})

@app_flask.route("/api/voice/test", methods=["POST"])
def api_voice_test():
    if not HAS_VOICE: return jsonify({"error":"No audio backend available"})
    data = request.json or {}; device_index = data.get("device_index", None)
    try:
        if HAS_SOUNDDEVICE and not HAS_PYAUDIO:
            import io, wave
            RATE = 16000; SECS = 5
            frames = sd.rec(int(RATE*SECS), samplerate=RATE, channels=1, dtype="int16", device=device_index, blocking=True)
            rms = float((frames.astype("float32")**2).mean()**0.5) / 32768
            if rms < 0.008: return jsonify({"error":"No speech detected — speak louder or check mic"})
            buf = io.BytesIO()
            with wave.open(buf,"wb") as wf:
                wf.setnchannels(1); wf.setsampwidth(2); wf.setframerate(RATE); wf.writeframes(frames.tobytes())
            buf.seek(0)
            rec = sr.Recognizer()
            with sr.AudioFile(buf) as src: audio = rec.record(src)
            text = rec.recognize_google(audio, language=cfg.get("voice_language","en-US"))
            return jsonify({"heard":text})
        else:
            rec = sr.Recognizer()
            kwargs = {} if device_index is None else {"device_index":device_index}
            mic = sr.Microphone(**kwargs)
            with mic as src: rec.adjust_for_ambient_noise(src, duration=0.5)
            with mic as src: audio = rec.listen(src, timeout=5, phrase_time_limit=7)
            text = rec.recognize_google(audio, language=cfg.get("voice_language","en-US"))
            return jsonify({"heard":text})
    except Exception as e: return jsonify({"error":str(e)})


# ══════════════════════════════════════════════════════════════════════
#  WINDOW CONTROL  (frameless titlebar support)
# ══════════════════════════════════════════════════════════════════════
WM_NCLBUTTONDOWN = 0x00A1
HTCAPTION        = 0x0002
WM_SYSCOMMAND    = 0x0112
SC_MINIMIZE      = 0xF020
SC_CLOSE         = 0xF060

@app_flask.route("/api/window/drag", methods=["POST"])
def api_window_drag():
    """Release mouse capture and send WM_NCLBUTTONDOWN so Windows moves the window."""
    if not HAS_WIN32: return jsonify({"ok": False})
    hwnd = _AOT_HWND or _find_app_hwnd()
    if not hwnd: return jsonify({"ok": False, "error": "window not found"})
    try:
        ctypes.windll.user32.ReleaseCapture()
        ctypes.windll.user32.SendMessageW(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

@app_flask.route("/api/window/drag_start", methods=["POST"])
def api_drag_start():
    global _is_dragging
    _is_dragging = True
    return jsonify({"ok": True})

@app_flask.route("/api/window/drag_end", methods=["POST"])
def api_drag_end():
    global _is_dragging
    _is_dragging = False
    return jsonify({"ok": True})

@app_flask.route("/api/window/minimize", methods=["POST"])
def api_window_minimize():
    if not HAS_WIN32: return jsonify({"ok": False})
    hwnd = _AOT_HWND or _find_app_hwnd()
    if not hwnd: return jsonify({"ok": False})
    try:
        ctypes.windll.user32.ShowWindow(hwnd, 6)  # SW_MINIMIZE = 6
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

@app_flask.route("/api/window/close", methods=["POST"])
def api_window_close():
    threading.Thread(target=lambda: (time.sleep(0.15), os._exit(0)), daemon=True).start()
    return jsonify({"ok": True})

@app_flask.route("/api/window/apply_frameless", methods=["POST"])
def api_apply_frameless():
    """Re-apply frameless style + resize borders. Call after window loads."""
    hwnd = _AOT_HWND or _find_app_hwnd()
    if hwnd: _apply_frameless_to_hwnd(hwnd)
    return jsonify({"ok": bool(hwnd), "hwnd": hwnd})

# ══════════════════════════════════════════════════════════════════════
#  YOUTUBE TRANSCRIPT
# ══════════════════════════════════════════════════════════════════════
def _extract_video_id(url_or_id):
    """Extract YouTube video ID from a URL or return as-is if already an ID."""
    url_or_id = url_or_id.strip()
    m = re.search(r'(?:v=|/v/|youtu\.be/|/embed/|/shorts/)([A-Za-z0-9_-]{11})', url_or_id)
    if m:
        return m.group(1)
    if re.fullmatch(r'[A-Za-z0-9_-]{11}', url_or_id):
        return url_or_id
    return url_or_id

@app_flask.route("/api/youtube_transcript", methods=["POST"])
def api_youtube_transcript():
    try:
        from youtube_transcript_api import YouTubeTranscriptApi, NoTranscriptFound, TranscriptsDisabled
    except ImportError:
        return jsonify({"ok": False, "error": "youtube-transcript-api not installed. Run: pip install youtube-transcript-api"}), 500

    data = request.get_json(silent=True) or {}
    raw = (data.get("manual_url") or data.get("video_id") or "").strip()
    if not raw:
        return jsonify({"ok": False, "error": "No URL or video ID provided"}), 400

    video_id = _extract_video_id(raw)
    try:
        # v1.x API: instantiate, then call .fetch()
        ytt = YouTubeTranscriptApi()
        fetched = ytt.fetch(video_id, languages=["en", "en-US", "en-GB"])
        text = " ".join(snippet.text for snippet in fetched).strip()
        return jsonify({"ok": True, "transcript": text})
    except NoTranscriptFound:
        # No English — try fetching whatever language is available
        try:
            ytt = YouTubeTranscriptApi()
            fetched = ytt.fetch(video_id)
            text = " ".join(snippet.text for snippet in fetched).strip()
            return jsonify({"ok": True, "transcript": text})
        except Exception as e2:
            _log_exc("youtube_transcript fallback", e2)
            return jsonify({"ok": False, "error": str(e2)}), 400
    except TranscriptsDisabled:
        return jsonify({"ok": False, "error": "Transcripts are disabled for this video"}), 400
    except Exception as e:
        _log_exc("youtube_transcript", e)
        return jsonify({"ok": False, "error": str(e)}), 500

# ══════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════
import socket, subprocess

def _kill_port(port):
    """Kill any process already holding our port so restart always works."""
    try:
        for proc in psutil.process_iter(["pid","name","connections"]):
            try:
                for conn in (proc.info.get("connections") or []):
                    if hasattr(conn, "laddr") and conn.laddr.port == port:
                        proc.kill(); time.sleep(0.3)
            except Exception: pass
    except Exception: pass

def _port_open(port):
    try: s = socket.create_connection(("127.0.0.1",port),timeout=0.3); s.close(); return True
    except Exception: return False

def _wait_for_flask(port=7890, timeout=15):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _port_open(port): return True
        time.sleep(0.08)
    return False

def run_flask():
    import logging; logging.getLogger("werkzeug").setLevel(logging.ERROR)
    app_flask.run(host="127.0.0.1", port=7890, debug=False, use_reloader=False, threaded=True)

def _clear_edge_locks():
    """
    Delete stale SingletonLock / SingletonCookie files from the Edge profile.
    When the app is force-closed these lock files remain and cause Edge to either
    crash on relaunch or refuse to open — which can also crash Windows Explorer
    due to shell extensions. Must run before launching Edge.
    """
    profile_dir = os.path.join(BASE, ".edge_profile")
    for fname in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
        for root, dirs, files in os.walk(profile_dir):
            if fname in files:
                try: os.remove(os.path.join(root, fname))
                except Exception: pass

def _find_browser():
    try:
        import winreg
    except ImportError:
        return None
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    for p in candidates:
        if os.path.isfile(p): return p
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe")
        path, _ = winreg.QueryValueEx(key, ""); winreg.CloseKey(key)
        if os.path.isfile(path): return path
    except Exception: pass
    return None

def _launch_app_window(url):
    _clear_edge_locks()   # remove stale lock files before launch — prevents crash/hang
    exe = _find_browser()
    if not exe: return None
    profile_dir = os.path.join(BASE, ".edge_profile"); os.makedirs(profile_dir, exist_ok=True)
    try:
        proc = subprocess.Popen([exe, "--app="+url, "--user-data-dir="+profile_dir,
            "--window-size=640,960", "--window-position=200,40",
            "--no-first-run", "--no-default-browser-check",
            "--disable-extensions", "--disable-default-apps",
            "--disable-background-networking", "--disable-sync",
            "--password-store=basic",
            # Performance flags — reduce compositor overhead for snappier window behavior
            "--disable-background-timer-throttling",
            "--disable-renderer-backgrounding",
            "--disable-backgrounding-occluded-windows",
            "--disable-ipc-flooding-protection",
            "--enable-features=VaapiVideoDecoder",
            "--force-device-scale-factor=1",
        ])
        # Store PID so window finder can target it precisely
        global _edge_pid
        _edge_pid = proc.pid
        return proc
    except Exception as e: print("Browser launch error: "+str(e)); return None

# Win32 constants
GWL_STYLE     = -16
WS_CAPTION    = 0x00C00000   # titlebar — remove this
WS_THICKFRAME = 0x00040000   # resize border — KEEP THIS
WS_SYSMENU    = 0x00080000   # system menu — remove this
WS_MINIMIZEBOX= 0x00020000
WS_MAXIMIZEBOX= 0x00010000
SWP_FRAMECHANGED  = 0x0020
SWP_NOMOVE        = 0x0002
SWP_NOSIZE        = 0x0001
SWP_NOZORDER      = 0x0004
SWP_NOOWNERZORDER = 0x0200
HWND_TOPMOST   = -1
HWND_NOTOPMOST = -2

def _apply_frameless_to_hwnd(hwnd):
    """Remove title bar but KEEP resize borders so the window is still resizable."""
    if not hwnd or not HAS_WIN32: return
    try:
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_STYLE)
        # Remove caption and sysmenu, keep thickframe + minimize/maximize
        style = (style & ~WS_CAPTION & ~WS_SYSMENU) | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_STYLE, style)
        ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOOWNERZORDER | SWP_FRAMECHANGED)
        _log(f"Frameless applied to hwnd={hwnd}")
    except Exception as e:
        _log(f"_apply_frameless_to_hwnd failed: {e}")





def _keep_alive(proc=None):
    """Keep the process alive until the browser window closes. No tkinter needed."""
    import threading
    stop_evt = threading.Event()
    def _watch():
        while not stop_evt.is_set():
            if proc and proc.poll() is not None:
                stop_evt.set(); break
            time.sleep(0.2)  # check every 200ms instead of 800ms
    if proc:
        t = threading.Thread(target=_watch, daemon=True)
        t.start()
        stop_evt.wait()  # blocks main thread until Edge closes
    else:
        stop_evt.wait()  # wait forever (fallback browser mode)


def main():
    _kill_port(7890)
    start_hotkeys()
    flask_thread = threading.Thread(target=run_flask, daemon=True); flask_thread.start()
    if not _wait_for_flask(): print("ERROR: Flask server failed to start."); sys.exit(1)
    url = "http://127.0.0.1:7890"

    # ── PRIMARY: pywebview ──
    # DIAGNOSTIC: print webview status so you can see why it might fall back
    print("="*55)
    print("PYWEBVIEW DIAGNOSTIC")
    print(f"  HAS_WEBVIEW : {HAS_WEBVIEW}")
    if HAS_WEBVIEW:
        try: print(f"  version     : {webview.__version__}")
        except Exception: print("  version     : unknown")
        try:
            import clr
            print("  backend     : pythonnet/EdgeChromium")
        except ImportError: pass
        try:
            from webview.platforms import winforms
            print("  platform    : winforms")
        except Exception as pe: print(f"  platform err: {pe}")
    else:
        print("  >> pywebview not installed — run: pip install pywebview")
    print("="*55)
    if HAS_WEBVIEW:
        try:
            os.environ["WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"] = " ".join([
                "--disable-background-timer-throttling",
                "--disable-renderer-backgrounding",
                "--disable-backgrounding-occluded-windows",
                "--disable-ipc-flooding-protection",
                "--disable-gpu-sandbox",
                "--force-device-scale-factor=1",
            ])
            global _webview_window

            class WVApi:
                def minimize(self):
                    if _webview_window:
                        _webview_window.minimize()
                def close(self):
                    os._exit(0)  # instant close, no delay

            _webview_window = webview.create_window(
                title="Fiuxxed's AutoTyper v9.1",
                url=url,
                width=620, height=960,
                min_size=(420, 600),
                resizable=True,
                frameless=True,
                easy_drag=False,   # Only drag from elements with class pywebview-drag-region (title bar)
                transparent=False,                     # FIX: was True – now solid background
                on_top=bool(cfg.get("always_on_top", True)),
                background_color="#07070e",            # FIX: match your theme
                js_api=WVApi(),
            )

            webview.start(debug=False)
            return
        except Exception as e:
            _log_exc("pywebview failed — falling back to Edge subprocess", e)

    # ── FALLBACK: Edge app mode ──
    proc = _launch_app_window(url)
    if proc: _keep_alive(proc); return

    import webbrowser; print("Opening in browser: "+url); webbrowser.open(url); _keep_alive()

# ══════════════════════════════════════════════════════════════════════
#  ADMIN ROUTES — IP-locked + password protected
# ══════════════════════════════════════════════════════════════════════

@app_flask.route("/api/admin/ping", methods=["POST"])
def api_admin_ping():
    """Returns whether this machine is the admin machine. No password needed."""
    return jsonify({"is_admin": _is_admin_ip(request)})

@app_flask.route("/api/admin/login", methods=["POST"])
def api_admin_login():
    ok, err = _require_admin(request)
    if not ok: return err
    data = request.json or {}
    if data.get("password") != _ADMIN_PASSWORD:
        return jsonify({"error": "Wrong password"}), 403
    import secrets
    token = secrets.token_hex(32)
    _ADMIN_SESSION.add(token)
    return jsonify({"ok": True, "token": token})

@app_flask.route("/api/admin/stats", methods=["POST"])
def api_admin_stats():
    ok, err = _require_admin(request)
    if not ok: return err
    data = request.json or {}
    if data.get("admin_token") not in _ADMIN_SESSION:
        return jsonify({"error": "Not authenticated"}), 403
    # Get online users from presence file (Supabase is client-side, so we track what we know locally)
    return jsonify({
        "ok": True,
        "uptime_s": int(time.time() - _start_time),
        "settings": cfg,
        "history_count": len(load_history()),
        "formula_count": len(load_formula_lib()),
        "chat_message_count": _admin_msg_count,
        "online_users": list(_admin_online_users.values()),
        "device_count": len(_app_devices),
        "devices": [{"username": v["username"]} for v in _app_devices.values()],
    })

@app_flask.route("/api/admin/presence_update", methods=["POST"])
def api_admin_presence_update():
    """Called by clients to register their presence with the server for admin tracking."""
    data = request.json or {}
    uid  = data.get("user_id","")
    name = data.get("username","")
    if uid and name:
        _admin_online_users[uid] = {"user_id": uid, "username": name, "last_seen": time.time()}
        # Expire stale entries
        stale = [k for k,v in _admin_online_users.items() if time.time()-v["last_seen"] > 15]
        for k in stale: del _admin_online_users[k]
    return jsonify({"ok": True})

@app_flask.route("/api/admin/gc", methods=["POST"])
def api_admin_gc():
    ok, err = _require_admin(request)
    if not ok: return err
    data = request.json or {}
    if data.get("admin_token") not in _ADMIN_SESSION:
        return jsonify({"error": "Not authenticated"}), 403
    import gc
    freed = gc.collect()
    try:
        import psutil, os as _os
        proc = psutil.Process(_os.getpid())
        mem_mb = round(proc.memory_info().rss / 1024 / 1024, 1)
    except: mem_mb = None
    return jsonify({"ok": True, "freed": freed, "memory_mb": mem_mb})

@app_flask.route("/api/admin/broadcast", methods=["POST"])
def api_admin_broadcast():
    """Store a system announcement that chat clients pick up on next poll."""
    ok, err = _require_admin(request)
    if not ok: return err
    data = request.json or {}
    if data.get("admin_token") not in _ADMIN_SESSION:
        return jsonify({"error": "Not authenticated"}), 403
    msg  = data.get("message","").strip()
    mtype= data.get("type","info")
    target=data.get("target","all")
    if not msg: return jsonify({"error": "No message"})
    _admin_announcements.append({"text": msg, "ts": time.time(), "type": mtype, "target": target})
    return jsonify({"ok": True})

@app_flask.route("/api/admin/announcement/poll", methods=["GET"])
def api_admin_announcement_poll():
    """Chat clients poll this to see if there's a pinned announcement."""
    since = float(request.args.get("since", 0))
    msgs = [a for a in _admin_announcements if a["ts"] > since]
    return jsonify({"announcements": msgs})

@app_flask.route("/api/admin/clear_history", methods=["POST"])
def api_admin_clear_history():
    ok, err = _require_admin(request)
    if not ok: return err
    data = request.json or {}
    if data.get("admin_token") not in _ADMIN_SESSION:
        return jsonify({"error": "Not authenticated"}), 403
    save_history_file([])
    return jsonify({"ok": True})

@app_flask.route("/api/admin/banned_words", methods=["POST"])
def api_admin_banned_words():
    ok, err = _require_admin(request)
    if not ok: return err
    data = request.json or {}
    if data.get("admin_token") not in _ADMIN_SESSION:
        return jsonify({"error": "Not authenticated"}), 403
    action = data.get("action")
    word   = (data.get("word") or "").strip().lower()
    if action == "get":
        return jsonify({"ok": True, "words": list(_admin_banned_words)})
    elif action == "add" and word:
        _admin_banned_words.add(word)
        return jsonify({"ok": True, "words": list(_admin_banned_words)})
    elif action == "remove" and word:
        _admin_banned_words.discard(word)
        return jsonify({"ok": True, "words": list(_admin_banned_words)})
    return jsonify({"error": "Unknown action"})

@app_flask.route("/api/admin/delete_message", methods=["POST"])
def api_admin_delete_message():
    ok, err = _require_admin(request)
    if not ok: return err
    data = request.json or {}
    if data.get("admin_token") not in _ADMIN_SESSION:
        return jsonify({"error": "Not authenticated"}), 403
    msg_id = data.get("message_id","")
    if not msg_id: return jsonify({"error": "No message_id"})
    try:
        import urllib.request as _ur, json as _json
        req = _ur.Request(
            f"{_SUPA_URL}/rest/v1/messages?id=eq.{msg_id}",
            method="DELETE",
            headers={"apikey": _SUPA_KEY, "Authorization": f"Bearer {_SUPA_SERVICE_KEY}"}
        )
        with _ur.urlopen(req) as resp:
            return jsonify({"ok": True, "status": resp.status})
    except Exception as e:
        _log(f"Admin delete failed: {e}")
        return jsonify({"error": str(e)})
    data = request.json or {}
    _ADMIN_SESSION.discard(data.get("admin_token",""))
    return jsonify({"ok": True})

# Admin state
_start_time = time.time()
_admin_announcements = []
_admin_banned_words  = set(["nigger","nigga","faggot","retard","cunt"])
_admin_online_users  = {}   # uid -> {user_id, username, last_seen}
_admin_msg_count     = 0
_app_devices         = {}   # session_id -> {username, last_seen}  — every open app instance

@app_flask.route("/api/admin/app_ping", methods=["POST"])
def api_admin_app_ping():
    """Called every 5s by every open app instance, signed in or not."""
    data = request.json or {}
    sid  = data.get("session_id","")
    name = data.get("username","Anonymous")
    if sid:
        _app_devices[sid] = {"username": name, "last_seen": time.time()}
        # Expire devices not seen in 15s
        stale = [k for k,v in _app_devices.items() if time.time()-v["last_seen"] > 15]
        for k in stale: del _app_devices[k]
    return jsonify({"ok": True, "device_count": len(_app_devices)})

@app_flask.route("/api/admin/app_close", methods=["POST"])
def api_admin_app_close():
    data = request.json or {}
    _app_devices.pop(data.get("session_id",""), None)
    return jsonify({"ok": True, "device_count": len(_app_devices)})

if __name__ == "__main__":
    main()