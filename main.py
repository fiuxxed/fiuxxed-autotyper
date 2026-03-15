# ══════════════════════════════════════════════════════════════════════
#  Fiuxxed's AutoTyper v9.1  —  main.py
# ══════════════════════════════════════════════════════════════════════
import threading, time, random, json, os, sys, re, io, base64, traceback, ctypes

BASE      = os.path.dirname(os.path.abspath(__file__))
SAVE_FILE = os.path.join(BASE, "settings.json")
WEB_DIR   = os.path.join(BASE, "web")

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
    HAS_PLOT = True
except ImportError: HAS_PLOT = False

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
def capture_window(hwnd=None):
    if not HAS_MSS: raise RuntimeError("mss/pillow not installed")
    if hwnd and HAS_WIN32:
        try: win32gui.SetForegroundWindow(hwnd); time.sleep(0.45)
        except Exception: pass
        rect = win32gui.GetWindowRect(hwnd)
        x1, y1, x2, y2 = rect; w = x2-x1; h = y2-y1
        if w < 10 or h < 10: raise RuntimeError("Window too small")

        # Use PrintWindow with PW_RENDERFULLCONTENT (flag=3) so GPU-accelerated apps
        # like Discord, Chrome, and games render correctly instead of returning a black bitmap.
        # PW_RENDERFULLCONTENT = 0x2, combined with PW_CLIENTONLY=0x1 gives flag 3.
        # Fall back to mss screen-grab if PrintWindow returns a blank (all-zero) result.
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
            # PW_RENDERFULLCONTENT forces the DWM compositor to copy GPU-rendered content
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

        # Fallback: mss screen-grab (fails for minimized/occluded GPU windows)
        if img is None:
            region = {"top": y1, "left": x1, "width": w, "height": h}
            with mss.mss() as sct:
                raw = sct.grab(region)
                img = Image.frombytes("RGB", raw.size, raw.bgra, "raw", "BGRX")
    else:
        with mss.mss() as sct:
            region = sct.monitors[1]
            raw = sct.grab(region)
            img = Image.frombytes("RGB", raw.size, raw.bgra, "raw", "BGRX")
    img = img.convert("RGB")
    # Scale up for better AI readability (scan/math); min 1600px width
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
    """Draw a geometry diagram from structured AI-extracted data."""
    if not HAS_PLOT: return None
    try:
        points  = diagram_data.get("points", {})   # {"A":[x,y], "B":[x,y], ...}
        edges   = diagram_data.get("edges", [])    # [["A","B"], ["B","C"], ...]
        angles  = diagram_data.get("angles", {})   # {"B": "90°", ...}
        labels  = diagram_data.get("labels", {})   # {"AB": "6y-16", "BA": "4y+6", ...}
        title   = diagram_data.get("title", "Diagram")

        if not points: return None

        fig, ax = plt.subplots(figsize=(5.5, 4.5), facecolor="#07070e")
        ax.set_facecolor("#0d0c18")
        ax.set_aspect("equal")

        # Draw edges
        for edge in edges:
            if len(edge) == 2 and edge[0] in points and edge[1] in points:
                x0,y0 = points[edge[0]]
                x1,y1 = points[edge[1]]
                ax.plot([x0,x1],[y0,y1], color="#a855f7", linewidth=2.2, zorder=3)
                # Draw edge label at midpoint
                key = edge[0]+edge[1]
                lbl = labels.get(key) or labels.get(edge[1]+edge[0])
                if lbl:
                    mx,my = (x0+x1)/2, (y0+y1)/2
                    # offset label slightly perpendicular
                    dx,dy = x1-x0, y1-y0
                    length = max((dx**2+dy**2)**0.5, 0.001)
                    ox,oy = -dy/length*0.3, dx/length*0.3
                    ax.text(mx+ox, my+oy, lbl, color="#f59e0b", fontsize=8.5,
                            ha="center", va="center",
                            bbox=dict(boxstyle="round,pad=0.2", fc="#07070e", ec="none", alpha=0.8))

        # Draw right-angle markers
        for pt_name, ang_label in angles.items():
            if pt_name not in points: continue
            if "90" in str(ang_label) or "right" in str(ang_label).lower():
                # find two edges meeting at this point
                connected = [e for e in edges if pt_name in e]
                if len(connected) >= 2:
                    px,py = points[pt_name]
                    def unit_toward(from_pt, to_name):
                        tx,ty = points[to_name]
                        dx,dy = tx-from_pt[0], ty-from_pt[1]
                        l = max((dx**2+dy**2)**0.5, 0.001)
                        return dx/l, dy/l
                    other1 = connected[0][1] if connected[0][0]==pt_name else connected[0][0]
                    other2 = connected[1][1] if connected[1][0]==pt_name else connected[1][0]
                    if other1 in points and other2 in points:
                        s = 0.25
                        u1 = unit_toward((px,py), other1)
                        u2 = unit_toward((px,py), other2)
                        sq = plt.Polygon([
                            [px+s*u1[0],          py+s*u1[1]],
                            [px+s*u1[0]+s*u2[0],  py+s*u1[1]+s*u2[1]],
                            [px+s*u2[0],          py+s*u2[1]],
                            [px,                  py],
                        ], fill=False, edgecolor="#10b981", linewidth=1.2)
                        ax.add_patch(sq)
            else:
                # label the angle as text near the vertex
                px,py = points[pt_name]
                ax.text(px, py-0.35, str(ang_label), color="#f59e0b", fontsize=8,
                        ha="center", va="top")

        # Draw points and labels
        xs = [v[0] for v in points.values()]
        ys = [v[1] for v in points.values()]
        ax.scatter(xs, ys, color="#a855f7", s=40, zorder=5)
        for name,(px,py) in points.items():
            # offset label away from centroid
            cx,cy = sum(xs)/len(xs), sum(ys)/len(ys)
            ox = 0.3 * (1 if px >= cx else -1)
            oy = 0.3 * (1 if py >= cy else -1)
            ax.text(px+ox, py+oy, name, color="#e9d5ff", fontsize=11,
                    fontweight="bold", ha="center", va="center")

        # Padding
        pad = 1.0
        ax.set_xlim(min(xs)-pad, max(xs)+pad)
        ax.set_ylim(min(ys)-pad, max(ys)+pad)

        ax.set_title(title, color="#e9d5ff", fontsize=10, pad=8)
        ax.tick_params(colors="#6b5b9e", labelsize=7)
        ax.grid(True, color="#28204a", alpha=0.4, linewidth=0.6)
        for sp in ax.spines.values(): sp.set_color("#28204a")

        plt.tight_layout()
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=110, bbox_inches="tight", facecolor="#07070e")
        plt.close(fig)
        return base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception:
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

def ai_scan(b64, strictness="flag_all", examine_examples=False, extra_context=None, extra_images=None):
    client = get_client()
    strict_note = ("Flag ANY answer that looks even slightly off, unclear, or potentially wrong."
        if strictness == "flag_all" else "Only flag answers that are clearly and definitively wrong.")
    img_parts = [{"type":"image_url","image_url":{"url":f"data:image/png;base64,{b64}"}}]
    for ei in (extra_images or []):
        if ei: img_parts.append({"type":"image_url","image_url":{"url":f"data:image/jpeg;base64,{ei}" if not ei.startswith("data:") else ei}})
    resp = client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[{"role":"user","content":img_parts + [
            {"type":"text","text":(
                (("EXTRA CONTEXT (use this to help answer questions):\n" + "\n\n".join(c.strip() for c in extra_context if c and c.strip()) + "\n\n") if extra_context and any(c.strip() for c in extra_context) else "") +
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
                "2. question: what is being asked. For passage-based boxes: 'What is the main idea of the passage about [topic]?'\n"
                "3. type: MULTIPLE_CHOICE, TRUE_FALSE, or WRITTEN\n"
                "4. answered: true ONLY if a real answer is already written (not just cursor, underline, or empty field)\n"
                "5. user_answer: what they wrote (null if unanswered)\n"
                "6. correct_answer: YOUR complete answer using passage/document content\n"
                "7. is_correct: true/false/null (null if unanswered)\n"
                f"8. {strict_note}\n"
                "9. correction: if wrong, the corrected answer in the same format\n"
                "10. confident: false only if you genuinely cannot read the content\n"
                + ("11. Where helpful, include a brief worked example.\n\n"
                   if examine_examples else
                   "11. Do NOT include worked examples — answer only.\n\n")
                + "Return ONLY valid JSON array (no markdown):\n"
                '[{"question_label":"Main Idea","question":"What is the main idea of the passage?","type":"WRITTEN",'
                '"answered":false,"user_answer":null,"correct_answer":"The Enlightenment was an intellectual movement that spread democratic ideas across Europe and the Americas in the 1700s.",'
                '"is_correct":null,"correction":null,"confident":true}]\n'
                "If no questions found: []"
            )}
        ]}],
        max_tokens=3500, temperature=0.1
    )
    raw = re.sub(r"```(?:json)?|```", "", resp.choices[0].message.content.strip()).strip()
    m = re.search(r'\[.*\]', raw, re.DOTALL)
    if m: raw = m.group(0)
    return json.loads(raw)

def ai_math(b64, examine_examples=False, extra_context=None, extra_images=None):
    client = get_client()
    # Build image list: primary screenshot + any extra images
    img_parts = [{"type":"image_url","image_url":{"url":f"data:image/png;base64,{b64}"}}]
    for ei in (extra_images or []):
        if ei: img_parts.append({"type":"image_url","image_url":{"url":f"data:image/jpeg;base64,{ei}" if not ei.startswith("data:") else ei}})
    resp = client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[{"role":"user","content":img_parts + [
            {"type":"text","text":(
                (("EXTRA CONTEXT (use this to help solve problems):\n" + "\n\n".join(c.strip() for c in extra_context if c and c.strip()) + "\n\n") if extra_context and any(c.strip() for c in extra_context) else "") +
                "Look at this screenshot. Find EVERY math or geometry problem visible.\n\n"
                "CRITICAL — HOW TO IDENTIFY PROBLEMS:\n"
                "- A number inside a GREEN CIRCLE or any colored circle (① ② ③) = problem number\n"
                "- Any diagram, triangle, or geometric figure NEXT TO or BELOW a circled number = that IS the problem\n"
                "- Plain numbers '1.', '2.', 'Q1' etc followed by text or a diagram = problems\n"
                "- NEVER skip a numbered item just because it shows a diagram\n"
                "- Geometry diagrams with angle measures, side lengths, or shape labels are ALL math problems\n\n"
                "PROBLEM TYPES — adapt your output to what you see:\n"
                "- EQUATION_SOLVE: solve for x (or variable). Give vertical_method as column layout.\n"
                "- FILL_BLANK: one or more blanks to fill. answer = the word/number(s) that go in the blank(s).\n"
                "- MULTIPLE_CHOICE: select one or more options. answer = the correct choice letter(s) and text.\n"
                "- GEOMETRY: diagram, angles, sides. Use diagram_data and diagram_description.\n"
                "- SHORT_ANSWER: brief numeric or symbolic answer.\n\n"
                "For EACH problem provide ALL fields:\n"
                "- problem_type: one of EQUATION_SOLVE, FILL_BLANK, MULTIPLE_CHOICE, GEOMETRY, SHORT_ANSWER\n"
                "- problem_label: number/label if visible (e.g. '1.', 'Q3') — null if not visible\n"
                "- problem: exact problem text. If it is a diagram problem describe it fully\n"
                "- answer: CLEAN final answer. For equations: 'x = 57' or 'y = 2'. Not a long expression. For fill-in: the exact blank value(s). For MCQ: e.g. 'B) 42'\n"
                "- steps: array of step-by-step solution strings\n"
                "- explanations: array of plain English explanation per step (SAME length as steps)\n"
                "- vertical_method: For EQUATION_SOLVE or when algebra is shown in columns, use a COLUMN layout: one equation per line, right-align style. Example (use \\n for newlines):\n"
                "  '    5x - 16 = 4x + 41\\n           x = 57'\n"
                "  Or: '  2x + 3 = 7\\n  2x = 4\\n  x = 2'. No prose, only equations. For fill-in/MCQ/geometry with no algebra column, use null or a brief calculation if needed.\n"
                "- graph_eq: for ANY plottable function provide 'y=expr'. For geometry problems provide null\n"
                "- diagram_data: for geometry/diagram problems, provide a JSON object with:\n"
                "    points: {\"A\":[x,y], \"B\":[x,y], ...} — place points on a coordinate grid that accurately reflects the shape. Use real proportions from the image.\n"
                "    edges: [[\"A\",\"B\"],[\"B\",\"C\"],...] — line segments to draw\n"
                "    angles: {\"B\":\"90°\",...} — angle label at each vertex (include 'right' or '90' for right angles so a square marker is drawn)\n"
                "    labels: {\"AB\":\"6y-16\",\"BC\":\"4y+6\",...} — side length expressions from the problem\n"
                "    title: short description e.g. 'Triangle ABC'\n"
                "  null for pure algebra/non-geometry problems\n"
                "- diagram_description: plain text description of the diagram for geometry problems. null for pure algebra\n"
                "- has_graph: true if graph_eq is not null\n"
                "- confident: false only if you cannot clearly read the problem\n\n"
                "RULES:\n"
                "- Every graphable function MUST have graph_eq\n"
                "- Geometry problems with triangles/shapes must have diagram_description\n"
                "- Steps and explanations must ALWAYS be present\n"
                + ("- Where it helps understanding, include a brief additional worked example after the main solution.\n\n"
                   if examine_examples else
                   "- Do NOT add extra worked examples or sample problems — solve the given problem only.\n\n")
                + "Return ONLY valid JSON array (no markdown):\n"
                '[{"problem_type":"EQUATION_SOLVE","problem_label":"1.","problem":"2x+3=7",'
                '"steps":["Subtract 3: 2x=4","Divide by 2: x=2"],'
                '"explanations":["Remove constant from left","Isolate x"],'
                '"vertical_method":"  2x + 3 = 7\\n  2x = 4\\n  x = 2","answer":"x = 2",'
                '"graph_eq":null,"diagram_data":null,"diagram_description":null,"has_graph":false,"confident":true}]\n'
                "No math found: []"
            )}
        ]}],
        max_tokens=4000, temperature=0.1
    )
    raw = re.sub(r"```(?:json)?|```", "", resp.choices[0].message.content.strip()).strip()
    m = re.search(r'\[.*\]', raw, re.DOTALL)
    if m: raw = m.group(0)
    return json.loads(raw)


def ai_double_check(question):
    client = get_client()
    resp = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
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

def ai_qa(question, history, is_followup, extra_context=None, extra_images=None):
    client = get_client()
    has_images = bool(extra_images and any(ei for ei in extra_images))
    model = "meta-llama/llama-4-scout-17b-16e-instruct" if has_images else "llama-3.3-70b-versatile"
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
        return text  # plain string — non-vision models require this
    if is_followup and history:
        messages = list(history) + [{"role":"user","content":_uc(f"{question}{ctx_block}\n\n[Follow-up. Be more detailed. 2-4 sentences max.]")}]
        max_tok = 400
    else:
        if has_images:
            sys_msg = (
                "You are an expert at analyzing images in detail. Look at the entire image carefully and use high-resolution detail. "
                "When identifying living things (animals, plants, fungi, etc.): use specific, precise names — including rare, uncommon, or lesser-known species. "
                "Prefer scientific names when the creature is unusual (e.g. 'axolotl, Ambystoma mexicanum' or 'satanic leaf-tailed gecko, Uroplatus phantasticus'). "
                "For common species still give the exact common name (e.g. 'blue morpho butterfly' not just 'butterfly'). "
                "For objects, text, or scenes: describe precisely and name anything you can identify. "
                "Answer directly and concisely but with enough detail to be accurate and specific."
            )
            max_tok = 400
        else:
            sys_msg = "You are a sharp, direct assistant. Answer in ONE sentence, max 20 words. No intros. Just the answer."
            max_tok = 120
        messages = [{"role":"system","content":sys_msg}, {"role":"user","content":_uc(question+ctx_block)}]
    resp = client.chat.completions.create(model=model, messages=messages, max_tokens=max_tok, temperature=0.4)
    answer = resp.choices[0].message.content.strip()
    if not is_followup:
        new_history = [{"role":"system","content":"You are a helpful, knowledgeable assistant."},
                       {"role":"user","content":question},{"role":"assistant","content":answer}]
    else:
        new_history = list(history) + [{"role":"user","content":question},{"role":"assistant","content":answer}]
    return answer, new_history


# ══════════════════════════════════════════════════════════════════════
#  FLASK APP
# ══════════════════════════════════════════════════════════════════════
app_flask   = Flask(__name__, static_folder=WEB_DIR)
cfg         = load_cfg()
engine      = TypingEngine()
_hk_listener    = None
_voice_listener = None
_type_state     = {"phase": "idle", "progress": 0, "status": "Ready"}
_webview_window = None
_windows_cache = []

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
    sk = km.get(cfg.get("hotkey_start","F8").lower())
    pk = km.get(cfg.get("hotkey_pause","F9").lower())
    stk = km.get(cfg.get("hotkey_stop","F10").lower())
    def _press(key):
        if sk  and key == sk:  _trigger_start()
        if pk  and key == pk:  _toggle_pause()
        if stk and key == stk: _do_stop()
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
    key = data.get("groq_api_key","").strip()
    cfg["groq_api_key"] = key; save_cfg(cfg)
    return jsonify({"ok": True, "saved": bool(key)})

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
        # Get our own PID so we can exclude the app window from the list
        our_pid = os.getpid()
        # Also collect child PIDs (Edge app window launched by us)
        our_pids = {our_pid}
        try:
            our_proc = psutil.Process(our_pid)
            for child in our_proc.children(recursive=True):
                our_pids.add(child.pid)
        except Exception: pass
        # Also exclude by _edge_pid if set
        if _edge_pid: our_pids.add(_edge_pid)

        # Build pid→name map once — avoids opening a new Process handle per window inside
        # the EnumWindows callback, which hammers the kernel on every Refresh call.
        pid_name = {}
        try:
            for proc in psutil.process_iter(["pid", "name"]):
                try: pid_name[proc.info["pid"]] = proc.info["name"].lower().replace(".exe", "")
                except Exception: pass
        except Exception: pass

        wins = []
        def cb(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd): return
            # Skip minimized windows — they cannot be screenshotted (PrintWindow returns black)
            if win32gui.IsIconic(hwnd): return
            title = win32gui.GetWindowText(hwnd)
            if not title or len(title) < 2: return
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                # Skip our own app window entirely
                if pid in our_pids: return
                name = pid_name.get(pid, "unknown")
            except Exception: name = "unknown"
            is_browser = any(b in name for b in BROWSERS)
            wins.append({"hwnd":hwnd,"title":title,"name":name,"is_browser":is_browser})
        win32gui.EnumWindows(cb, None)

        # Sort: non-browsers first (alphabetical), then browsers (alphabetical)
        # This puts Google Docs, Word, etc. at the top and YouTube tabs at the bottom
        wins.sort(key=lambda x: (1 if x["is_browser"] else 0, x["title"].lower()))

        # Mark the best suggestion — first non-browser, or first browser if nothing else
        if wins:
            wins[0]["suggested"] = True

        _windows_cache = wins
        return jsonify({"windows": wins})
    except Exception as e:
        # FIX: Always return a proper JSON response, even on error
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
            return jsonify({"ok":True,"result":{"problems":problems}})
        else:
            questions = ai_scan(b64, strictness, examine_examples=cfg.get("examine_examples", False), extra_context=extra_context, extra_images=extra_images)
            add_to_history("scan", questions)
            return jsonify({"ok":True,"result":{"questions":questions}})
    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()})

@app_flask.route("/api/screenshot/preview", methods=["POST"])
def api_screenshot_preview():
    """Take a screenshot and return full-res b64 for region selector.
    The UI displays it scaled down via CSS but fractions are computed against
    the naturalWidth/naturalHeight of the img element (= actual pixel dims here).
    """
    data = request.json or {}
    hwnd = data.get("hwnd")
    try:
        img = capture_window(hwnd)
        ow, oh = img.size
        b64 = img_to_b64(img)
        return jsonify({"ok":True,"b64":b64,"w":ow,"h":oh})
    except Exception as e:
        return jsonify({"error":str(e)})

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
    if not question: return jsonify({"error":"No question"})
    extra_images = data.get("extra_images", [])
    try:
        answer, new_history = ai_qa(question, history, is_followup, extra_context, extra_images)
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
            model="llama-3.3-70b-versatile",
            messages=[{"role":"user","content":p}],
            max_tokens=min(tokens, 600), temperature=0.75
        )
        return jsonify({"result": resp.choices[0].message.content.strip()})
    except Exception as e:
        return jsonify({"error": str(e)})

@app_flask.route("/api/formula_library", methods=["GET"])
def api_formula_library(): return jsonify({"formulas": load_formula_lib()})

@app_flask.route("/api/formula_library/clear", methods=["POST"])
def api_formula_library_clear(): save_formula_lib([]); return jsonify({"ok":True})

@app_flask.route("/api/history", methods=["GET"])
def api_history(): return jsonify({"history": load_history()})

@app_flask.route("/api/history/clear", methods=["POST"])
def api_history_clear(): save_history_file([]); return jsonify({"ok":True})

# Local Music
AUDIO_EXTS = {".mp3",".flac",".wav",".ogg",".m4a",".aac",".wma",".opus",".aiff"}

def _get_music_roots():
    roots = []
    home = os.path.expanduser("~")
    for c in [os.path.join(home,"Music"),os.path.join(home,"Downloads"),os.path.join(home,"Desktop"),r"C:\Users\Public\Music"]:
        if os.path.isdir(c): roots.append(c)
    return roots

def _scan_music_dir(directory, max_files=400):
    results = []
    try:
        for root, dirs, files in os.walk(directory):
            dirs[:] = [d for d in dirs if not d.startswith(".")]
            for fname in files:
                ext = os.path.splitext(fname)[1].lower()
                if ext in AUDIO_EXTS:
                    full = os.path.join(root, fname)
                    results.append({"path":full,"name":os.path.splitext(fname)[0],"ext":ext.lstrip("."),"size":os.path.getsize(full)})
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

# Cache for YouTube search pagination: one query's full result set (up to 50), keyed by query
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
            # Use cache if same query and we have enough results for this page
            if query == _music_search_cache_query and _music_search_cache_results is not None:
                if start < len(_music_search_cache_results):
                    return return_paginated(_music_search_cache_results)
                # Page beyond cache — need to refetch (e.g. they went to page 4)
                if page == 1:
                    _music_search_cache_results = None  # force refetch below

            if _music_search_cache_query != query or _music_search_cache_results is None or page > 1:
                # Fetch up to 50 so we have ~3 pages
                ydl_opts = {
                    "quiet": True,
                    "no_warnings": True,
                    "extract_flat": "in_playlist",
                    "default_search": "ytsearch50",
                    "skip_download": True,
                    "ignoreerrors": True,
                }
                with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                    info = ydl.extract_info(f"ytsearch50:{query}", download=False)
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
    mime_map = {".mp3":"audio/mpeg",".flac":"audio/flac",".wav":"audio/wav",".ogg":"audio/ogg",".m4a":"audio/mp4",".aac":"audio/aac",".wma":"audio/x-ms-wma",".opus":"audio/opus",".aiff":"audio/aiff"}
    mime = mime_map.get(ext, "audio/mpeg")
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
    # Instant exit — don't wait for pywebview/Edge cleanup, just die
    threading.Thread(target=lambda: (time.sleep(0.15), os._exit(0)), daemon=True).start()
    return jsonify({"ok": True})

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
            "--window-size=620,960", "--window-position=200,40",
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
SWP_FRAMECHANGED = 0x0020

# _edge_pid: PID of the Edge process we launched — most reliable finder anchor
_edge_pid = None





def _keep_alive(proc=None):
    """Keep the process alive until the browser window closes. No tkinter needed."""
    import threading
    stop_evt = threading.Event()
    def _watch():
        while not stop_evt.is_set():
            if proc and proc.poll() is not None:
                stop_evt.set(); break
            time.sleep(0.8)
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
                    threading.Thread(target=lambda: (time.sleep(0.1), os._exit(0)), daemon=True).start()

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

if __name__ == "__main__":
    main()