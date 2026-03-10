# ══════════════════════════════════════════════════════════════════════
#  Fiuxxed's AutoTyper v9.0  —  main.py
# ══════════════════════════════════════════════════════════════════════
import threading, time, random, json, os, sys, re, io, base64, traceback, ctypes

BASE      = os.path.dirname(os.path.abspath(__file__))
SAVE_FILE = os.path.join(BASE, "settings.json")
WEB_DIR   = os.path.join(BASE, "web")

DEFAULTS = {
    "text": "", "wpm": 80, "countdown": 5,
    "repeat_count": 1, "repeat_delay": 2.0,
    "stop_after_chars": 0, "stop_after_words": 0, "stop_after_lines": 0,
    "typo_chance": 0, "stutter_chance": 40, "stutter_duration": 2.0,
    "thinking_pause_chance": 3, "thinking_pause_min": 300, "thinking_pause_max": 800,
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
                    elif not iswc and inw: inw = False; stutter = False
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
        region = {"top": y1, "left": x1, "width": w, "height": h}
    else:
        with mss.mss() as sct: region = sct.monitors[1]
    with mss.mss() as sct:
        raw = sct.grab(region)
        img = Image.frombytes("RGB", raw.size, raw.bgra, "raw", "BGRX")
    img = img.convert("RGB")
    if img.width < 1200:
        scale = 1200 / img.width
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

def ai_scan(b64, strictness="flag_all", examine_examples=False):
    client = get_client()
    strict_note = ("Flag ANY answer that looks even slightly off, unclear, or potentially wrong."
        if strictness == "flag_all" else "Only flag answers that are clearly and definitively wrong.")
    resp = client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[{"role":"user","content":[
            {"type":"image_url","image_url":{"url":f"data:image/png;base64,{b64}"}},
            {"type":"text","text":(
                "Look at this screenshot carefully. Find EVERY question visible on screen.\n\n"
                "CRITICAL — HOW TO IDENTIFY QUESTIONS:\n"
                "- A number inside a GREEN CIRCLE or any colored/highlighted circle (like ① ② ③) is ALWAYS a question number\n"
                "- Any diagram, triangle, geometric shape, or figure that appears NEXT TO or BELOW a circled/numbered label is part of that question\n"
                "- Plain numbers: '1.', '2.', 'Q1', 'a)', 'b)' — all questions\n"
                "- A geometric diagram (triangle with angles/sides labeled) next to a number = geometry question, describe it fully\n"
                "- NEVER skip a numbered item just because it contains a diagram instead of text\n"
                "- Look for question marks, fill-in blanks, answer boxes\n"
                "- Include ALL questions even if already answered\n\n"
                "For EACH question:\n"
                "1. question_label: the number/label (e.g. '1.', 'Q2', 'a)') — null if not visible\n"
                "2. question: full question text. If question involves a diagram, describe it (e.g. 'Triangle with angles 97 28 55 degrees — Order sides shortest to longest')\n"
                "3. type: MULTIPLE_CHOICE, TRUE_FALSE, or WRITTEN\n"
                "4. answered: true if person has written/selected an answer\n"
                "5. user_answer: what they wrote/selected (null if unanswered)\n"
                "6. correct_answer: the correct answer\n"
                "7. is_correct: true/false/null (null if unanswered)\n"
                f"8. {strict_note}\n"
                "9. correction: if wrong, correction in EXACT SAME FORMAT as their answer\n"
                "10. confident: false only if you genuinely cannot read the content\n"
                + ("11. Where helpful, include a brief worked example showing HOW to get the correct answer.\n\n"
                   if examine_examples else
                   "11. Do NOT include worked examples or sample problems — answer only.\n\n")
                + "Return ONLY valid JSON array (no markdown):\n"
                '[{"question_label":"1.","question":"What is...","type":"WRITTEN",'
                '"answered":false,"user_answer":null,"correct_answer":"Paris",'
                '"is_correct":null,"correction":null,"confident":true}]\n'
                "If no questions found: []"
            )}
        ]}],
        max_tokens=3000, temperature=0.1
    )
    raw = re.sub(r"```(?:json)?|```", "", resp.choices[0].message.content.strip()).strip()
    m = re.search(r'\[.*\]', raw, re.DOTALL)
    if m: raw = m.group(0)
    return json.loads(raw)

def ai_math(b64, examine_examples=False):
    client = get_client()
    resp = client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[{"role":"user","content":[
            {"type":"image_url","image_url":{"url":f"data:image/png;base64,{b64}"}},
            {"type":"text","text":(
                "Look at this screenshot. Find EVERY math or geometry problem visible.\n\n"
                "CRITICAL — HOW TO IDENTIFY PROBLEMS:\n"
                "- A number inside a GREEN CIRCLE or any colored circle (① ② ③) = problem number\n"
                "- Any diagram, triangle, or geometric figure NEXT TO or BELOW a circled number = that IS the problem\n"
                "- Plain numbers '1.', '2.', 'Q1' etc followed by text or a diagram = problems\n"
                "- NEVER skip a numbered item just because it shows a diagram\n"
                "- Geometry diagrams with angle measures, side lengths, or shape labels are ALL math problems\n\n"
                "For EACH problem provide ALL fields:\n"
                "- problem_label: number/label if visible (e.g. '1.', 'Q3') — null if not visible\n"
                "- problem: exact problem text. If it is a diagram problem describe it fully\n"
                "  (e.g. 'Triangle LJK with angles: L=97 degrees, J=28 degrees, K=55 degrees. Order the sides from shortest to longest.')\n"
                "- answer: the final answer clearly stated\n"
                "- steps: array of step-by-step solution strings\n"
                "- explanations: array of plain English explanation per step (SAME length as steps)\n"
                "- vertical_method: vertical/column layout as multi-line string if applicable (long division, column addition etc). null otherwise\n"
                "- graph_eq: for ANY plottable function provide 'y=expr'. For geometry problems that can be visualized provide null\n"
                "- diagram_description: for geometry/diagram problems, describe what to draw (e.g. 'Triangle with angles 97, 28, 55 degrees'). null for pure algebra\n"
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
                '[{"problem_label":"1.","problem":"2x+3=7",'
                '"steps":["Subtract 3: 2x=4","Divide by 2: x=2"],'
                '"explanations":["Remove constant from left","Isolate x"],'
                '"vertical_method":null,"answer":"x = 2",'
                '"graph_eq":null,"diagram_description":null,"has_graph":false,"confident":true}]\n'
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

def ai_qa(question, history, is_followup):
    client = get_client()
    if is_followup and history:
        messages = list(history) + [{"role":"user","content":f"{question}\n\n[Follow-up. Be more detailed. 2-4 sentences max.]"}]
        max_tok = 400
    else:
        messages = [
            {"role":"system","content":"You are a sharp, direct assistant. Answer in ONE sentence, max 20 words. No intros, no filler. Just the answer."},
            {"role":"user","content":question}
        ]
        max_tok = 120
    resp = client.chat.completions.create(model="llama-3.3-70b-versatile", messages=messages, max_tokens=max_tok, temperature=0.4)
    answer = resp.choices[0].message.content.strip()
    if not is_followup:
        new_history = [{"role":"system","content":"You are a helpful, knowledgeable assistant."},{"role":"user","content":question},{"role":"assistant","content":answer}]
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

_AOT_HWND   = None   # cached window handle
_aot_thread = None   # background watcher thread

HWND_TOPMOST    = -1
HWND_NOTOPMOST  = -2
SWP_NOMOVE      = 0x0002
SWP_NOSIZE      = 0x0001
SWP_NOACTIVATE  = 0x0010
SWP_FLAGS       = SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE

def _find_app_hwnd():
    """
    Find our app window. Matches on title OR on the process name being
    msedge/chrome running our localhost URL (Edge app-mode strips the title).
    """
    if not HAS_WIN32:
        return None

    TITLE_HINTS   = ("AutoTyper", "Fiuxxed", "127.0.0.1:7890", "localhost:7890")
    # Never grab these — they're console/system windows
    SKIP_TITLES   = ("cmd", "command prompt", "powershell", "python", "administrator",
                     "c:\\windows", "conhost", "run_hidden")
    result = [None]
    best_score = [0]

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if not title:
            return
        tl = title.lower()
        # Hard skip console/system windows
        if any(s in tl for s in SKIP_TITLES):
            return
        score = 0
        for hint in TITLE_HINTS:
            if hint in title:
                score += 2 if hint in ("AutoTyper","Fiuxxed") else 1
        if score == 0:
            return
        # Prefer windows in the right size range for our app
        try:
            rect = win32gui.GetWindowRect(hwnd)
            w = rect[2] - rect[0]
            if 300 < w < 900:
                score += 1
        except Exception:
            pass
        if score > best_score[0]:
            best_score[0] = score
            result[0] = hwnd

    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass

    # Fallback: find by process — look for msedge/chrome with our port in cmdline
    if result[0] is None and HAS_WIN32:
        try:
            for proc in psutil.process_iter(["pid","name","cmdline"]):
                name = (proc.info.get("name") or "").lower()
                if "edge" not in name and "chrome" not in name:
                    continue
                cmdline = " ".join(proc.info.get("cmdline") or [])
                if "7890" not in cmdline and "AutoTyper" not in cmdline:
                    continue
                def pid_cb(hwnd, pid):
                    if not win32gui.IsWindowVisible(hwnd): return
                    try:
                        _, wpid = win32process.GetWindowThreadProcessId(hwnd)
                        if wpid == pid and win32gui.GetWindowText(hwnd):
                            result[0] = hwnd
                    except Exception: pass
                win32gui.EnumWindows(pid_cb, proc.info["pid"])
                if result[0]: break
        except Exception:
            pass

    # Final safety check — make sure the handle belongs to msedge or chrome process
    # This prevents ever accidentally operating on Explorer or system windows
    if result[0] and HAS_WIN32:
        try:
            _, pid = win32process.GetWindowThreadProcessId(result[0])
            proc_name = ""
            for p in psutil.process_iter(["pid","name"]):
                if p.info["pid"] == pid:
                    proc_name = (p.info["name"] or "").lower()
                    break
            if "edge" not in proc_name and "chrome" not in proc_name and "python" not in proc_name:
                return None  # refuse to operate on non-browser windows
        except Exception:
            pass

    return result[0]

def _set_hwnd_topmost(hwnd, on_top):
    """Use raw win32 SetWindowPos to pin/unpin a window."""
    try:
        insert = HWND_TOPMOST if on_top else HWND_NOTOPMOST
        ctypes.windll.user32.SetWindowPos(hwnd, insert, 0, 0, 0, 0, SWP_FLAGS)
    except Exception:
        pass

def apply_always_on_top(val):
    """Called whenever the setting changes — immediately applies it."""
    global _webview_window, _AOT_HWND
    on_top = bool(val)

    # pywebview path
    if _webview_window:
        try:
            _webview_window.on_top = on_top
        except Exception:
            pass

    # win32 direct path — works for Edge app mode too
    if HAS_WIN32:
        hwnd = _AOT_HWND or _find_app_hwnd()
        if hwnd:
            _AOT_HWND = hwnd
            _set_hwnd_topmost(hwnd, on_top)

def _aot_watcher():
    """
    Background thread — re-applies always_on_top every second.
    Always re-finds the window instead of trusting a cached handle,
    so if Edge recreates its window or the handle goes stale we self-correct.
    """
    global _AOT_HWND
    time.sleep(2.5)
    while True:
        try:
            if bool(cfg.get("always_on_top", True)):
                # Always re-find — never trust stale cache for AOT
                hwnd = _find_app_hwnd()
                if hwnd:
                    _AOT_HWND = hwnd
                    _set_hwnd_topmost(hwnd, True)
        except Exception:
            pass
        time.sleep(1)

def start_aot_watcher():
    global _aot_thread
    _aot_thread = threading.Thread(target=_aot_watcher, daemon=True)
    _aot_thread.start()

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
    _type_state.update({"phase":"typing","progress":0,"status":"Typing..."})
    def on_prog(idx, total, pct): _type_state.update({"progress":pct,"status":f"Typing… {pct}%  ({idx}/{total} chars)"})
    def on_done(): _type_state.update({"phase":"done","progress":100,"status":"✓ All done!"})
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
    wins = []
    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd): return
        title = win32gui.GetWindowText(hwnd)
        if not title or len(title) < 2: return
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            name = psutil.Process(pid).name().lower().replace(".exe","")
        except Exception: name = "unknown"
        wins.append({"hwnd":hwnd,"title":title,"name":name,"is_browser":any(b in name for b in BROWSERS)})
    win32gui.EnumWindows(cb, None)
    wins.sort(key=lambda x: x["title"].lower())
    _windows_cache = wins
    return jsonify({"windows": wins})

@app_flask.route("/api/screenshot", methods=["POST"])
def api_screenshot():
    data = request.json or {}
    hwnd = data.get("hwnd"); mode = data.get("mode","scan")
    strictness = data.get("strictness", cfg.get("scanner_wrong_answer_strictness","flag_all"))
    try:
        img = capture_window(hwnd); b64 = img_to_b64(img)
        if mode == "math":
            problems = ai_math(b64, examine_examples=cfg.get("examine_examples", False))
            for p in problems:
                eq = p.get("graph_eq")
                p["graph_b64"] = make_graph(eq) if (eq and cfg.get("math_show_graphs", True)) else None
                if "diagram_description" not in p:
                    p["diagram_description"] = None
                if cfg.get("math_formula_library", True) and p.get("answer"):
                    add_to_formula_lib(p.get("problem",""), p.get("answer",""), p.get("steps",[]))
            add_to_history("math", problems)
            return jsonify({"ok":True,"result":{"problems":problems}})
        else:
            questions = ai_scan(b64, strictness, examine_examples=cfg.get("examine_examples", False))
            add_to_history("scan", questions)
            return jsonify({"ok":True,"result":{"questions":questions}})
    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()})

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
    if not question: return jsonify({"error":"No question"})
    try:
        answer, new_history = ai_qa(question, history, is_followup)
        return jsonify({"answer":answer,"history":new_history})
    except Exception as e: return jsonify({"error":str(e)})

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

@app_flask.route("/api/music/search")
def api_music_search():
    """Search YouTube via yt-dlp (reliable) or Invidious fallback."""
    query = request.args.get("q", "").strip()
    if not query:
        return jsonify({"error": "No query", "results": []})

    # ── Primary: yt-dlp ──────────────────────────────────────
    if HAS_YTDLP:
        try:
            ydl_opts = {
                "quiet": True,
                "no_warnings": True,
                "extract_flat": "in_playlist",
                "default_search": "ytsearch15",
                "skip_download": True,
                "ignoreerrors": True,
            }
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(f"ytsearch15:{query}", download=False)
            entries = (info or {}).get("entries") or []
            results = []
            for v in entries[:15]:
                if not v:
                    continue
                vid = v.get("id") or v.get("url","")
                thumb = v.get("thumbnail","")
                if not thumb and vid:
                    thumb = f"https://i.ytimg.com/vi/{vid}/mqdefault.jpg"
                results.append({
                    "videoId":  vid,
                    "title":    v.get("title","Untitled"),
                    "author":   v.get("uploader") or v.get("channel",""),
                    "duration": v.get("duration") or 0,
                    "thumb":    thumb,
                })
            if results:
                return jsonify({"results": results})
        except Exception as e:
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
            url = (f"{base}/api/v1/search?q={urllib.parse.quote(query)}"
                   f"&type=video&fields=videoId,title,author,lengthSeconds,videoThumbnails")
            req = urllib.request.Request(url, headers={"User-Agent":"Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=5, context=ctx) as resp:
                data = json.loads(resp.read().decode())
            if not isinstance(data, list) or not data:
                continue
            results = []
            for v in data[:15]:
                thumbs = v.get("videoThumbnails",[])
                thumb = next((t["url"] for t in thumbs if t.get("quality")=="medium"),
                             thumbs[0]["url"] if thumbs else "")
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
                return jsonify({"results": results})
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
    if not HAS_WIN32: return jsonify({"ok": False})
    hwnd = _AOT_HWND or _find_app_hwnd()
    if hwnd:
        try:
            ctypes.windll.user32.PostMessageW(hwnd, WM_SYSCOMMAND, SC_CLOSE, 0)
        except Exception:
            pass
    # Also shut down Flask so the process fully exits
    threading.Thread(target=lambda: (time.sleep(0.3), os._exit(0)), daemon=True).start()
    return jsonify({"ok": True})

# ══════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════
import socket, subprocess

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

def _find_browser():
    import winreg
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
        ])
        # Strip the native titlebar via SetWindowLong after window appears
        threading.Thread(target=_strip_titlebar_later, daemon=True).start()
        return proc
    except Exception as e: print("Browser launch error: "+str(e)); return None

# Win32 constants for stripping the titlebar
GWL_STYLE      = -16
GWL_EXSTYLE    = -20
WS_CAPTION     = 0x00C00000  # titlebar (WS_BORDER | WS_DLGFRAME)
WS_THICKFRAME  = 0x00040000  # resizable border — DO NOT REMOVE (breaks resize)
WS_BORDER      = 0x00800000
WS_DLGFRAME    = 0x00400000
WS_SYSMENU     = 0x00080000  # system menu (also part of caption area)
WS_EX_TOOLWINDOW = 0x00000080  # hides from taskbar, removes caption chrome

def _apply_frame_strip(hwnd):
    """
    Remove WS_CAPTION (titlebar) only — keep WS_THICKFRAME (resize border).
    Do NOT set WS_EX_TOOLWINDOW — that hides the app from the taskbar.
    """
    SWP_FRAMECHANGED = 0x0020
    try:
        style     = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_STYLE)
        new_style = style & ~WS_CAPTION & ~WS_SYSMENU
        if new_style == style:
            return False  # already stripped, nothing to do
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_STYLE, new_style)
        ctypes.windll.user32.SetWindowPos(
            hwnd, 0, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED
        )
        return True
    except Exception:
        return False

_strip_done = False   # becomes True once we've successfully stripped

def _strip_titlebar_later():
    """
    Persistent loop that:
    1. Waits for the Edge window to appear
    2. Strips the native titlebar
    3. Keeps re-stripping every 500ms for the first 10s (Edge redraws its
       frame after page-load, which puts the bar back — we catch that)
    4. Then checks every 3s forever in case the window was recreated
    """
    global _AOT_HWND, _strip_done
    if not HAS_WIN32:
        return

    # Phase 1: wait for window (up to 15s)
    deadline = time.time() + 15
    hwnd = None
    while time.time() < deadline:
        hwnd = _find_app_hwnd()
        if hwnd:
            break
        time.sleep(0.3)

    if not hwnd:
        return

    _AOT_HWND = hwnd

    # Phase 2: aggressively re-strip for first 12 seconds (catches Edge redraw)
    strip_deadline = time.time() + 12
    while time.time() < strip_deadline:
        try:
            cur_hwnd = _find_app_hwnd() or hwnd
            if cur_hwnd:
                _AOT_HWND = cur_hwnd
                _apply_frame_strip(cur_hwnd)
                on_top = bool(cfg.get("always_on_top", True))
                _set_hwnd_topmost(cur_hwnd, on_top)
        except Exception:
            pass
        time.sleep(0.5)

    _strip_done = True

    # Phase 3: periodic maintenance forever
    while True:
        try:
            cur_hwnd = _find_app_hwnd()
            if cur_hwnd:
                _AOT_HWND = cur_hwnd
                _apply_frame_strip(cur_hwnd)
                on_top = bool(cfg.get("always_on_top", True))
                _set_hwnd_topmost(cur_hwnd, on_top)
        except Exception:
            pass
        time.sleep(3)

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
    start_hotkeys()
    start_aot_watcher()   # always-on-top background watcher
    flask_thread = threading.Thread(target=run_flask, daemon=True); flask_thread.start()
    if not _wait_for_flask(): print("ERROR: Flask server failed to start."); sys.exit(1)
    url = "http://127.0.0.1:7890"
    proc = _launch_app_window(url)
    if proc: _keep_alive(proc); return
    if HAS_WEBVIEW:
        try:
            global _webview_window
            _webview_window = webview.create_window(title="Fiuxxed's AutoTyper", url=url,
                width=600, height=940, min_size=(420,600), resizable=True,
                on_top=bool(cfg.get("always_on_top",True)), background_color="#07070e")
            webview.start(debug=False); return
        except Exception as e: print("pywebview error: "+str(e))
    import webbrowser; print("Opening in browser: "+url); webbrowser.open(url); _keep_alive()

if __name__ == "__main__":
    main()
