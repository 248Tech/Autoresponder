"""
AutoResponder Bot - Windows AI-powered typing assistant
With Screen Capture (OCR) + AutoType

Requirements:
    pip install openai pyperclip pynput keyboard customtkinter pygetwindow
    pip install pillow pytesseract mss pyautogui

Also install Tesseract OCR engine:
    Download: https://github.com/UB-Mannheim/tesseract/wiki
    Install to: C:/Program Files/Tesseract-OCR/tesseract.exe
"""

import tkinter as tk
import customtkinter as ctk
import threading
import time
import json
import os
import pyperclip
import keyboard
from openai import OpenAI
from pynput.keyboard import Controller, Key
import queue
from PIL import Image, ImageTk, ImageGrab
import mss
import mss.tools
import pytesseract
import pyautogui
import ctypes

# ── Tesseract path (adjust if installed elsewhere) ───────────────────────
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ── Config ───────────────────────────────────────────────────────────────
CONFIG_FILE = "autoresponder_config.json"
DEFAULT_CONFIG = {
    "api_key": "",
    "model": "gpt-4o-mini",
    "hotkey": "ctrl+shift+space",
    "capture_hotkey": "ctrl+shift+c",
    "tone": "Friendly",
    "auto_send": False,
    "delay_ms": 500,
    "preset_messages": [
        "Thanks for reaching out! I'll get back to you shortly.",
        "Great question! Let me look into that for you.",
        "I appreciate your patience. Here's what I found:",
        "Happy to help! Here's what you need to know:",
    ],
    "platform": "Generic",
    "context_prompt": "",
    "typing_speed_cps": 30,
    "autotype_enabled": True,
    "capture_region": None,  # [x, y, w, h] or None for full screen
    "capture_monitor": "all",  # all | primary | monitor:<index>
    "watch_interval_ms": 1500,
    "detect_existing_text": True,
}

TONES = ["Friendly", "Professional", "Casual", "Formal", "Empathetic", "Concise", "Enthusiastic", "Technical"]
PLATFORMS = ["Generic", "Discord", "Slack", "WhatsApp Web", "Teams", "Gmail", "Telegram Web"]

TONE_DESCRIPTIONS = {
    "Friendly":     "warm, approachable, and personable",
    "Professional": "formal, polished, and business-appropriate",
    "Casual":       "relaxed, conversational, and informal",
    "Formal":       "strictly formal with proper etiquette",
    "Empathetic":   "understanding, compassionate, and supportive",
    "Concise":      "brief, direct, and to the point",
    "Enthusiastic": "energetic, positive, and upbeat",
    "Technical":    "precise, detailed, and technically accurate",
}

PLATFORM_HINTS = {
    "Generic":      "Keep response suitable for any text input field.",
    "Discord":      "Use Discord markdown if helpful. Keep responses chat-length.",
    "Slack":        "Use Slack formatting if helpful. Keep professional but conversational.",
    "WhatsApp Web": "Keep it short and conversational, no markdown.",
    "Teams":        "Professional tone appropriate for Microsoft Teams.",
    "Gmail":        "Format as a proper email reply with greeting and sign-off.",
    "Telegram Web": "Short, conversational reply.",
}

keyboard_controller = Controller()

# ── OpenAI ────────────────────────────────────────────────────────────────
def generate_reply(prompt: str, cfg: dict) -> str:
    client = OpenAI(api_key=cfg["api_key"])
    tone, platform = cfg["tone"], cfg["platform"]
    sys_prompt = (
        f"You are an AI assistant helping the user auto-respond to messages.\n"
        f"Tone: {tone} — {TONE_DESCRIPTIONS.get(tone, '')}\n"
        f"Platform: {platform} — {PLATFORM_HINTS.get(platform, '')}\n"
        f"{'Additional context: ' + cfg['context_prompt'] if cfg['context_prompt'] else ''}\n"
        "Write ONLY the reply text. No explanations, no meta-commentary."
    )
    resp = client.chat.completions.create(
        model=cfg["model"],
        messages=[
            {"role": "system", "content": sys_prompt},
            {"role": "user",   "content": f"Generate a reply to this message:\n\n{prompt}"},
        ],
        max_tokens=300,
    )
    return resp.choices[0].message.content.strip()

def refine_reply(existing_reply: str, instruction: str, cfg: dict) -> str:
    client = OpenAI(api_key=cfg["api_key"])
    tone, platform = cfg["tone"], cfg["platform"]
    sys_prompt = (
        "You rewrite draft replies.\n"
        f"Tone: {tone} - {TONE_DESCRIPTIONS.get(tone, '')}\n"
        f"Platform: {platform} - {PLATFORM_HINTS.get(platform, '')}\n"
        f"{'Additional context: ' + cfg['context_prompt'] if cfg['context_prompt'] else ''}\n"
        "Return ONLY the revised reply text."
    )
    resp = client.chat.completions.create(
        model=cfg["model"],
        messages=[
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": f"Instruction: {instruction}\n\nDraft reply:\n{existing_reply}"},
        ],
        max_tokens=320,
    )
    return resp.choices[0].message.content.strip()

def type_text(text: str, cps: int = 30):
    delay = 1.0 / max(cps, 1)
    for ch in text:
        keyboard_controller.type(ch)
        time.sleep(delay)

# ── Config I/O ────────────────────────────────────────────────────────────
def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE) as f:
            saved = json.load(f)
        return {**DEFAULT_CONFIG, **saved}
    return DEFAULT_CONFIG.copy()

def save_config(cfg: dict):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)

# ── Region Selector Overlay ───────────────────────────────────────────────
class RegionSelector(tk.Toplevel):
    """Transparent overlay for drag-to-select a capture region."""
    def __init__(self, master, callback, bounds=None):
        super().__init__(master)
        self.callback = callback
        self.start = None
        self.rect = None
        self.bounds = bounds or {
            "left": 0,
            "top": 0,
            "width": self.winfo_screenwidth(),
            "height": self.winfo_screenheight(),
        }
        self.left = int(self.bounds.get("left", 0))
        self.top = int(self.bounds.get("top", 0))
        self.width = int(self.bounds.get("width", self.winfo_screenwidth()))
        self.height = int(self.bounds.get("height", self.winfo_screenheight()))

        self.overrideredirect(True)
        self.geometry(f"{self.width}x{self.height}+{self.left}+{self.top}")
        self.attributes("-alpha", 0.3)
        self.attributes("-topmost", True)
        self.configure(bg="black", cursor="crosshair")

        self.canvas = tk.Canvas(self, bg="black", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.create_text(
            self.width // 2, 40,
            text="Click and drag to select capture region. Press Esc to cancel.",
            fill="white", font=("Segoe UI", 16)
        )

        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Escape>", lambda e: self.destroy())

    def _on_press(self, e):
        self.start = (e.x, e.y)
        if self.rect:
            self.canvas.delete(self.rect)

    def _on_drag(self, e):
        if self.rect:
            self.canvas.delete(self.rect)
        x0, y0 = self.start
        self.rect = self.canvas.create_rectangle(
            x0, y0, e.x, e.y, outline="#00d4ff", width=2, fill="#00d4ff", stipple="gray25"
        )

    def _on_release(self, e):
        x0, y0 = self.start
        x1, y1 = e.x, e.y
        w = abs(x1 - x0)
        h = abs(y1 - y0)
        if w < 3 or h < 3:
            self.destroy()
            return
        region = [min(x0, x1) + self.left, min(y0, y1) + self.top, w, h]
        self.destroy()
        self.callback(region)

# ── Screenshot Preview Window ─────────────────────────────────────────────
class ScreenshotViewer(tk.Toplevel):
    def __init__(self, master, img: Image.Image, ocr_text: str, on_use):
        super().__init__(master)
        self.title("Captured Screen Region")
        self.attributes("-topmost", True)
        self.resizable(True, True)

        # Thumbnail
        thumb = img.copy()
        thumb.thumbnail((600, 300))
        self.photo = ImageTk.PhotoImage(thumb)
        tk.Label(self, image=self.photo, bg="#1a1a2e").pack(padx=10, pady=10)

        # OCR text
        ctk.CTkLabel(self, text="Extracted Text (editable):", anchor="w").pack(fill="x", padx=10)
        self.text_box = ctk.CTkTextbox(self, height=120, font=("Consolas", 12))
        self.text_box.pack(fill="both", expand=True, padx=10)
        self.text_box.insert("1.0", ocr_text)

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(pady=8)
        ctk.CTkButton(btn_row, text="✅ Use as Input", fg_color="#0078d4",
                      command=lambda: on_use(self.text_box.get("1.0","end").strip())).pack(side="left", padx=6)
        ctk.CTkButton(btn_row, text="❌ Discard", fg_color="#444",
                      command=self.destroy).pack(side="left", padx=6)


# ── AutoType Preview ──────────────────────────────────────────────────────
class AutoTypePreview(ctk.CTkToplevel):
    """Shows reply with countdown before typing begins."""
    def __init__(self, master, text: str, delay_s: float, on_confirm, on_cancel):
        super().__init__(master)
        self.title("⌨️ AutoType Preview")
        self.attributes("-topmost", True)
        self.geometry("480x320")
        self.on_confirm = on_confirm
        self.on_cancel = on_cancel
        self.cancelled = False

        ctk.CTkLabel(self, text="Reply will be typed in:", font=("Segoe UI", 13)).pack(pady=(16,4))
        self.countdown_lbl = ctk.CTkLabel(self, text=f"{delay_s:.1f}s", font=("Segoe UI", 36, "bold"),
                                           text_color="#00d4ff")
        self.countdown_lbl.pack()

        ctk.CTkLabel(self, text="Text to type:", anchor="w").pack(fill="x", padx=14, pady=(10,2))
        self.preview = ctk.CTkTextbox(self, height=100, font=("Segoe UI", 12))
        self.preview.pack(fill="both", expand=True, padx=14)
        self.preview.insert("1.0", text)

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(pady=10)
        ctk.CTkButton(btn_row, text="⌨️ Type Now", fg_color="#107c10",
                      command=self._confirm_now).pack(side="left", padx=8)
        ctk.CTkButton(btn_row, text="❌ Cancel", fg_color="#c00",
                      command=self._cancel).pack(side="left", padx=8)

        self._remaining = delay_s
        self._tick()

    def _tick(self):
        if self.cancelled or not self.winfo_exists():
            return
        if self._remaining <= 0:
            self._confirm_now()
            return
        self.countdown_lbl.configure(text=f"{self._remaining:.1f}s")
        self._remaining = round(self._remaining - 0.1, 1)
        self.after(100, self._tick)

    def _confirm_now(self):
        text = self.preview.get("1.0", "end").strip()
        if not self.cancelled:
            self.cancelled = True
            self.destroy()
            self.on_confirm(text)

    def _cancel(self):
        self.cancelled = True
        self.destroy()
        self.on_cancel()


# ── Main App ──────────────────────────────────────────────────────────────
class AutoResponderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.cfg = load_config()
        self.reply_queue = queue.Queue()
        self.captured_region = self.cfg.get("capture_region")
        self.monitor_targets = self._discover_monitor_targets()
        self.cfg.setdefault("capture_monitor", "all")
        if self.cfg["capture_monitor"] not in self.monitor_targets:
            self.cfg["capture_monitor"] = "all"
        self._auto_generate_after_capture = False
        self.watch_running = False
        self.watch_stop_event = threading.Event()
        self.watch_thread = None
        self.last_detected_text = ""
        self.watch_detect_existing = bool(self.cfg.get("detect_existing_text", True))
        self.watch_interval_s = max(0.3, int(self.cfg.get("watch_interval_ms", 1500)) / 1000.0)
        self._hotkey_hooks = []

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title("🤖 AutoResponder Bot")
        self.geometry("820x760")
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self._build_ui()
        self._register_hotkeys()
        self.after(200, self._poll_queue)

    # ── UI ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        hdr = ctk.CTkFrame(self, corner_radius=0, height=60, fg_color="#1a1a2e")
        hdr.pack(fill="x")
        ctk.CTkLabel(hdr, text="🤖  AutoResponder Bot", font=("Segoe UI", 20, "bold"),
                     text_color="#00d4ff").pack(side="left", padx=20, pady=10)
        self.status_dot = ctk.CTkLabel(hdr, text="● IDLE", font=("Segoe UI", 12), text_color="#888")
        self.status_dot.pack(side="right", padx=20)

        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=10)
        for t in ["Compose", "Screen Capture", "AutoType", "Presets", "Settings", "Log"]:
            self.tabs.add(t)

        self._build_compose_tab()
        self._build_capture_tab()
        self._build_autotype_tab()
        self._build_presets_tab()
        self._build_settings_tab()
        self._build_log_tab()
        self._bind_live_counters()
        self._update_all_stats()

    # ── Compose Tab ───────────────────────────────────────────────────────
    def _build_compose_tab(self):
        tab = self.tabs.tab("Compose")
        ctk.CTkLabel(tab, text="Incoming message:", anchor="w").pack(fill="x", padx=5, pady=(10,2))
        self.input_box = ctk.CTkTextbox(tab, height=110, font=("Segoe UI", 13))
        self.input_box.pack(fill="x", padx=5)
        self.input_stats_lbl = ctk.CTkLabel(tab, text="0 chars - 0 words", text_color="#8f9bb3")
        self.input_stats_lbl.pack(anchor="e", padx=8, pady=(2, 0))

        row = ctk.CTkFrame(tab, fg_color="transparent")
        row.pack(fill="x", padx=5, pady=6)
        ctk.CTkLabel(row, text="Tone:").pack(side="left")
        self.tone_var = ctk.StringVar(value=self.cfg["tone"])

        ctk.CTkOptionMenu(row, values=TONES, variable=self.tone_var,
                          command=lambda v: None, width=150).pack(side="left", padx=8)
        ctk.CTkLabel(row, text="Platform:").pack(side="left", padx=(16,0))
        self.platform_var = ctk.StringVar(value=self.cfg["platform"])

        ctk.CTkOptionMenu(row, values=PLATFORMS, variable=self.platform_var,
                          width=150).pack(side="left", padx=8)

        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.pack(fill="x", padx=5, pady=4)
        ctk.CTkButton(btn_row, text="⚡ Generate Reply", command=self._generate_clicked,
                      fg_color="#0078d4", hover_color="#005a9e", width=160).pack(side="left")
        ctk.CTkButton(btn_row, text="📋 Paste Clipboard", command=self._paste_from_clipboard,
                      width=150).pack(side="left", padx=8)
        ctk.CTkButton(btn_row, text="🗑 Clear", command=self._clear_all,
                      fg_color="#444", width=80).pack(side="left")

        ctk.CTkLabel(tab, text="Generated reply:", anchor="w").pack(fill="x", padx=5, pady=(10,2))
        self.output_box = ctk.CTkTextbox(tab, height=140, font=("Segoe UI", 13))
        self.output_box.pack(fill="x", padx=5)
        self.output_stats_lbl = ctk.CTkLabel(tab, text="0 chars - 0 words", text_color="#8f9bb3")
        self.output_stats_lbl.pack(anchor="e", padx=8, pady=(2, 0))

        action_row = ctk.CTkFrame(tab, fg_color="transparent")
        action_row.pack(fill="x", padx=5, pady=8)
        ctk.CTkButton(action_row, text="⌨️ AutoType Reply",
                      command=self._launch_autotype_preview,
                      fg_color="#107c10", hover_color="#0a5c0a", width=160).pack(side="left")
        ctk.CTkButton(action_row, text="📋 Copy Reply", command=self._copy_reply,
                      width=120).pack(side="left", padx=8)
        self.auto_send_var = ctk.BooleanVar(value=self.cfg["auto_send"])

        ctk.CTkCheckBox(action_row, text="Auto-send after typing",
                        variable=self.auto_send_var).pack(side="left", padx=10)

        refine_row = ctk.CTkFrame(tab, fg_color="transparent")
        refine_row.pack(fill="x", padx=5, pady=(0, 4))
        ctk.CTkLabel(refine_row, text="Polish reply:", text_color="#9aa5b1").pack(side="left")
        ctk.CTkButton(refine_row, text="Shorter", width=90,
                      command=lambda: self._refine_clicked("Make this shorter and clearer."),
                      fg_color="#2d5bff", hover_color="#2549cc").pack(side="left", padx=6)
        ctk.CTkButton(refine_row, text="Friendlier", width=90,
                      command=lambda: self._refine_clicked("Make this warmer and more friendly."),
                      fg_color="#2d5bff", hover_color="#2549cc").pack(side="left", padx=6)
        ctk.CTkButton(refine_row, text="More Formal", width=100,
                      command=lambda: self._refine_clicked("Make this more professional and formal."),
                      fg_color="#2d5bff", hover_color="#2549cc").pack(side="left", padx=6)

        ctk.CTkLabel(tab, text=f"Hotkeys: Generate+Type={self.cfg['hotkey']}  |  Capture={self.cfg['capture_hotkey']}",
                     font=("Segoe UI", 11), text_color="#888").pack(pady=4)

    # ── Screen Capture Tab ────────────────────────────────────────────────
    def _build_capture_tab(self):
        tab = self.tabs.tab("Screen Capture")

        reg_frame = ctk.CTkFrame(tab, corner_radius=8)
        reg_frame.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(reg_frame, text="📐 Capture Region", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=6)

        target_row = ctk.CTkFrame(reg_frame, fg_color="transparent")
        target_row.pack(fill="x", padx=10, pady=(0, 4))
        ctk.CTkLabel(target_row, text="Screen target:").pack(side="left")
        self.capture_target_names = {v: k for k, v in self.monitor_targets.items()}
        current_target = self.monitor_targets.get(self.cfg.get("capture_monitor", "all"), "All Monitors")
        self.capture_target_var = ctk.StringVar(value=current_target)
        self.capture_target_menu = ctk.CTkOptionMenu(
            target_row,
            values=list(self.monitor_targets.values()),
            variable=self.capture_target_var,
            command=self._on_capture_target_change,
            width=200,
        )
        self.capture_target_menu.pack(side="left", padx=8)

        self.region_lbl = ctk.CTkLabel(reg_frame, text=self._region_text(), text_color="#aaa")
        self.region_lbl.pack(anchor="w", padx=10)

        rb = ctk.CTkFrame(reg_frame, fg_color="transparent")
        rb.pack(fill="x", padx=10, pady=8)
        ctk.CTkButton(rb, text="🖱 Select Region (drag)", command=self._pick_region,
                      width=180).pack(side="left")
        ctk.CTkButton(rb, text="🖥 Full Screen", command=self._use_full_screen,
                      width=140).pack(side="left", padx=8)
        ctk.CTkButton(rb, text="🗑 Clear Region", command=self._clear_region,
                      fg_color="#444", width=120).pack(side="left")

        cap_frame = ctk.CTkFrame(tab, corner_radius=8)
        cap_frame.pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(cap_frame, text="📸 Capture & OCR", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=6)

        cb = ctk.CTkFrame(cap_frame, fg_color="transparent")
        cb.pack(fill="x", padx=10, pady=6)
        ctk.CTkButton(cb, text="📸 Capture Now", command=self._capture_now,
                      fg_color="#0078d4", width=150).pack(side="left")
        ctk.CTkButton(cb, text="⚡ Capture + Generate", command=self._capture_and_generate,
                      fg_color="#107c10", width=180).pack(side="left", padx=8)
        ctk.CTkButton(cb, text=f"🔑 Hotkey: {self.cfg['capture_hotkey']}",
                      fg_color="#333", width=180).pack(side="left", padx=8)

        ctk.CTkLabel(cap_frame, text="OCR Result (editable - will be used as input):", anchor="w").pack(fill="x", padx=10, pady=(8,2))
        self.ocr_box = ctk.CTkTextbox(cap_frame, height=140, font=("Consolas", 12))
        self.ocr_box.pack(fill="both", expand=True, padx=10, pady=(0,8))
        self.ocr_stats_lbl = ctk.CTkLabel(cap_frame, text="0 chars - 0 words", text_color="#8f9bb3")
        self.ocr_stats_lbl.pack(anchor="e", padx=12, pady=(0, 6))

        ob = ctk.CTkFrame(cap_frame, fg_color="transparent")
        ob.pack(fill="x", padx=10, pady=(0,8))
        ctk.CTkButton(ob, text="➡️ Send to Compose & Generate",
                      command=self._ocr_to_compose, fg_color="#107c10", width=230).pack(side="left")
        ctk.CTkButton(ob, text="📋 Copy OCR Text",
                      command=lambda: pyperclip.copy(self.ocr_box.get("1.0","end").strip()),
                      width=140).pack(side="left", padx=8)

        watch_frame = ctk.CTkFrame(tab, corner_radius=8)
        watch_frame.pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(watch_frame, text="▶ Auto Watch", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=6)

        interval_row = ctk.CTkFrame(watch_frame, fg_color="transparent")
        interval_row.pack(fill="x", padx=10, pady=(0, 6))
        ctk.CTkLabel(interval_row, text="Scan interval:").pack(side="left")
        self.watch_interval_var = ctk.IntVar(value=int(self.cfg.get("watch_interval_ms", 1500)))
        ctk.CTkSlider(interval_row, from_=500, to=8000, variable=self.watch_interval_var, width=220).pack(side="left", padx=8)
        self.watch_interval_lbl = ctk.CTkLabel(interval_row, text=f"{self.watch_interval_var.get()} ms", width=80)
        self.watch_interval_lbl.pack(side="left")
        self.watch_interval_var.trace_add("write", lambda *_: self.watch_interval_lbl.configure(text=f"{self.watch_interval_var.get()} ms"))

        self.detect_existing_var = ctk.BooleanVar(value=self.cfg.get("detect_existing_text", True))
        ctk.CTkCheckBox(watch_frame, text="Detect existing text first (ignore already on-screen text at Start)",
                        variable=self.detect_existing_var).pack(anchor="w", padx=10, pady=(0, 8))

        watch_buttons = ctk.CTkFrame(watch_frame, fg_color="transparent")
        watch_buttons.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkButton(watch_buttons, text="▶ Start", command=self._start_watch,
                      fg_color="#107c10", hover_color="#0a5c0a", width=120).pack(side="left")
        ctk.CTkButton(watch_buttons, text="⏹ Stop", command=self._stop_watch,
                      fg_color="#b71c1c", hover_color="#8e1515", width=120).pack(side="left", padx=8)
        self.watch_status_lbl = ctk.CTkLabel(watch_buttons, text="Status: Stopped", text_color="#aaa")
        self.watch_status_lbl.pack(side="left", padx=12)

        self.thumb_lbl = ctk.CTkLabel(tab, text="No capture yet", text_color="#555")
        self.thumb_lbl.pack(pady=8)

    # ── AutoType Tab ──────────────────────────────────────────────────────
    def _build_autotype_tab(self):
        tab = self.tabs.tab("AutoType")

        info = ctk.CTkFrame(tab, corner_radius=8, fg_color="#1e2a1e")
        info.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(info, text="⌨️ AutoType Engine", font=("Segoe UI", 14, "bold"),
                     text_color="#4caf50").pack(anchor="w", padx=10, pady=6)
        ctk.CTkLabel(info, text="Simulates keypresses directly into any focused window.\nWorks with Discord, browsers, apps — anything that accepts keyboard input.",
                     text_color="#aaa", justify="left").pack(anchor="w", padx=10, pady=(0,8))

        # Settings
        sf = ctk.CTkFrame(tab, corner_radius=8)
        sf.pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(sf, text="⚙️ AutoType Settings", font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=10, pady=6)

        def srow(label, widget_fn):
            r = ctk.CTkFrame(sf, fg_color="transparent")
            r.pack(fill="x", padx=10, pady=4)
            ctk.CTkLabel(r, text=label, width=220, anchor="w").pack(side="left")
            widget_fn(r)

        self.autotype_enabled_var = ctk.BooleanVar(value=self.cfg.get("autotype_enabled", True))
        srow("Enable AutoType:", lambda p: ctk.CTkSwitch(p, variable=self.autotype_enabled_var,
             text="").pack(side="left"))

        self.cps_var = ctk.IntVar(value=self.cfg["typing_speed_cps"])
        def cps_widget(p):
            ctk.CTkSlider(p, from_=5, to=120, variable=self.cps_var, width=200).pack(side="left")
            self.cps_lbl = ctk.CTkLabel(p, text=f"{self.cps_var.get()} cps", width=60)
            self.cps_lbl.pack(side="left", padx=6)
            self.cps_var.trace_add("write", lambda *_: self.cps_lbl.configure(text=f"{self.cps_var.get()} cps"))
        srow("Typing Speed:", cps_widget)

        self.delay_var = ctk.IntVar(value=self.cfg["delay_ms"])
        def delay_widget(p):
            ctk.CTkSlider(p, from_=0, to=5000, variable=self.delay_var, width=200).pack(side="left")
            self.delay_lbl = ctk.CTkLabel(p, text=f"{self.delay_var.get()} ms", width=70)
            self.delay_lbl.pack(side="left", padx=6)
            self.delay_var.trace_add("write", lambda *_: self.delay_lbl.configure(text=f"{self.delay_var.get()} ms"))
        srow("Pre-type Delay:", delay_widget)

        self.preview_before_var = ctk.BooleanVar(value=True)
        srow("Show Preview Window:", lambda p: ctk.CTkSwitch(p, variable=self.preview_before_var,
             text="").pack(side="left"))

        self.auto_send_var2 = ctk.BooleanVar(value=self.cfg["auto_send"])
        srow("Press Enter after typing:", lambda p: ctk.CTkSwitch(p, variable=self.auto_send_var2,
             text="").pack(side="left"))

        # Manual type box
        mf = ctk.CTkFrame(tab, corner_radius=8)
        mf.pack(fill="both", expand=True, padx=10, pady=6)
        ctk.CTkLabel(mf, text="🔤 Manual AutoType — type any text into the active window",
                     font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=10, pady=6)
        self.manual_type_box = ctk.CTkTextbox(mf, height=100, font=("Segoe UI", 13))
        self.manual_type_box.pack(fill="both", expand=True, padx=10)
        mb = ctk.CTkFrame(mf, fg_color="transparent")
        mb.pack(fill="x", padx=10, pady=8)
        ctk.CTkButton(mb, text="⌨️ Type This Text", fg_color="#0078d4",
                      command=self._manual_type, width=160).pack(side="left")
        ctk.CTkButton(mb, text="📋 Paste & Type Clipboard", command=self._paste_and_type,
                      width=180).pack(side="left", padx=8)

    # ── Presets Tab ───────────────────────────────────────────────────────
    def _build_presets_tab(self):
        tab = self.tabs.tab("Presets")
        ctk.CTkLabel(tab, text="Click a preset to load it into the output. Edit / add below.",
                     text_color="#aaa").pack(pady=8)
        self.preset_frame = ctk.CTkScrollableFrame(tab, height=260)
        self.preset_frame.pack(fill="both", expand=True, padx=5)
        self._render_presets()
        ctk.CTkLabel(tab, text="New preset:").pack(anchor="w", padx=10, pady=(8,2))
        self.new_preset_box = ctk.CTkTextbox(tab, height=70)
        self.new_preset_box.pack(fill="x", padx=10)
        br = ctk.CTkFrame(tab, fg_color="transparent")
        br.pack(fill="x", padx=10, pady=6)
        ctk.CTkButton(br, text="➕ Add Preset", command=self._add_preset, width=130).pack(side="left")
        ctk.CTkButton(br, text="💾 Save", command=self._save_presets,
                      fg_color="#107c10", width=100).pack(side="left", padx=8)

    def _render_presets(self):
        for w in self.preset_frame.winfo_children():
            w.destroy()
        for i, msg in enumerate(self.cfg["preset_messages"]):
            row = ctk.CTkFrame(self.preset_frame, fg_color="#1e1e2e", corner_radius=6)
            row.pack(fill="x", pady=3, padx=2)
            ctk.CTkButton(row, text=msg[:90]+("…" if len(msg)>90 else ""), anchor="w",
                          fg_color="transparent", hover_color="#333",
                          command=lambda m=msg: self._load_preset(m)).pack(side="left", fill="x", expand=True, padx=6)
            ctk.CTkButton(row, text="⌨️", width=36, fg_color="#0078d4",
                          command=lambda m=msg: self._type_preset(m)).pack(side="right", padx=2, pady=4)
            ctk.CTkButton(row, text="🗑", width=36, fg_color="#c00",
                          command=lambda idx=i: self._delete_preset(idx)).pack(side="right", padx=2, pady=4)

    # ── Settings Tab ──────────────────────────────────────────────────────
    def _build_settings_tab(self):
        tab = self.tabs.tab("Settings")
        frm = ctk.CTkScrollableFrame(tab)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        def row(label, widget_fn):
            r = ctk.CTkFrame(frm, fg_color="transparent")
            r.pack(fill="x", pady=5)
            ctk.CTkLabel(r, text=label, width=220, anchor="w").pack(side="left")
            widget_fn(r)

        self.api_key_var = ctk.StringVar(value=self.cfg["api_key"])
        row("OpenAI API Key:", lambda p: ctk.CTkEntry(p, textvariable=self.api_key_var, show="*", width=320).pack(side="left"))

        self.model_var = ctk.StringVar(value=self.cfg["model"])
        row("Model:", lambda p: ctk.CTkOptionMenu(p, values=["gpt-4o-mini","gpt-4o","gpt-3.5-turbo"],
            variable=self.model_var, width=200).pack(side="left"))

        self.hotkey_var = ctk.StringVar(value=self.cfg["hotkey"])
        row("Generate+Type Hotkey:", lambda p: ctk.CTkEntry(p, textvariable=self.hotkey_var, width=200).pack(side="left"))

        self.cap_hotkey_var = ctk.StringVar(value=self.cfg["capture_hotkey"])
        row("Capture Hotkey:", lambda p: ctk.CTkEntry(p, textvariable=self.cap_hotkey_var, width=200).pack(side="left"))

        self.tesseract_var = ctk.StringVar(value=pytesseract.pytesseract.tesseract_cmd)
        row("Tesseract Path:", lambda p: ctk.CTkEntry(p, textvariable=self.tesseract_var, width=320).pack(side="left"))

        row("Context / Persona:", lambda p: None)
        self.context_box = ctk.CTkTextbox(frm, height=80)
        self.context_box.pack(fill="x", pady=4)
        self.context_box.insert("1.0", self.cfg["context_prompt"])

        ctk.CTkButton(frm, text="💾 Save Settings", command=self._save_settings,
                      fg_color="#107c10").pack(pady=14)

    # ── Log Tab ───────────────────────────────────────────────────────────
    def _build_log_tab(self):
        tab = self.tabs.tab("Log")
        self.log_box = ctk.CTkTextbox(tab, font=("Consolas", 12), state="disabled")
        self.log_box.pack(fill="both", expand=True, padx=5, pady=5)
        ctk.CTkButton(tab, text="🗑 Clear Log", command=self._clear_log,
                      fg_color="#444", width=120).pack(pady=6)

    # ── Capture + Text Helpers ───────────────────────────────────────────
    def _discover_monitor_targets(self):
        targets = {"all": "All Monitors", "primary": "Primary Monitor"}
        try:
            with mss.mss() as sct:
                for i in range(1, len(sct.monitors)):
                    mon = sct.monitors[i]
                    targets[f"monitor:{i}"] = f"Monitor {i} ({mon['width']}x{mon['height']})"
        except Exception:
            pass
        return targets

    def _get_virtual_screen_bounds(self):
        user32 = ctypes.windll.user32
        return {
            "left": int(user32.GetSystemMetrics(76)),
            "top": int(user32.GetSystemMetrics(77)),
            "width": int(user32.GetSystemMetrics(78)),
            "height": int(user32.GetSystemMetrics(79)),
        }

    def _selected_monitor_key(self):
        if hasattr(self, "capture_target_var"):
            selected_label = self.capture_target_var.get()
            return self.capture_target_names.get(selected_label, "all")
        return self.cfg.get("capture_monitor", "all")

    def _resolve_capture_monitor(self, sct):
        key = self._selected_monitor_key()
        monitors = sct.monitors
        if key == "primary":
            return monitors[1] if len(monitors) > 1 else monitors[0]
        if key.startswith("monitor:"):
            try:
                idx = int(key.split(":", 1)[1])
                if 0 < idx < len(monitors):
                    return monitors[idx]
            except (TypeError, ValueError):
                pass
        return monitors[0]

    def _monitor_target_name(self):
        return self.monitor_targets.get(self._selected_monitor_key(), "All Monitors")

    def _on_capture_target_change(self, selected_label):
        key = self.capture_target_names.get(selected_label, "all")
        self.cfg["capture_monitor"] = key
        self.region_lbl.configure(text=self._region_text())
        self._log(f"🖥 Capture target set to: {self._monitor_target_name()}")

    def _bind_live_counters(self):
        self.input_box.bind("<KeyRelease>", lambda _e: self._update_text_stats(self.input_box, self.input_stats_lbl))
        self.output_box.bind("<KeyRelease>", lambda _e: self._update_text_stats(self.output_box, self.output_stats_lbl))
        self.ocr_box.bind("<KeyRelease>", lambda _e: self._update_text_stats(self.ocr_box, self.ocr_stats_lbl))

    def _update_text_stats(self, textbox, label):
        text = textbox.get("1.0", "end").strip()
        chars = len(text)
        words = len(text.split()) if text else 0
        label.configure(text=f"{chars} chars - {words} words")

    def _update_all_stats(self):
        self._update_text_stats(self.input_box, self.input_stats_lbl)
        self._update_text_stats(self.output_box, self.output_stats_lbl)
        self._update_text_stats(self.ocr_box, self.ocr_stats_lbl)

    # ── Screen Capture Logic ──────────────────────────────────────────────
    def _region_text(self):
        r = self.captured_region
        prefix = f"Target: {self._monitor_target_name()}"
        if not r:
            return f"{prefix} | Full screen"
        return f"{prefix} | x={r[0]}, y={r[1]}, w={r[2]}, h={r[3]}"

    def _pick_region(self):
        self.iconify()
        time.sleep(0.4)
        bounds = self._get_virtual_screen_bounds()
        try:
            with mss.mss() as sct:
                bounds = self._resolve_capture_monitor(sct)
        except Exception:
            pass
        sel = RegionSelector(self, self._on_region_selected, bounds=bounds)
        sel.wait_window()
        self.deiconify()

    def _on_region_selected(self, region):
        self.captured_region = region
        self.cfg["capture_region"] = region
        self.region_lbl.configure(text=self._region_text())
        self._log(f"📐 Region set: {region}")

    def _use_full_screen(self):
        self.captured_region = None
        self.cfg["capture_region"] = None
        self.region_lbl.configure(text=self._region_text())
        self._log(f"🖥 Using full screen capture ({self._monitor_target_name()}).")

    def _clear_region(self):
        self._use_full_screen()

    def _capture_image_and_ocr(self):
        with mss.mss() as sct:
            if self.captured_region:
                x, y, w, h = self.captured_region
                monitor = {"left": x, "top": y, "width": w, "height": h}
            else:
                monitor = self._resolve_capture_monitor(sct)
            shot = sct.grab(monitor)
            img = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
        text = pytesseract.image_to_string(img).strip()
        return img, text

    def _normalize_detected_text(self, text: str):
        return " ".join(text.split())

    def _capture_now(self):
        self._log("📸 Capturing...")
        threading.Thread(target=self._do_capture, daemon=True).start()

    def _capture_and_generate(self):
        self._auto_generate_after_capture = True
        self._capture_now()

    def _do_capture(self):
        try:
            img, ocr_text = self._capture_image_and_ocr()
            self.reply_queue.put(("capture_done", (img, ocr_text)))
        except Exception as e:
            self.reply_queue.put(("error", f"Capture failed: {e}"))

    def _start_watch(self):
        if self.watch_running:
            self._log("ℹ️ Watch mode is already running.")
            return
        if not self._current_cfg().get("api_key"):
            self._log("❌ No API key. Go to Settings before starting watch mode.")
            return
        self.cfg["watch_interval_ms"] = int(self.watch_interval_var.get())
        self.cfg["detect_existing_text"] = bool(self.detect_existing_var.get())
        self.watch_detect_existing = bool(self.detect_existing_var.get())
        self.watch_interval_s = max(0.3, int(self.watch_interval_var.get()) / 1000.0)
        self.watch_running = True
        self.watch_stop_event.clear()
        self.watch_status_lbl.configure(text="Status: Starting...", text_color="#f0a500")

        if self.watch_detect_existing:
            try:
                _img, baseline = self._capture_image_and_ocr()
                self.last_detected_text = self._normalize_detected_text(baseline)
                if self.last_detected_text:
                    self._log("👀 Baseline captured. Waiting for new text...")
            except Exception as e:
                self._log(f"⚠ Baseline capture failed: {e}")
                self.last_detected_text = ""
        else:
            self.last_detected_text = ""

        self.watch_thread = threading.Thread(target=self._watch_loop, daemon=True)
        self.watch_thread.start()
        self.watch_status_lbl.configure(text="Status: Running", text_color="#00c853")
        self._log("▶ Watch mode started.")

    def _stop_watch(self):
        if not self.watch_running:
            self.watch_status_lbl.configure(text="Status: Stopped", text_color="#aaa")
            return
        self.watch_running = False
        self.watch_stop_event.set()
        self.watch_status_lbl.configure(text="Status: Stopped", text_color="#aaa")
        self._log("⏹ Watch mode stopped.")

    def _watch_loop(self):
        while not self.watch_stop_event.is_set():
            try:
                img, ocr_text = self._capture_image_and_ocr()
                normalized = self._normalize_detected_text(ocr_text)
                if not normalized:
                    time.sleep(self.watch_interval_s)
                    continue

                if self.watch_detect_existing and normalized == self.last_detected_text:
                    time.sleep(self.watch_interval_s)
                    continue

                self.last_detected_text = normalized
                self.reply_queue.put(("watch_detected", (img, ocr_text)))

                cfg = self._current_cfg()
                if not cfg["api_key"]:
                    self.reply_queue.put(("error", "No API key. Go to Settings."))
                    self.reply_queue.put(("watch_status", "stopped"))
                    return
                reply = generate_reply(ocr_text, cfg)
                self.reply_queue.put(("watch_reply", reply))
            except Exception as e:
                self.reply_queue.put(("error", f"Watch loop error: {e}"))
            time.sleep(self.watch_interval_s)

    def _show_capture_result(self, img: Image.Image, ocr_text: str):
        self.ocr_box.delete("1.0", "end")
        self.ocr_box.insert("1.0", ocr_text)
        self._update_text_stats(self.ocr_box, self.ocr_stats_lbl)
        thumb = img.copy()
        thumb.thumbnail((400, 200))
        self._thumb_photo = ImageTk.PhotoImage(thumb)
        self.thumb_lbl.configure(image=self._thumb_photo, text="")
        self._log(f"✅ Captured. OCR extracted {len(ocr_text)} chars.")
        if getattr(self, "_auto_generate_after_capture", False):
            self._auto_generate_after_capture = False
            self._ocr_to_compose()

    def _ocr_to_compose(self):
        text = self.ocr_box.get("1.0", "end").strip()
        if not text:
            self._log("⚠ No OCR text to send.")
            return
        self.input_box.delete("1.0", "end")
        self.input_box.insert("1.0", text)
        self._update_text_stats(self.input_box, self.input_stats_lbl)
        self.tabs.set("Compose")
        self._generate_clicked()

    # ── AutoType Logic ────────────────────────────────────────────────────
    def _launch_autotype_preview(self):
        reply = self.output_box.get("1.0", "end").strip()
        if not reply:
            self._log("⚠ No reply to type.")
            return
        delay_s = self.delay_var.get() / 1000
        if self.preview_before_var.get():
            AutoTypePreview(self, reply, delay_s,
                            on_confirm=self._do_type_text,
                            on_cancel=lambda: self._log("⌨️ AutoType cancelled."))
        else:
            self.iconify()
            threading.Thread(target=lambda: (time.sleep(delay_s), self._do_type_text(reply)), daemon=True).start()

    def _do_type_text(self, text: str):
        self._set_status("TYPING…", "#f0a500")
        self.iconify()
        cps = self.cps_var.get()
        auto_send = self.auto_send_var2.get()
        def worker():
            time.sleep(0.3)
            type_text(text, cps)
            if auto_send:
                time.sleep(0.1)
                keyboard_controller.press(Key.enter)
                keyboard_controller.release(Key.enter)
            self.reply_queue.put(("type_done", None))
        threading.Thread(target=worker, daemon=True).start()

    def _manual_type(self):
        text = self.manual_type_box.get("1.0", "end").strip()
        if text:
            delay_s = self.delay_var.get() / 1000
            if self.preview_before_var.get():
                AutoTypePreview(self, text, delay_s,
                                on_confirm=self._do_type_text,
                                on_cancel=lambda: None)
            else:
                threading.Thread(target=lambda: (time.sleep(delay_s), self._do_type_text(text)), daemon=True).start()

    def _paste_and_type(self):
        text = pyperclip.paste().strip()
        if text:
            self.manual_type_box.delete("1.0", "end")
            self.manual_type_box.insert("1.0", text)
            self._manual_type()

    def _type_preset(self, msg: str):
        self._do_type_text(msg)

    # ── Generate Logic ────────────────────────────────────────────────────
    def _generate_clicked(self):
        prompt = self.input_box.get("1.0", "end").strip()
        if not prompt:
            self._log("⚠ No input message.")
            return
        self._set_status("GENERATING…", "#f0a500")
        threading.Thread(target=self._generate_worker, args=(prompt,), daemon=True).start()

    def _generate_worker(self, prompt: str):
        cfg = self._current_cfg()
        if not cfg["api_key"]:
            self.reply_queue.put(("error", "No API key. Go to Settings."))
            return
        try:
            reply = generate_reply(prompt, cfg)
            self.reply_queue.put(("reply", reply))
        except Exception as e:
            self.reply_queue.put(("error", str(e)))

    def _refine_clicked(self, instruction: str):
        reply = self.output_box.get("1.0", "end").strip()
        if not reply:
            self._log("⚠ No reply to refine.")
            return
        self._set_status("POLISHING...", "#2d5bff")
        threading.Thread(target=self._refine_worker, args=(reply, instruction), daemon=True).start()

    def _refine_worker(self, reply: str, instruction: str):
        cfg = self._current_cfg()
        if not cfg["api_key"]:
            self.reply_queue.put(("error", "No API key. Go to Settings."))
            return
        try:
            revised = refine_reply(reply, instruction, cfg)
            self.reply_queue.put(("refined", revised))
        except Exception as e:
            self.reply_queue.put(("error", str(e)))

    # ── Queue Poller ──────────────────────────────────────────────────────
    def _poll_queue(self):
        while not self.reply_queue.empty():
            kind, val = self.reply_queue.get()
            if kind == "reply":
                self.output_box.delete("1.0", "end")
                self.output_box.insert("1.0", val)
                self._update_text_stats(self.output_box, self.output_stats_lbl)
                self._set_status("READY", "#00c853")
                self._log(f"✅ Reply generated ({len(val)} chars)")
            elif kind == "refined":
                self.output_box.delete("1.0", "end")
                self.output_box.insert("1.0", val)
                self._update_text_stats(self.output_box, self.output_stats_lbl)
                self._set_status("READY", "#00c853")
                self._log("✨ Reply polished.")
            elif kind == "watch_detected":
                img, ocr = val
                self._show_capture_result(img, ocr)
                self._set_status("NEW TEXT", "#2d5bff")
            elif kind == "watch_reply":
                self.output_box.delete("1.0", "end")
                self.output_box.insert("1.0", val)
                self._update_text_stats(self.output_box, self.output_stats_lbl)
                self._set_status("READY", "#00c853")
                self._log("🤖 Watch mode generated a reply.")
                if self.autotype_enabled_var.get():
                    self._do_type_text(val)
            elif kind == "watch_status":
                if val == "stopped":
                    self._stop_watch()
            elif kind == "error":
                self._set_status("ERROR", "#e53935")
                self._log(f"❌ {val}")
            elif kind == "type_done":
                self._set_status("TYPED ✓", "#00c853")
                self.deiconify()
                self._log("⌨️ Reply typed into active window.")
            elif kind == "capture_done":
                img, ocr = val
                self._show_capture_result(img, ocr)
        self.after(200, self._poll_queue)

    # ── Hotkeys ───────────────────────────────────────────────────────────
    def _register_hotkeys(self):
        for h in self._hotkey_hooks:
            try: keyboard.remove_hotkey(h)
            except: pass
        self._hotkey_hooks = []
        try:
            h1 = keyboard.add_hotkey(self.cfg["hotkey"], self._hotkey_generate_type)
            h2 = keyboard.add_hotkey(self.cfg["capture_hotkey"], self._hotkey_capture)
            self._hotkey_hooks = [h1, h2]
            self._log(f"⌨️ Hotkeys: {self.cfg['hotkey']} | {self.cfg['capture_hotkey']}")
        except Exception as e:
            self._log(f"❌ Hotkey error: {e}")

    def _hotkey_generate_type(self):
        prompt = pyperclip.paste().strip()
        if not prompt:
            self._log("⚠ Hotkey: clipboard empty.")
            return
        self._log("🔥 Hotkey: generate+type triggered")
        self.input_box.after(0, lambda: (self.input_box.delete("1.0","end"), self.input_box.insert("1.0", prompt)))
        cfg = self._current_cfg()
        if not cfg["api_key"]:
            self._log("❌ No API key.")
            return
        def worker():
            try:
                reply = generate_reply(prompt, cfg)
                self.reply_queue.put(("reply", reply))
                time.sleep(0.3)
                self._do_type_text(reply)
            except Exception as e:
                self.reply_queue.put(("error", str(e)))
        threading.Thread(target=worker, daemon=True).start()

    def _hotkey_capture(self):
        self._log("🔥 Capture hotkey triggered")
        threading.Thread(target=self._do_capture, daemon=True).start()

    # ── Helpers ───────────────────────────────────────────────────────────
    def _log(self, msg: str):
        ts = time.strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{ts}] {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _set_status(self, text: str, color: str = "#888"):
        self.status_dot.configure(text=f"● {text}", text_color=color)

    def _paste_from_clipboard(self):
        text = pyperclip.paste()
        self.input_box.delete("1.0", "end")
        self.input_box.insert("1.0", text)
        self._update_text_stats(self.input_box, self.input_stats_lbl)
        self._log(f"📋 Pasted {len(text)} chars.")

    def _copy_reply(self):
        pyperclip.copy(self.output_box.get("1.0","end").strip())
        self._log("📋 Reply copied.")

    def _clear_all(self):
        self.input_box.delete("1.0","end")
        self.output_box.delete("1.0","end")
        self._update_text_stats(self.input_box, self.input_stats_lbl)
        self._update_text_stats(self.output_box, self.output_stats_lbl)

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0","end")
        self.log_box.configure(state="disabled")

    def _load_preset(self, msg: str):
        self.output_box.delete("1.0","end")
        self.output_box.insert("1.0", msg)
        self._update_text_stats(self.output_box, self.output_stats_lbl)
        self.tabs.set("Compose")

    def _add_preset(self):
        msg = self.new_preset_box.get("1.0","end").strip()
        if msg:
            self.cfg["preset_messages"].append(msg)
            self.new_preset_box.delete("1.0","end")
            self._render_presets()

    def _delete_preset(self, idx):
        self.cfg["preset_messages"].pop(idx)
        self._render_presets()

    def _save_presets(self):
        save_config(self.cfg)
        self._log("💾 Presets saved.")

    def _save_settings(self):
        self.cfg["api_key"] = self.api_key_var.get().strip()
        self.cfg["model"] = self.model_var.get()
        self.cfg["typing_speed_cps"] = self.cps_var.get()
        self.cfg["delay_ms"] = self.delay_var.get()
        self.cfg["context_prompt"] = self.context_box.get("1.0","end").strip()
        self.cfg["autotype_enabled"] = self.autotype_enabled_var.get()
        self.cfg["auto_send"] = self.auto_send_var2.get()
        self.cfg["capture_monitor"] = self._selected_monitor_key()
        self.cfg["watch_interval_ms"] = int(self.watch_interval_var.get())
        self.cfg["detect_existing_text"] = bool(self.detect_existing_var.get())
        pytesseract.pytesseract.tesseract_cmd = self.tesseract_var.get()
        changed = (self.hotkey_var.get() != self.cfg["hotkey"] or
                   self.cap_hotkey_var.get() != self.cfg["capture_hotkey"])
        self.cfg["hotkey"] = self.hotkey_var.get().strip()
        self.cfg["capture_hotkey"] = self.cap_hotkey_var.get().strip()
        if changed:
            self._register_hotkeys()
        save_config(self.cfg)
        self._log("💾 Settings saved.")

    def _current_cfg(self):
        return {
            **self.cfg,
            "tone": self.tone_var.get(),
            "platform": self.platform_var.get(),
            "typing_speed_cps": self.cps_var.get(),
            "auto_send": self.auto_send_var2.get(),
        }

    def on_close(self):
        self.cfg["typing_speed_cps"] = self.cps_var.get()
        self.cfg["delay_ms"] = self.delay_var.get()
        self.cfg["capture_monitor"] = self._selected_monitor_key()
        self.cfg["watch_interval_ms"] = int(self.watch_interval_var.get())
        self.cfg["detect_existing_text"] = bool(self.detect_existing_var.get())
        self._stop_watch()
        save_config(self.cfg)
        for h in self._hotkey_hooks:
            try: keyboard.remove_hotkey(h)
            except: pass
        self.destroy()

if __name__ == "__main__":
    app = AutoResponderApp()
    app.mainloop()














