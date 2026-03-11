"""
AutoResponder Bot - Windows AI-powered typing assistant
Requirements:
    pip install openai pyperclip pynput keyboard tkinter customtkinter pygetwindow

Run: python autoresponder.py
"""

import tkinter as tk
import customtkinter as ctk
import threading
import time
import json
import os
import pyperclip
import keyboard
import pygetwindow as gw
from openai import OpenAI
from pynput import mouse, keyboard as kb
from pynput.keyboard import Controller, Key
import queue

# ── Config ──────────────────────────────────────────────────────────────────
CONFIG_FILE = "autoresponder_config.json"
DEFAULT_CONFIG = {
    "api_key": "",
    "model": "gpt-4o-mini",
    "hotkey": "ctrl+shift+space",
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
    "typing_speed_cps": 30,  # chars per second
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
    "Discord":      "Use Discord markdown if helpful (bold, italics). Keep responses chat-length.",
    "Slack":        "Use Slack formatting if helpful. Keep professional but conversational.",
    "WhatsApp Web": "Keep it short and conversational, no markdown.",
    "Teams":        "Professional tone appropriate for Microsoft Teams.",
    "Gmail":        "Format as a proper email reply with greeting and sign-off.",
    "Telegram Web": "Short, conversational reply.",
}

# ── OpenAI Client ─────────────────────────────────────────────────────────
keyboard_controller = Controller()

def generate_reply(prompt: str, cfg: dict) -> str:
    client = OpenAI(api_key=cfg["api_key"])
    tone = cfg["tone"]
    platform = cfg["platform"]
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

def type_text(text: str, cps: int = 30):
    """Type text character by character at a human-like speed."""
    delay = 1.0 / max(cps, 1)
    for ch in text:
        keyboard_controller.type(ch)
        time.sleep(delay)

# ── Config I/O ───────────────────────────────────────────────────────────
def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE) as f:
            saved = json.load(f)
        cfg = {**DEFAULT_CONFIG, **saved}
        return cfg
    return DEFAULT_CONFIG.copy()

def save_config(cfg: dict):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)

# ── Main App ─────────────────────────────────────────────────────────────
class AutoResponderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.cfg = load_config()
        self.running = False
        self.hotkey_hook = None
        self.reply_queue = queue.Queue()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title("🤖 AutoResponder Bot")
        self.geometry("780x720")
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self._build_ui()
        self._register_hotkey()
        self.after(200, self._poll_queue)

    # ── UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = ctk.CTkFrame(self, corner_radius=0, height=60, fg_color="#1a1a2e")
        hdr.pack(fill="x")
        ctk.CTkLabel(hdr, text="🤖  AutoResponder Bot", font=("Segoe UI", 20, "bold"),
                     text_color="#00d4ff").pack(side="left", padx=20, pady=10)
        self.status_dot = ctk.CTkLabel(hdr, text="● IDLE", font=("Segoe UI", 12),
                                        text_color="#888")
        self.status_dot.pack(side="right", padx=20)

        # Tabs
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=10)
        for t in ["Compose", "Presets", "Settings", "Log"]:
            self.tabs.add(t)

        self._build_compose_tab()
        self._build_presets_tab()
        self._build_settings_tab()
        self._build_log_tab()

    # ── Compose Tab ─────────────────────────────────────────────────────
    def _build_compose_tab(self):
        tab = self.tabs.tab("Compose")

        ctk.CTkLabel(tab, text="Paste / type the incoming message:", anchor="w").pack(fill="x", padx=5, pady=(10,2))
        self.input_box = ctk.CTkTextbox(tab, height=120, font=("Segoe UI", 13))
        self.input_box.pack(fill="x", padx=5)

        row = ctk.CTkFrame(tab, fg_color="transparent")
        row.pack(fill="x", padx=5, pady=8)
        ctk.CTkLabel(row, text="Tone:").pack(side="left")
        self.tone_var = ctk.StringVar(value=self.cfg["tone"])
        tone_menu = ctk.CTkOptionMenu(row, values=TONES, variable=self.tone_var,
                                       command=self._on_tone_change, width=150)
        tone_menu.pack(side="left", padx=8)

        ctk.CTkLabel(row, text="Platform:").pack(side="left", padx=(20,0))
        self.platform_var = ctk.StringVar(value=self.cfg["platform"])
        ctk.CTkOptionMenu(row, values=PLATFORMS, variable=self.platform_var,
                           command=self._on_platform_change, width=150).pack(side="left", padx=8)

        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.pack(fill="x", padx=5, pady=4)
        ctk.CTkButton(btn_row, text="⚡ Generate Reply", command=self._generate_clicked,
                      fg_color="#0078d4", hover_color="#005a9e", width=160).pack(side="left")
        ctk.CTkButton(btn_row, text="📋 Copy Input from Clipboard",
                      command=self._paste_from_clipboard, width=200).pack(side="left", padx=10)
        ctk.CTkButton(btn_row, text="🗑 Clear", command=self._clear_all,
                      fg_color="#444", hover_color="#666", width=80).pack(side="left")

        ctk.CTkLabel(tab, text="Generated reply:", anchor="w").pack(fill="x", padx=5, pady=(12,2))
        self.output_box = ctk.CTkTextbox(tab, height=150, font=("Segoe UI", 13))
        self.output_box.pack(fill="x", padx=5)

        action_row = ctk.CTkFrame(tab, fg_color="transparent")
        action_row.pack(fill="x", padx=5, pady=8)
        ctk.CTkButton(action_row, text="⌨️ Type Reply into Active Window",
                      command=self._type_reply, fg_color="#107c10", hover_color="#0a5c0a",
                      width=220).pack(side="left")
        ctk.CTkButton(action_row, text="📋 Copy Reply",
                      command=self._copy_reply, width=120).pack(side="left", padx=10)
        self.auto_send_var = ctk.BooleanVar(value=self.cfg["auto_send"])
        ctk.CTkCheckBox(action_row, text="Auto-send (press Enter after typing)",
                        variable=self.auto_send_var).pack(side="left", padx=10)

        ctk.CTkLabel(tab, text=f"💡 Hotkey: {self.cfg['hotkey']}  — triggers Generate + Type from clipboard",
                     font=("Segoe UI", 11), text_color="#888").pack(pady=4)

    # ── Presets Tab ──────────────────────────────────────────────────────
    def _build_presets_tab(self):
        tab = self.tabs.tab("Presets")
        ctk.CTkLabel(tab, text="Click a preset to load it into the output box. Edit below.",
                     text_color="#aaa").pack(pady=8)

        self.preset_frame = ctk.CTkScrollableFrame(tab, height=280)
        self.preset_frame.pack(fill="both", expand=True, padx=5)
        self._render_presets()

        ctk.CTkLabel(tab, text="Add / edit preset:").pack(anchor="w", padx=10, pady=(10,2))
        self.new_preset_box = ctk.CTkTextbox(tab, height=70)
        self.new_preset_box.pack(fill="x", padx=10)
        btn_r = ctk.CTkFrame(tab, fg_color="transparent")
        btn_r.pack(fill="x", padx=10, pady=6)
        ctk.CTkButton(btn_r, text="➕ Add Preset", command=self._add_preset, width=130).pack(side="left")
        ctk.CTkButton(btn_r, text="💾 Save Presets", command=self._save_presets,
                      fg_color="#107c10", width=130).pack(side="left", padx=10)

    def _render_presets(self):
        for w in self.preset_frame.winfo_children():
            w.destroy()
        for i, msg in enumerate(self.cfg["preset_messages"]):
            row = ctk.CTkFrame(self.preset_frame, fg_color="#1e1e2e", corner_radius=6)
            row.pack(fill="x", pady=3, padx=2)
            ctk.CTkButton(row, text=msg[:80] + ("…" if len(msg) > 80 else ""),
                          anchor="w", fg_color="transparent", hover_color="#333",
                          command=lambda m=msg: self._load_preset(m)).pack(side="left", fill="x", expand=True, padx=6)
            ctk.CTkButton(row, text="🗑", width=36, fg_color="#c00", hover_color="#900",
                          command=lambda idx=i: self._delete_preset(idx)).pack(side="right", padx=4, pady=4)

    # ── Settings Tab ─────────────────────────────────────────────────────
    def _build_settings_tab(self):
        tab = self.tabs.tab("Settings")
        frm = ctk.CTkScrollableFrame(tab)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        def row(label, widget_fn):
            r = ctk.CTkFrame(frm, fg_color="transparent")
            r.pack(fill="x", pady=5)
            ctk.CTkLabel(r, text=label, width=200, anchor="w").pack(side="left")
            widget_fn(r)

        self.api_key_var = ctk.StringVar(value=self.cfg["api_key"])
        row("OpenAI API Key:", lambda p: ctk.CTkEntry(p, textvariable=self.api_key_var,
            show="*", width=320).pack(side="left"))

        self.model_var = ctk.StringVar(value=self.cfg["model"])
        row("Model:", lambda p: ctk.CTkOptionMenu(p, values=["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"],
            variable=self.model_var, width=200).pack(side="left"))

        self.hotkey_var = ctk.StringVar(value=self.cfg["hotkey"])
        row("Hotkey:", lambda p: ctk.CTkEntry(p, textvariable=self.hotkey_var, width=200).pack(side="left"))

        self.delay_var = ctk.IntVar(value=self.cfg["delay_ms"])
        row("Hotkey Delay (ms):", lambda p: ctk.CTkSlider(p, from_=0, to=2000,
            variable=self.delay_var, width=200).pack(side="left"))

        self.cps_var = ctk.IntVar(value=self.cfg["typing_speed_cps"])
        row("Typing Speed (chars/sec):", lambda p: ctk.CTkSlider(p, from_=5, to=120,
            variable=self.cps_var, width=200).pack(side="left"))

        row("Context / Persona:", lambda p: None)
        self.context_box = ctk.CTkTextbox(frm, height=80)
        self.context_box.pack(fill="x", pady=4)
        self.context_box.insert("1.0", self.cfg["context_prompt"])

        ctk.CTkButton(frm, text="💾 Save Settings", command=self._save_settings,
                      fg_color="#107c10").pack(pady=14)

    # ── Log Tab ──────────────────────────────────────────────────────────
    def _build_log_tab(self):
        tab = self.tabs.tab("Log")
        self.log_box = ctk.CTkTextbox(tab, font=("Consolas", 12), state="disabled")
        self.log_box.pack(fill="both", expand=True, padx=5, pady=5)
        ctk.CTkButton(tab, text="🗑 Clear Log", command=self._clear_log,
                      fg_color="#444", width=120).pack(pady=6)

    # ── Logic ────────────────────────────────────────────────────────────
    def _log(self, msg: str):
        ts = time.strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{ts}] {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _set_status(self, text: str, color: str = "#888"):
        self.status_dot.configure(text=f"● {text}", text_color=color)

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
            self.reply_queue.put(("error", "No API key set. Go to Settings."))
            return
        try:
            reply = generate_reply(prompt, cfg)
            self.reply_queue.put(("reply", reply))
        except Exception as e:
            self.reply_queue.put(("error", str(e)))

    def _poll_queue(self):
        while not self.reply_queue.empty():
            kind, val = self.reply_queue.get()
            if kind == "reply":
                self.output_box.delete("1.0", "end")
                self.output_box.insert("1.0", val)
                self._set_status("READY", "#00c853")
                self._log(f"✅ Reply generated ({len(val)} chars)")
            elif kind == "error":
                self._set_status("ERROR", "#e53935")
                self._log(f"❌ {val}")
            elif kind == "type_done":
                self._set_status("TYPED ✓", "#00c853")
                self._log("⌨️ Reply typed into active window.")
        self.after(200, self._poll_queue)

    def _type_reply(self):
        reply = self.output_box.get("1.0", "end").strip()
        if not reply:
            self._log("⚠ No reply to type.")
            return
        self._set_status("TYPING…", "#f0a500")
        self.minimize_to_tray()
        delay_s = self.cfg.get("delay_ms", 500) / 1000
        cps = self.cps_var.get()
        auto_send = self.auto_send_var.get()

        def do_type():
            time.sleep(delay_s)
            type_text(reply, cps)
            if auto_send:
                time.sleep(0.1)
                keyboard_controller.press(Key.enter)
                keyboard_controller.release(Key.enter)
            self.reply_queue.put(("type_done", None))

        threading.Thread(target=do_type, daemon=True).start()

    def minimize_to_tray(self):
        self.iconify()
        self.after(2000, self.deiconify)

    def _paste_from_clipboard(self):
        text = pyperclip.paste()
        self.input_box.delete("1.0", "end")
        self.input_box.insert("1.0", text)
        self._log(f"📋 Pasted {len(text)} chars from clipboard.")

    def _copy_reply(self):
        reply = self.output_box.get("1.0", "end").strip()
        pyperclip.copy(reply)
        self._log("📋 Reply copied to clipboard.")

    def _clear_all(self):
        self.input_box.delete("1.0", "end")
        self.output_box.delete("1.0", "end")

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def _on_tone_change(self, val):
        self.cfg["tone"] = val

    def _on_platform_change(self, val):
        self.cfg["platform"] = val

    # ── Presets ──────────────────────────────────────────────────────────
    def _load_preset(self, msg: str):
        self.output_box.delete("1.0", "end")
        self.output_box.insert("1.0", msg)
        self.tabs.set("Compose")

    def _add_preset(self):
        msg = self.new_preset_box.get("1.0", "end").strip()
        if msg:
            self.cfg["preset_messages"].append(msg)
            self.new_preset_box.delete("1.0", "end")
            self._render_presets()

    def _delete_preset(self, idx: int):
        self.cfg["preset_messages"].pop(idx)
        self._render_presets()

    def _save_presets(self):
        save_config(self.cfg)
        self._log("💾 Presets saved.")

    # ── Settings ─────────────────────────────────────────────────────────
    def _save_settings(self):
        self.cfg["api_key"] = self.api_key_var.get().strip()
        self.cfg["model"] = self.model_var.get()
        self.cfg["delay_ms"] = self.delay_var.get()
        self.cfg["typing_speed_cps"] = self.cps_var.get()
        self.cfg["context_prompt"] = self.context_box.get("1.0", "end").strip()
        new_hk = self.hotkey_var.get().strip()
        if new_hk != self.cfg["hotkey"]:
            self.cfg["hotkey"] = new_hk
            self._register_hotkey()
        save_config(self.cfg)
        self._log("💾 Settings saved.")

    def _current_cfg(self) -> dict:
        return {
            **self.cfg,
            "tone": self.tone_var.get(),
            "platform": self.platform_var.get(),
            "auto_send": self.auto_send_var.get(),
        }

    # ── Hotkey ────────────────────────────────────────────────────────────
    def _register_hotkey(self):
        if self.hotkey_hook:
            try:
                keyboard.remove_hotkey(self.hotkey_hook)
            except Exception:
                pass
        try:
            self.hotkey_hook = keyboard.add_hotkey(self.cfg["hotkey"], self._hotkey_triggered)
            self._log(f"⌨️ Hotkey registered: {self.cfg['hotkey']}")
        except Exception as e:
            self._log(f"❌ Hotkey error: {e}")

    def _hotkey_triggered(self):
        """Grab clipboard → generate → type."""
        self._log("🔥 Hotkey triggered!")
        self._set_status("HOTKEY…", "#f0a500")
        prompt = pyperclip.paste().strip()
        if not prompt:
            self._log("⚠ Clipboard empty — nothing to reply to.")
            return
        self.input_box.delete("1.0", "end")
        self.input_box.insert("1.0", prompt)
        cfg = self._current_cfg()
        if not cfg["api_key"]:
            self._log("❌ No API key. Open Settings.")
            return
        def worker():
            try:
                reply = generate_reply(prompt, cfg)
                self.reply_queue.put(("reply", reply))
                self.output_box.after(0, lambda: self.output_box.delete("1.0", "end"))
                self.output_box.after(0, lambda: self.output_box.insert("1.0", reply))
                time.sleep(self.cfg["delay_ms"] / 1000)
                type_text(reply, cfg["typing_speed_cps"])
                if cfg["auto_send"]:
                    time.sleep(0.1)
                    keyboard_controller.press(Key.enter)
                    keyboard_controller.release(Key.enter)
                self.reply_queue.put(("type_done", None))
            except Exception as e:
                self.reply_queue.put(("error", str(e)))
        threading.Thread(target=worker, daemon=True).start()

    # ── Lifecycle ────────────────────────────────────────────────────────
    def on_close(self):
        save_config(self.cfg)
        if self.hotkey_hook:
            try:
                keyboard.remove_hotkey(self.hotkey_hook)
            except Exception:
                pass
        self.destroy()


# ── Entry Point ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = AutoResponderApp()
    app.mainloop()