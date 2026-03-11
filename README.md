# Autoresponder

AI-powered Windows desktop autoresponder that can generate replies from copied text and type them into the active window, with a v2 option that adds OCR-based screen capture.

## About
This project provides a local Python GUI assistant for drafting and optionally auto-typing responses using OpenAI models. It is designed for fast message handling across common chat and communication platforms.

## Features
- Configurable response tone and platform style
- Global hotkey to generate replies
- Optional auto-send behavior with typing delay controls
- Preset message support
- `autoresponder-v2.py` adds OCR and capture-region workflows

## Project Files
- `autoresponder.py`: core autoresponder app
- `autoresponder-v2.py`: enhanced version with OCR/screen capture

## Requirements
Install Python 3.10+ and then install dependencies.

Core app:

```bash
pip install openai pyperclip pynput keyboard tkinter customtkinter pygetwindow
```

V2 extras:

```bash
pip install pillow pytesseract mss pyautogui
```

For OCR, install Tesseract and ensure its path matches the script setting:
- Suggested installer: https://github.com/UB-Mannheim/tesseract/wiki
- Default path used in code: `C:\Program Files\Tesseract-OCR\tesseract.exe`

## Quick Start
1. Set your OpenAI API key in the app settings.
2. Run one of the scripts:
   - `python autoresponder.py`
   - `python autoresponder-v2.py`
3. Use your configured hotkeys to generate and type replies.

## Notes
- This is a desktop automation tool intended for responsible use.
- Review generated messages before sending when accuracy is important.

## License
No license file is currently included.
