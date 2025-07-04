# v1.0.0
"""
Clipboard ➜ Windows TTS reader

Core design (simplified):
1. Single-threaded. We keep one SAPI.SpVoice instance.
2. Use SVSFlagsAsync | SVSFPurgeBeforeSpeak (value 3) so every new Speak() call
   automatically stops any current narration and starts the new one.
3. A separate Stop hot-key just purges the queue (Skip) without new speech.
4. The main loop pumps COM messages so async playback proceeds while the
   program waits for hot-keys.
"""

import time
import configparser
import os
import pyperclip
import keyboard
import pythoncom
import win32com.client
import queue

# ---------------------------------------------------------------------------
# Configuration ----------------------------------------------------------------
SETTINGS_FILE = "settings.ini"
DEFAULT_SETTINGS = {
    "speech_rate": 200,   # WPM ~200 ≈ SAPI rate 0
    "volume": 0.9,        # 0-1 range → SAPI 0-100
    "voice_id": ""        # Empty = default voice
}

settings: dict[str, object] = DEFAULT_SETTINGS.copy()

# ---------------------------------------------------------------------------
# Helpers ----------------------------------------------------------------------

def load_settings() -> None:
    """Load/merge settings from settings.ini if present."""
    cfg = configparser.ConfigParser()
    if os.path.exists(SETTINGS_FILE):
        cfg.read(SETTINGS_FILE)
        if "TTS_SETTINGS" in cfg:
            sect = cfg["TTS_SETTINGS"]
            settings["speech_rate"] = sect.getint("speech_rate", fallback=settings["speech_rate"])
            settings["volume"] = sect.getfloat("volume", fallback=settings["volume"])
            settings["voice_id"] = sect.get("voice_id", fallback=settings["voice_id"])


def wpm_to_sapi_rate(wpm: int) -> int:
    """Convert ~100-300 WPM to SAPI rate (-10 … 10). Rough mapping."""
    return max(-10, min(10, int((wpm - 200) / 10)))

# ---------------------------------------------------------------------------
# TTS Engine --------------------------------------------------------------------
pythoncom.CoInitialize()
_voice = win32com.client.Dispatch("SAPI.SpVoice")

# event queue for cross-thread communication
_event_q: queue.Queue = queue.Queue()

def configure_voice() -> None:
    """Apply current settings to the global _voice instance."""
    _voice.Rate = wpm_to_sapi_rate(int(settings["speech_rate"]))
    _voice.Volume = int(float(settings["volume"]) * 100)

    if settings["voice_id"]:
        for v in _voice.GetVoices():
            if v.Id == settings["voice_id"]:
                _voice.Voice = v
                break


def list_voices() -> None:
    print("\n=== Available Voices ===")
    for i, v in enumerate(_voice.GetVoices()):
        print(f"{i+1}. {v.GetDescription()}")
        print(f"   ID: {v.Id}")
        print("-" * 40)
    print("=" * 40)

# ---------------------------------------------------------------------------
# Hot-key handlers -------------------------------------------------------------

SVSFlagsAsync = 1  # speak async
SVSFPurgeBeforeSpeak = 2  # purge before speak
FLAGS_PURGE_ASYNC = SVSFlagsAsync | SVSFPurgeBeforeSpeak  # 3


def read_clipboard() -> None:
    text: str = pyperclip.paste().strip()
    if not text:
        print("Clipboard empty – nothing to read.")
        return
    _event_q.put(("read", text))


def stop_speech() -> None:
    _event_q.put(("stop", None))

# ---------------------------------------------------------------------------
# Main -------------------------------------------------------------------------

def main() -> None:
    print("=== Clipboard → TTS ===")
    load_settings()
    configure_voice()
    list_voices()

    hotkey = "alt+q"
    stop_hotkey = "alt+shift+q"
    print("\nControls:")
    print(f" {hotkey}  – read clipboard / interrupt & read new")
    print(f" {stop_hotkey} – stop speech without reading")
    print(" Ctrl+C – exit\n")

    keyboard.add_hotkey(hotkey, read_clipboard)
    keyboard.add_hotkey(stop_hotkey, stop_speech)

    print("Listening for hot-keys… (press Ctrl+C to quit)")
    try:
        while True:
            # process COM messages so async speech plays
            pythoncom.PumpWaitingMessages()
            # handle queued events coming from hot-keys
            try:
                evt, payload = _event_q.get_nowait()
                if evt == "read":
                    sample = (payload[:80] + "…") if len(payload) > 80 else payload
                    print(f"Reading: {sample}")
                    _voice.Speak(payload, FLAGS_PURGE_ASYNC)
                elif evt == "stop":
                    _voice.Skip("Sentence", 1000)
                    print("Speech stopped.")
            except queue.Empty:
                pass
            time.sleep(0.05)
    except KeyboardInterrupt:
        print("\nExiting…")
    finally:
        stop_speech()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main() 