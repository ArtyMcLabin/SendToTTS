# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

SendToTTS is a Windows application that converts clipboard text to speech using Windows SAPI (Speech API). It runs completely windowless in the system tray by default with global hotkeys for clipboard reading and speech control.

## Architecture

- **Single-file application**: All functionality is contained in `main.py` (~580 lines)
- **Threading model**: Uses separate threads for TTS operations, system tray, and main event loop
- **COM integration**: Heavily uses Windows COM objects for SAPI TTS and clipboard access
- **Global hotkey system**: Uses `keyboard` library with self-healing hotkey registration
- **Configuration**: Simple INI file for speech rate and volume settings

### Key Components

- **Voice Management** (`main.py:84-205`): Handles TTS voice initialization, language detection (Russian/English), and voice switching
- **Clipboard Operations** (`main.py:122-164`): Unicode-aware clipboard reading with retry logic
- **Hotkey System** (`main.py:270-335`): Robust global hotkey registration with automatic recovery
- **System Tray** (`main.py:337-422`): Windowless operation with tray icon and menu
- **Main Event Loop** (`main.py:441-578`): COM message pumping, heartbeat monitoring, and hotkey health checks

## Common Commands

### Running the Application

**Normal operation (system tray, windowless):**
```bash
pythonw main.py
# or
run.bat
```

**Debug mode (console window with detailed logging):**
```bash
python main.py --debug
# or
debug.bat
```

### Dependencies
```bash
pip install -r requirements.txt
```

### Configuration
Edit `settings.ini` to modify:
- `speech_rate`: TTS speed (-10 to 10, default: 2)
- `volume`: TTS volume (0 to 100, default: 100)

## Language Detection Logic

The application automatically switches TTS voices based on text content:
- **Russian text** (Cyrillic detection): Uses Microsoft Irina voice
- **English text** (default): Uses Microsoft Zira voice  
- **Hebrew text**: Detected but no voice available

Voice switching occurs in `detect_language()` and `set_voice_by_language()` functions.

## Error Recovery Systems

- **Hotkey auto-recovery**: Tests hotkeys every 60 seconds and re-registers if needed
- **COM error handling**: Automatic TTS voice reinitialization on COM failures
- **Clipboard retry logic**: 3 attempts with delays for clipboard access
- **Thread safety**: Proper COM initialization per thread

## Development Notes

- Windows-only application (requires pywin32, Windows SAPI)
- No unit tests present - testing requires Windows environment with TTS voices installed
- Logging configured differently for tray vs debug mode (null handler vs console+file)
- Global variables used extensively for voice, hotkey handlers, and application state