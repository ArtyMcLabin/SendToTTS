# SendToTTS

A Windows application that reads clipboard content using Text-to-Speech (TTS) with global hotkeys. Runs completely windowless in the system tray by default.

## Features

- **System Tray Integration**: Runs completely windowless in the background with tray icon
- **Global Hotkeys**: 
  - `Alt+Q` - Read clipboard content / interrupt current speech and read new content
  - `Alt+Shift+Q` - Stop speech without reading new content
- **Debug Mode**: Console window with Enter key fallback for troubleshooting
- **Automatic Language Detection**: Supports Russian (Cyrillic) and English text
- **Voice Auto-Selection**: Automatically switches between Russian and English voices
- **Configurable Settings**: Adjust speech rate and volume via settings.ini
- **24/7 Operation**: Runs continuously with automatic error recovery
- **Robust Error Handling**: Self-healing hotkey system with automatic re-registration
- **No Multiple Hotkey Firing**: Fixed hotkey registration to prevent multiple executions

## Installation

1. Install Python 3.7 or higher
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Normal Operation (System Tray)
Double-click `run.bat` or run:
```bash
pythonw main.py
```
The application will run completely windowless in the system tray. Right-click the tray icon to access menu options.

### Debug Mode (Console Window)
Double-click `debug.bat` or run:
```bash
python main.py --debug
```
This opens a console window with detailed logging and Enter key fallback.

### Controls
- **Alt+Q**: Read clipboard content (global hotkey)
- **Alt+Shift+Q**: Stop current speech (global hotkey)  
- **Enter**: Read clipboard (debug mode only - works when console window is focused)
- **System Tray**: Right-click for menu options including quit

## Configuration

Edit `settings.ini` to customize:
- `speech_rate`: Speech speed (0-10, default: 3)
- `volume`: Speech volume (0.0-1.0, default: 0.8)

## Language Support

The application automatically detects and switches between:
- **Russian text** → Uses Microsoft Irina voice (if available)
- **English text** → Uses Microsoft Zira voice (if available)

## Troubleshooting

- Use debug mode (`debug.bat`) for detailed console output and Enter key fallback
- Check `tts_debug.log` for detailed error information
- The application automatically attempts to recover from hotkey failures
- Ensure Windows TTS voices are installed (Irina for Russian, Zira for English)
- Right-click system tray icon to quit if needed

## Requirements

- Windows 10/11
- Python 3.7+
- Windows Speech API (SAPI) voices installed 