# SendToTTS

A 24/7 Python application to send clipboard content to Windows' built-in Text-to-Speech (TTS) engine using hotkeys. Features configurable speech settings and voice selection.

## Features

- Reads text from the clipboard using Windows' native TTS engine
- Configurable speech rate, volume, and voice selection via `settings.ini`
- Toggle speech playback with `Alt+Q` (start/stop)
- Force stop speech with `Alt+Shift+Q`
- Displays all available TTS voices on startup
- Runs perpetually in the background (24/7)
- Real-time settings loading from configuration file

## Prerequisites

- Python 3.x
- Windows OS

## Setup and Usage

1.  **Clone the repository or download the files.**

2.  **Install the required Python packages:**
    Open a terminal or command prompt in the project directory and run:
    ```bash
    pip install -r requirements.txt
    ```

3.  **Configure settings (optional):**
    Edit `settings.ini` to customize:
    - `speech_rate`: Speech speed in words per minute (100-300, default: 200)
    - `voice_id`: Specific voice ID (see available voices when app starts)
    - `volume`: Volume level (0.0-1.0, default: 0.9)

4.  **Run the application:**
    ```bash
    python main.py
    ```

5.  **How to use:**
    - Copy any text to your clipboard
    - Press `Alt+Q` to start reading or stop current speech
    - Press `Alt+Shift+Q` to force stop current speech
    - Press `Ctrl+C` in the terminal to exit the application

## Configuration

### settings.ini

```ini
[TTS_SETTINGS]
# Speech rate (words per minute) - typical range: 100-300
speech_rate = 200

# Voice selection - use voice ID from the list shown on startup
voice_id = 

# Volume level (0.0 to 1.0)
volume = 0.9
```

To use a specific voice:
1. Run the application to see available voices
2. Copy the desired voice ID from the output
3. Paste it into the `voice_id` field in `settings.ini`
4. Restart the application

## Controls

- `Alt+Q`: Toggle speech (start reading clipboard or stop current speech)
- `Alt+Shift+Q`: Force stop current speech
- `Ctrl+C`: Exit the application

## How it works

The application uses the following Python libraries:
- `pyttsx3`: For Text-to-Speech functionality
- `pyperclip`: To access the clipboard
- `keyboard`: To listen for global hotkeys
- `configparser`: To load settings from INI file 