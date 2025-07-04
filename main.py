# SendToTTS Application v1.0.5
# Reads clipboard content and converts to speech using Windows SAPI
# Global hotkeys: Alt+Q (read/interrupt), Alt+Shift+Q (stop only)
# Local fallback: Enter key (works only when this window is focused)

import win32com.client
import win32clipboard
import win32con
import pythoncom
import keyboard
import time
import threading
import queue
import logging
import configparser
import sys
import os
import msvcrt

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('tts_debug.log'),
        logging.StreamHandler()
    ]
)

# Global variables
voice = None
event_queue = queue.Queue()
last_hotkey_time = time.time()
running = True
hotkey_handlers = []

def load_settings():
    """Load settings from settings.ini"""
    config = configparser.ConfigParser()
    
    # Default settings
    defaults = {
        'speech_rate': '0',
        'volume': '100',
        'voice_id': ''
    }
    
    if os.path.exists('settings.ini'):
        config.read('settings.ini')
        settings = {}
        for key in defaults:
            settings[key] = config.get('DEFAULT', key, fallback=defaults[key])
    else:
        settings = defaults
        # Create default settings file
        config['DEFAULT'] = defaults
        with open('settings.ini', 'w') as f:
            config.write(f)
    
    return settings

def list_available_voices():
    """List all available TTS voices"""
    try:
        temp_voice = win32com.client.Dispatch("SAPI.SpVoice")
        voices = temp_voice.GetVoices()
        
        print("\n=== Available Voices ===")
        for i in range(voices.Count):
            voice_info = voices.Item(i)
            name = voice_info.GetDescription()
            voice_id = voice_info.Id
            print(f"{i+1}. {name}")
            print(f"   ID: {voice_id}")
            print("-" * 40)
        print("=" * 40)
        
        return voices
    except Exception as e:
        logging.error(f"Error listing voices: {e}")
        return None

def setup_voice():
    """Initialize and configure the TTS voice"""
    global voice
    
    try:
        pythoncom.CoInitialize()
        voice = win32com.client.Dispatch("SAPI.SpVoice")
        settings = load_settings()
        
        # Set voice if specified
        if settings['voice_id']:
            voices = voice.GetVoices()
            for i in range(voices.Count):
                if voices.Item(i).Id == settings['voice_id']:
                    voice.Voice = voices.Item(i)
                    break
        
        # Set speech rate (-10 to 10)
        voice.Rate = int(settings['speech_rate'])
        
        # Set volume (0 to 100)
        voice.Volume = int(settings['volume'])
        
        logging.info(f"Voice configured: Rate={voice.Rate}, Volume={voice.Volume}")
        
    except Exception as e:
        logging.error(f"Error setting up voice: {e}")
        voice = None

def read_clipboard():
    """Read text from clipboard with retry logic"""
    logging.debug("read_clipboard() called")
    
    for attempt in range(3):
        try:
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(win32con.CF_TEXT):
                    text = win32clipboard.GetClipboardData(win32con.CF_TEXT)
                    if isinstance(text, bytes):
                        text = text.decode('utf-8', errors='ignore')
                    text = text.strip()
                    if text:
                        logging.debug(f"Clipboard text length: {len(text)}")
                        return text
                    else:
                        logging.debug(f"Clipboard attempt {attempt + 1}: empty")
                else:
                    logging.debug(f"Clipboard attempt {attempt + 1}: no text format")
            finally:
                win32clipboard.CloseClipboard()
        except Exception as e:
            logging.debug(f"Clipboard attempt {attempt + 1} failed: {e}")
        
        if attempt < 2:  # Don't sleep after the last attempt
            time.sleep(0.1)
    
    logging.info("Clipboard empty after 3 attempts – nothing to read.")
    return None

def speak_text(text):
    """Convert text to speech using SAPI with interruption support"""
    global voice
    
    if not voice:
        logging.error("Voice not initialized")
        return
    
    try:
        # Ensure COM is initialized for this thread
        pythoncom.CoInitialize()
        
        # Stop any current speech and clear the queue
        voice.Speak("", 3)  # SVSFlagsAsync | SVSFPurgeBeforeSpeak
        
        # Start new speech
        logging.info(f"Starting TTS for text of length {len(text)}")
        voice.Speak(text, 3)  # SVSFlagsAsync | SVSFPurgeBeforeSpeak
        
    except Exception as e:
        logging.error(f"Error in speak_text: {e}")
        # Try to reinitialize voice on COM error
        if "CoInitialize" in str(e) or "-2147221008" in str(e) or "-2147352567" in str(e):
            logging.warning("COM error detected - attempting to reinitialize voice")
            try:
                pythoncom.CoInitialize()
                voice = win32com.client.Dispatch("SAPI.SpVoice")
                settings = load_settings()
                voice.Rate = int(settings['speech_rate'])
                voice.Volume = int(settings['volume'])
                voice.Speak(text, 3)
                logging.info("Voice reinitialized successfully")
            except Exception as reinit_error:
                logging.error(f"Failed to reinitialize voice: {reinit_error}")

def handle_read_request():
    """Handle read clipboard request"""
    global last_hotkey_time
    last_hotkey_time = time.time()
    
    text = read_clipboard()
    if text:
        event_queue.put('read')
        print(f"Reading: {text}")
        speak_text(text)
    else:
        print("Clipboard empty – nothing to read.")

def handle_stop_request():
    """Handle stop speech request"""
    global last_hotkey_time
    last_hotkey_time = time.time()
    
    event_queue.put('stop')
    try:
        if voice:
            voice.Speak("", 3)  # Stop current speech
        print("🛑 Speech stopped")
        logging.info("Speech stopped by user")
    except Exception as e:
        logging.error(f"Error stopping speech: {e}")

def register_hotkeys():
    """Register global hotkeys using keyboard library with robust error handling"""
    global hotkey_handlers
    
    try:
        # Clear any existing hotkeys
        for handler in hotkey_handlers:
            try:
                keyboard.remove_hotkey(handler)
            except:
                pass
        hotkey_handlers.clear()
        
        # Register new hotkeys with multiple attempts
        for attempt in range(3):
            try:
                handler1 = keyboard.add_hotkey('alt+q', handle_read_request, suppress=True)
                handler2 = keyboard.add_hotkey('alt+shift+q', handle_stop_request, suppress=True)
                
                if handler1 and handler2:
                    hotkey_handlers = [handler1, handler2]
                    logging.info("Keyboard library hotkeys registered successfully")
                    return True
                else:
                    logging.warning(f"Hotkey registration attempt {attempt + 1} returned None")
                    
            except Exception as e:
                logging.warning(f"Hotkey registration attempt {attempt + 1} failed: {e}")
                
            if attempt < 2:
                time.sleep(0.5)  # Wait before retry
        
        logging.error("Failed to register hotkeys after 3 attempts")
        return False
        
    except Exception as e:
        logging.error(f"Error in register_hotkeys: {e}")
        return False

def unregister_hotkeys():
    """Unregister global hotkeys"""
    global hotkey_handlers
    
    try:
        for handler in hotkey_handlers:
            try:
                keyboard.remove_hotkey(handler)
                logging.debug(f"Removed hotkey handler: {handler}")
            except Exception as e:
                logging.warning(f"Error removing hotkey handler {handler}: {e}")
        
        hotkey_handlers.clear()
        logging.info("Hotkeys unregistered")
        
    except Exception as e:
        logging.error(f"Error unregistering hotkeys: {e}")

def check_for_enter_key():
    """Check for Enter key press (local fallback)"""
    if msvcrt.kbhit():
        key = msvcrt.getch()
        if key == b'\r':  # Enter key
            logging.info("Enter key pressed (local fallback)")
            print("🔄 Enter pressed - reading clipboard...")
            handle_read_request()

def test_hotkeys():
    """Test if hotkeys are still working"""
    try:
        # Simple test - check if keyboard module is responsive
        keyboard.is_pressed('alt')
        return True
    except Exception as e:
        logging.warning(f"Hotkey test failed: {e}")
        return False

def main():
    """Main application entry point"""
    global running
    
    print("Starting Clipboard to TTS Application...")
    print("\n=== Clipboard → TTS ===")
    logging.info("Application starting")
    
    # Initialize COM first
    pythoncom.CoInitialize()
    
    # List available voices
    list_available_voices()
    
    # Setup voice
    setup_voice()
    if not voice:
        print("❌ Failed to initialize TTS voice")
        return
    
    print("\nControls:")
    print(" alt+q  – read clipboard / interrupt & read new (GLOBAL)")
    print(" alt+shift+q – stop speech without reading (GLOBAL)")
    print(" Enter – read clipboard (LOCAL - works only in this window)")
    print(" Ctrl+C – exit")
    
    # Register hotkeys
    if register_hotkeys():
        print("✅ Global hotkeys registered successfully")
    else:
        print("⚠️  Global hotkeys failed - Enter key fallback available")
    
    print("Listening for hot-keys… (press Ctrl+C to quit)")
    logging.info("Main loop starting")
    
    last_heartbeat = time.time()
    last_hotkey_test = time.time()
    
    try:
        while running:
            # Check for Enter key (local fallback)
            check_for_enter_key()
            
            # Process event queue
            try:
                event = event_queue.get_nowait()
                logging.debug(f"Processing event: {event}")
            except queue.Empty:
                pass
            
            # Pump COM messages for SAPI
            pythoncom.PumpWaitingMessages()
            
            current_time = time.time()
            
            # Heartbeat every 30 seconds
            if current_time - last_heartbeat >= 30:
                time_since_hotkey = current_time - last_hotkey_time
                logging.debug(f"Heartbeat - Last hotkey: {time_since_hotkey:.1f}s ago")
                last_heartbeat = current_time
            
            # Test hotkeys every 60 seconds and re-register if needed
            if current_time - last_hotkey_test >= 60:
                if not test_hotkeys():
                    logging.warning("Hotkey test failed - attempting to re-register")
                    if register_hotkeys():
                        print("🔄 Hotkeys re-registered successfully")
                    else:
                        print("⚠️  Hotkey re-registration failed - Enter key still available")
                last_hotkey_test = current_time
            
            time.sleep(0.01)  # Small delay to prevent excessive CPU usage
            
    except KeyboardInterrupt:
        print("\n🔄 Shutting down...")
        logging.info("Application shutting down")
    finally:
        running = False
        unregister_hotkeys()
        if voice:
            try:
                voice.Speak("", 3)  # Stop any ongoing speech
            except:
                pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main() 