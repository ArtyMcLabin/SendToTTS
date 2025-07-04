# SendToTTS Application v1.1.1
# Reads clipboard content and converts to speech using Windows SAPI
# Global hotkeys: Alt+Q (read/interrupt), Alt+Shift+Q (stop only)
# Runs completely windowless in system tray by default, use --debug for console mode

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
import re
import argparse
import win32gui
import win32con
import pystray
from PIL import Image, ImageDraw
import ctypes
from ctypes import wintypes

# Logging will be configured in main() based on debug mode

# Global variables
voice = None
event_queue = queue.Queue()
last_hotkey_time = time.time()
running = True
hotkey_handlers = []
debug_mode = False
tray_icon = None

def load_settings():
    """Load settings from settings.ini"""
    config = configparser.ConfigParser()
    
    # Default settings (removed voice_id since we auto-detect language)
    defaults = {
        'speech_rate': '0',
        'volume': '100'
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
        
        # Apply initial settings (will be reapplied when voice changes)
        apply_voice_settings()
        
        logging.info(f"Voice configured: Rate={voice.Rate}, Volume={voice.Volume}")
        
    except Exception as e:
        logging.error(f"Error setting up voice: {e}")
        voice = None

def apply_voice_settings():
    """Apply speech rate and volume settings to current voice"""
    global voice
    
    if not voice:
        return
    
    try:
        settings = load_settings()
        
        # Set speech rate (-10 to 10)
        voice.Rate = int(settings['speech_rate'])
        
        # Set volume (0 to 100)
        voice.Volume = int(settings['volume'])
        
        logging.debug(f"Applied settings: Rate={voice.Rate}, Volume={voice.Volume}")
        
    except Exception as e:
        logging.error(f"Error applying voice settings: {e}")

def read_clipboard():
    """Read text from clipboard with retry logic and proper Unicode support"""
    logging.debug("read_clipboard() called")
    
    for attempt in range(3):
        try:
            win32clipboard.OpenClipboard()
            try:
                # Try Unicode format first (CF_UNICODETEXT)
                if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                    text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                    text = text.strip()
                    if text:
                        logging.debug(f"Clipboard text length: {len(text)} (Unicode)")
                        # Log a safe representation of the text for debugging
                        safe_text = text.encode('ascii', errors='replace').decode('ascii')
                        logging.debug(f"Text preview: {safe_text[:50]}...")
                        return text
                    else:
                        logging.debug(f"Clipboard attempt {attempt + 1}: empty (Unicode)")
                # Fallback to regular text format
                elif win32clipboard.IsClipboardFormatAvailable(win32con.CF_TEXT):
                    text = win32clipboard.GetClipboardData(win32con.CF_TEXT)
                    if isinstance(text, bytes):
                        text = text.decode('utf-8', errors='ignore')
                    text = text.strip()
                    if text:
                        logging.debug(f"Clipboard text length: {len(text)} (ANSI)")
                        return text
                    else:
                        logging.debug(f"Clipboard attempt {attempt + 1}: empty (ANSI)")
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

def detect_language(text):
    """Detect language of text and return appropriate voice ID"""
    # Check for Cyrillic characters (Russian)
    cyrillic_match = re.search(r'[а-яё]', text.lower())
    if cyrillic_match:
        logging.info(f"Russian text detected (found: '{cyrillic_match.group()}')")
        return 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\TTS_MS_RU-RU_IRINA_11.0'
    
    # Check for Hebrew characters
    hebrew_match = re.search(r'[\u0590-\u05ff]', text)
    if hebrew_match:
        logging.info("Hebrew text detected, but no Hebrew voice available")
        return None
    
    # Default to English
    logging.info("English text detected (default)")
    return 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\TTS_MS_EN-US_ZIRA_11.0'

def set_voice_by_language(text):
    """Set the appropriate voice based on detected language"""
    global voice
    
    if not voice:
        return
    
    try:
        voice_id = detect_language(text)
        if voice_id:
            voices = voice.GetVoices()
            for i in range(voices.Count):
                if voices.Item(i).Id == voice_id:
                    voice.Voice = voices.Item(i)
                    logging.info(f"Switched to voice: {voices.Item(i).GetDescription()}")
                    
                    # Apply speech rate and volume settings to the new voice
                    apply_voice_settings()
                    break
    except Exception as e:
        logging.error(f"Error setting voice by language: {e}")

def speak_text(text):
    """Convert text to speech using SAPI with interruption support"""
    global voice
    
    if not voice:
        logging.error("Voice not initialized")
        return
    
    try:
        # Ensure COM is initialized for this thread
        pythoncom.CoInitialize()
        
        # Set appropriate voice based on language
        set_voice_by_language(text)
        
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
                # Set appropriate voice for the text and apply settings
                set_voice_by_language(text)
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
    
    # Skip if hotkeys are already registered
    if hotkey_handlers:
        logging.debug("Hotkeys already registered, skipping")
        return True
    
    try:
        # Ensure clean state - remove all hotkeys first
        try:
            keyboard.unhook_all_hotkeys()
        except:
            pass
        
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
        # Use unhook_all_hotkeys for clean removal
        keyboard.unhook_all_hotkeys()
        hotkey_handlers.clear()
        logging.info("Hotkeys unregistered")
        
    except Exception as e:
        logging.error(f"Error unregistering hotkeys: {e}")



def test_hotkeys():
    """Test if hotkeys are still working"""
    try:
        # Simple test - check if keyboard module is responsive
        keyboard.is_pressed('alt')
        return True
    except Exception as e:
        logging.warning(f"Hotkey test failed: {e}")
        return False

def create_tray_icon():
    """Create a simple icon for the system tray"""
    # Create a simple icon
    width = height = 64
    image = Image.new('RGB', (width, height), color='blue')
    draw = ImageDraw.Draw(image)
    
    # Draw a simple microphone icon
    draw.ellipse([16, 40, 48, 56], fill='white')  # Base
    draw.rectangle([28, 20, 36, 40], fill='white')  # Handle
    draw.ellipse([20, 8, 44, 32], fill='white')  # Mic head
    
    return image

def show_notification(title, message):
    """Show a system notification"""
    global tray_icon
    if tray_icon and not debug_mode:
        tray_icon.notify(message, title)
    elif debug_mode:
        print(f"{title}: {message}")

def quit_application():
    """Quit the application"""
    global running, tray_icon
    running = False
    unregister_hotkeys()
    if voice:
        try:
            voice.Speak("", 3)  # Stop any ongoing speech
        except:
            pass
    if tray_icon:
        tray_icon.stop()
    pythoncom.CoUninitialize()
    sys.exit(0)

def show_about():
    """Show about information"""
    about_text = """SendToTTS v1.1.0
    
Clipboard to Text-to-Speech converter
Supports Russian and English auto-detection

Hotkeys:
• Alt+Q - Read clipboard / interrupt & read new
• Alt+Shift+Q - Stop speech

Settings: Edit settings.ini to adjust speech rate and volume"""
    show_notification("About SendToTTS", about_text)

def create_tray_menu():
    """Create the system tray menu"""
    return pystray.Menu(
        pystray.MenuItem("SendToTTS v1.1.1", lambda: None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Read Clipboard (Alt+Q)", lambda: handle_read_request()),
        pystray.MenuItem("Stop Speech (Alt+Shift+Q)", lambda: handle_stop_request()),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("About", lambda: show_about()),
        pystray.MenuItem("Quit", lambda: quit_application())
    )

def setup_tray():
    """Setup system tray icon"""
    global tray_icon
    
    if debug_mode:
        return
    
    try:
        icon_image = create_tray_icon()
        tray_icon = pystray.Icon(
            "SendToTTS",
            icon_image,
            "SendToTTS - Clipboard to Speech",
            menu=create_tray_menu()
        )
        
        # Run tray icon in a separate thread
        tray_thread = threading.Thread(target=tray_icon.run, daemon=True)
        tray_thread.start()
        
        logging.info("System tray icon created")
        
    except Exception as e:
        logging.error(f"Failed to create system tray icon: {e}")
        # Continue without tray icon

def check_for_enter_key():
    """Check for Enter key press (local fallback) - only in debug mode"""
    if debug_mode and msvcrt.kbhit():
        key = msvcrt.getch()
        if key == b'\r':  # Enter key
            logging.info("Enter key pressed (local fallback)")
            print("🔄 Enter pressed - reading clipboard...")
            handle_read_request()

def hide_console_window():
    """Hide the console window for windowless operation"""
    if not debug_mode:
        # Get console window handle
        console_window = ctypes.windll.kernel32.GetConsoleWindow()
        if console_window:
            # Hide the console window
            ctypes.windll.user32.ShowWindow(console_window, 0)  # SW_HIDE = 0

def main():
    """Main application entry point"""
    global running, debug_mode
    
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='SendToTTS - Clipboard to Text-to-Speech')
    parser.add_argument('--debug', action='store_true', 
                       help='Run in debug mode with console window')
    args = parser.parse_args()
    
    debug_mode = args.debug
    
    # Hide console window if not in debug mode
    hide_console_window()
    
    # Configure logging based on mode
    if debug_mode:
        # Console mode - show everything and log to file
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('tts_debug.log'),
                logging.StreamHandler()
            ]
        )
        print("Starting Clipboard to TTS Application...")
        print("\n=== Clipboard → TTS ===")
    else:
        # Tray mode - no file logging, only console (which is hidden)
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.StreamHandler()
            ]
        )
    
    logging.info("Application starting")
    
    # Initialize COM first
    pythoncom.CoInitialize()
    
    # List available voices (only in debug mode)
    if debug_mode:
        list_available_voices()
    
    # Setup voice
    setup_voice()
    if not voice:
        if debug_mode:
            print("❌ Failed to initialize TTS voice")
        logging.error("Failed to initialize TTS voice")
        return
    
    if debug_mode:
        print("\nControls:")
        print(" alt+q  – read clipboard / interrupt & read new (GLOBAL)")
        print(" alt+shift+q – stop speech without reading (GLOBAL)")
        print(" Enter – read clipboard (LOCAL - works only in this window)")
        print(" Ctrl+C – exit")
    
    # Register hotkeys
    if register_hotkeys():
        if debug_mode:
            print("✅ Global hotkeys registered successfully")
        logging.info("Global hotkeys registered successfully")
    else:
        if debug_mode:
            print("⚠️  Global hotkeys failed - Enter key fallback available")
        logging.warning("Global hotkeys failed")
    
    # Setup system tray (only in non-debug mode)
    setup_tray()
    
    if debug_mode:
        print("Listening for hot-keys… (press Ctrl+C to quit)")
    else:
        show_notification("SendToTTS Started", "Clipboard to speech is ready. Use Alt+Q to read clipboard.")
    
    logging.info("Main loop starting")
    
    last_heartbeat = time.time()
    last_hotkey_test = time.time()
    
    try:
        while running:
            # Check for Enter key (local fallback) - only in debug mode
            if debug_mode:
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
                    # Clear existing hotkeys first
                    unregister_hotkeys()
                    if register_hotkeys():
                        if debug_mode:
                            print("🔄 Hotkeys re-registered successfully")
                        logging.info("Hotkeys re-registered successfully")
                    else:
                        if debug_mode:
                            print("⚠️  Hotkey re-registration failed - Enter key still available")
                        logging.warning("Hotkey re-registration failed")
                else:
                    logging.debug("Hotkey test passed - no re-registration needed")
                last_hotkey_test = current_time
            
            time.sleep(0.01)  # Small delay to prevent excessive CPU usage
            
    except KeyboardInterrupt:
        if debug_mode:
            print("\n🔄 Shutting down...")
        logging.info("Application shutting down")
    finally:
        quit_application()

if __name__ == "__main__":
    main() 