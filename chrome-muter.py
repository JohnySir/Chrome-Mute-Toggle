# YouTube Clip Title: I Fixed The Bug In My Ultimate Lazy Mute Script.

import time
import msvcrt
import ctypes
from ctypes import wintypes
from datetime import datetime
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities

# --- Globals for Timestamping ---
last_toggle_time = None

# --- Windows API constants and structures for mouse input ---
STD_INPUT_HANDLE = -10
ENABLE_EXTENDED_FLAGS = 0x0080
ENABLE_MOUSE_INPUT = 0x0010

# Event types
KEY_EVENT = 0x0001
MOUSE_EVENT = 0x0002

# This was the line causing the error. 
# In some versions/environments, COORD is not a public member of wintypes.
# The error message correctly suggests using the internal '_COORD'.
class MOUSE_EVENT_RECORD(ctypes.Structure):
    _fields_ = [("dwMousePosition", wintypes._COORD),
                ("dwButtonState", wintypes.DWORD),
                ("dwControlKeyState", wintypes.DWORD),
                ("dwEventFlags", wintypes.DWORD)]

class KEY_EVENT_RECORD(ctypes.Structure):
    _fields_ = [("bKeyDown", wintypes.BOOL),
                ("wRepeatCount", wintypes.WORD),
                ("wVirtualKeyCode", wintypes.WORD),
                ("wVirtualScanCode", wintypes.WORD),
                ("uChar", wintypes.CHAR),
                ("dwControlKeyState", wintypes.DWORD)]

class INPUT_RECORD(ctypes.Structure):
    class _U(ctypes.Union):
        _fields_ = [("KeyEvent", KEY_EVENT_RECORD),
                    ("MouseEvent", MOUSE_EVENT_RECORD)]
    _fields_ = [("EventType", wintypes.WORD),
                ("Event", _U)]

# --- Core Audio Functions ---

def set_chrome_mute_status(mute_status):
    """
    Finds all active audio sessions for 'chrome.exe' and sets their mute status.

    Args:
        mute_status (bool): True to mute, False to unmute.

    Returns:
        bool: True if any Chrome sessions were found and modified, False otherwise.
    """
    sessions = AudioUtilities.GetAllSessions()
    chrome_found = False
    for session in sessions:
        if session.Process and session.Process.name() == "chrome.exe":
            chrome_found = True
            volume = session.SimpleAudioVolume
            volume.SetMute(mute_status, None)
    return chrome_found

def toggle_chrome_mute():
    """
    Toggles the mute status of all 'chrome.exe' audio sessions and logs the time.
    It finds the first Chrome session to determine the current state and then
    applies the opposite state to all Chrome sessions.
    """
    global last_toggle_time
    
    sessions = AudioUtilities.GetAllSessions()
    chrome_sessions = [s for s in sessions if s.Process and s.Process.name() == "chrome.exe"]

    if not chrome_sessions:
        print("-> Chrome is not running or not playing any audio.")
        return

    # Determine the current mute state from the first found Chrome session
    volume_interface = chrome_sessions[0].SimpleAudioVolume
    is_currently_muted = volume_interface.GetMute()
    new_mute_state = not is_currently_muted

    # Apply the new mute state to all Chrome sessions
    for session in chrome_sessions:
        session.SimpleAudioVolume.SetMute(new_mute_state, None)

    # --- Timestamp and Interval Logging ---
    current_time = datetime.now()
    interval_str = ""
    if last_toggle_time:
        interval = current_time - last_toggle_time
        # Format interval to be more readable
        total_seconds = interval.total_seconds()
        minutes, seconds = divmod(total_seconds, 60)
        interval_str = f" (Interval: {int(minutes)}m {seconds:.2f}s)"
    
    last_toggle_time = current_time
    timestamp_str = current_time.strftime('%Y-%m-%d %H:%M:%S')

    if new_mute_state:
        print(f"[{timestamp_str}] Chrome MUTED.{interval_str}")
    else:
        print(f"[{timestamp_str}] Chrome UNMUTED.{interval_str}")


def main():
    """
    Main function to run the mute toggle utility.
    Listens for spacebar presses, mouse clicks, and 'q' to quit.
    """
    print("=======================================")
    print("   Chrome Mute/Click Tracker by Johny's AI   ")
    print("=======================================")
    print("\nThis script runs in this window.")
    print("Press SPACEBAR or LEFT-CLICK to mute/unmute Chrome.")
    print("Press 'q' to quit the script.")
    print("\nWaiting for input...")

    # Get handle to the console input buffer
    h_in = ctypes.windll.kernel32.GetStdHandle(STD_INPUT_HANDLE)
    
    # Set console mode to enable mouse input
    # This is crucial for capturing click events
    ctypes.windll.kernel32.SetConsoleMode(h_in, ENABLE_EXTENDED_FLAGS | ENABLE_MOUSE_INPUT)

    input_record = INPUT_RECORD()
    num_read = wintypes.DWORD()

    try:
        while True:
            # Wait for and read a console input event
            ctypes.windll.kernel32.ReadConsoleInputW(h_in, ctypes.byref(input_record), 1, ctypes.byref(num_read))

            # --- Process Keyboard Input ---
            if input_record.EventType == KEY_EVENT and input_record.Event.KeyEvent.bKeyDown:
                key = input_record.Event.KeyEvent.uChar.lower()
                
                if key == ' ':
                    toggle_chrome_mute()
                elif key == 'q':
                    print("\nExiting script. Goodbye!")
                    break
            
            # --- Process Mouse Input ---
            elif input_record.EventType == MOUSE_EVENT:
                # Check for a left-button click (press, not release)
                if input_record.Event.MouseEvent.dwButtonState == 1 and input_record.Event.MouseEvent.dwEventFlags == 0:
                    toggle_chrome_mute()
            
            # Small delay to be kind to the CPU, though ReadConsoleInput is blocking
            time.sleep(0.05)

    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
        print("Please ensure you have the necessary libraries installed:")
        print("pip install pycaw comtypes")
    
    finally:
        # As a final cleanup, ensure Chrome is unmuted if the user wants
        try:
            choice = input("Do you want to leave Chrome unmuted? (y/n): ").lower()
            if choice == 'y':
                if set_chrome_mute_status(False):
                     print("Chrome has been unmuted.")
                else:
                     print("No active Chrome audio session found to unmute.")
        except Exception as e:
            print(f"Could not perform final unmute operation. Error: {e}")


if __name__ == "__main__":
    # Before starting, check for dependencies.
    try:
        from pycaw.pycaw import AudioUtilities
    except ImportError:
        print("Error: 'pycaw' library not found.")
        print("Please install it by running: pip install pycaw")
        exit()
        
    main()
