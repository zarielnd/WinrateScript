import keyboard
import win32gui
import win32con
import win32api
import time
import threading
from threading import Thread
import ctypes
from ctypes import wintypes

# Load user32.dll for Windows API calls
user32 = ctypes.WinDLL('user32', use_last_error=True)


# Define necessary structures and types for SendInput
class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG)),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG)),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", wintypes.DWORD),
        ("wParamL", wintypes.WORD),
        ("wParamH", wintypes.WORD),
    ]


class INPUT_union(ctypes.Union):
    _fields_ = [
        ("mi", MOUSEINPUT),
        ("ki", KEYBDINPUT),
        ("hi", HARDWAREINPUT),
    ]


class INPUT(ctypes.Structure):
    _fields_ = [
        ("type", wintypes.DWORD),
        ("union", INPUT_union),
    ]


# Constants
INPUT_KEYBOARD = 1
KEYEVENTF_KEYDOWN = 0x0000
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_SCANCODE = 0x0008
VK_P = 0x50  # Virtual key code for P
VK_RETURN = 0x0D  # Virtual key code for Enter

# Target application
TARGET_APP = "LimbusCompany"

# Global variables
running = True
auto_loop_active = False
loop_interval = 1.0  # Default interval in seconds


def find_limbus_window():
    """Find the LimbusCompany window handle"""
    result = []

    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd) and TARGET_APP in win32gui.GetWindowText(hwnd):
            windows.append(hwnd)
        return True

    win32gui.EnumWindows(callback, result)
    return result[0] if result else None


def send_input_key(vk_code, down=True):
    """Send keyboard input using SendInput API which works better for games"""
    extra = ctypes.c_ulong(0)
    ii_ = INPUT_union()
    ii_.ki = KEYBDINPUT(vk_code, 0, KEYEVENTF_KEYDOWN if down else KEYEVENTF_KEYUP, 0, ctypes.pointer(extra))
    x = INPUT(INPUT_KEYBOARD, ii_)
    user32.SendInput(1, ctypes.byref(x), ctypes.sizeof(x))


def send_to_game():
    """Send P+Enter to the game by setting it as foreground first"""
    limbus_hwnd = find_limbus_window()

    if not limbus_hwnd:
        print("LimbusCompany window not found")
        return False

    # Store the current foreground window
    current_hwnd = win32gui.GetForegroundWindow()

    try:
        # Set LimbusCompany as the foreground window
        win32gui.SetForegroundWindow(limbus_hwnd)
        time.sleep(0.1)  # Give time for window to come to foreground

        # Send P key down and up
        send_input_key(VK_P, down=True)
        time.sleep(0.05)
        send_input_key(VK_P, down=False)

        time.sleep(0.05)  # Small delay between keys

        # Send Enter key down and up
        send_input_key(VK_RETURN, down=True)
        time.sleep(0.05)
        send_input_key(VK_RETURN, down=False)

        time.sleep(0.1)  # Small delay before restoring window

        # Restore original foreground window
        win32gui.SetForegroundWindow(current_hwnd)

        return True
    except Exception as e:
        print(f"Error sending keys: {e}")
        # Attempt to restore original window
        try:
            win32gui.SetForegroundWindow(current_hwnd)
        except:
            pass
        return False


def auto_loop_function():
    """Function that runs in a separate thread to automatically send P+Enter at intervals"""
    global auto_loop_active

    print(f"Auto loop started - sending P+Enter every {loop_interval} seconds")
    while running and auto_loop_active:
        if send_to_game():
            print(f"P+Enter sent to LimbusCompany (Next in {loop_interval} seconds)")
        else:
            print("Failed to send P+Enter - LimbusCompany window not found or error occurred")
            time.sleep(5)  # Longer delay if error
            continue

        # Wait for the specified interval
        time.sleep(loop_interval)


def toggle_auto_loop():
    """Toggle the automatic loop on/off"""
    global auto_loop_active

    auto_loop_active = not auto_loop_active

    if auto_loop_active:
        # Start the auto loop in a new thread
        loop_thread = Thread(target=auto_loop_function, daemon=True)
        loop_thread.start()
        print("Auto loop ACTIVATED")
    else:
        print("Auto loop DEACTIVATED")


def increase_interval():
    """Increase the time interval between automatic presses"""
    global loop_interval
    loop_interval = min(loop_interval + 0.5, 10.0)
    print(f"Interval increased to {loop_interval} seconds")


def decrease_interval():
    """Decrease the time interval between automatic presses"""
    global loop_interval
    loop_interval = max(loop_interval - 0.5, 0.5)
    print(f"Interval decreased to {loop_interval} seconds")


def main():
    global running

    print(f"Starting P+Enter Auto-Clicker for {TARGET_APP}")
    print("\nThis script works by briefly bringing LimbusCompany to the foreground")
    print("to send the keystrokes, then restoring your previous window.")
    print("\nKEYBOARD CONTROLS:")
    print("  F6        - Toggle auto loop on/off")
    print("  F7        - Decrease interval (faster)")
    print("  F8        - Increase interval (slower)")
    print("  Alt+P     - Send P+Enter once (manual trigger)")
    print("  Ctrl+C    - Exit script (in command prompt)")
    print(f"\nCurrent interval: {loop_interval} seconds")

    # Register hotkeys
    keyboard.add_hotkey('f6', toggle_auto_loop)
    keyboard.add_hotkey('f7', decrease_interval)
    keyboard.add_hotkey('f8', increase_interval)
    keyboard.add_hotkey('alt+p', lambda: send_to_game())

    try:
        # Keep the main thread alive
        while True:
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("Script terminated by user")
    finally:
        # Clean up
        running = False
        keyboard.unhook_all()


if __name__ == "__main__":
    main()