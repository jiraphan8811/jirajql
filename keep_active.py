import pyautogui
import time
import random
import ctypes
import time


ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001

def prevent_sleep():
    # Call SetThreadExecutionState with ES_CONTINUOUS | ES_SYSTEM_REQUIRED
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)

def restore_settings():
    # Restore original power settings
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

def move_mouse():
    pyautogui.FAILSAFE = False  # Correctly disable the fail-safe
    while True:
        # Get the current position of the mouse
        x, y = pyautogui.position()
        
        # Move the mouse to the right by 50 pixels and back to original position
        pyautogui.moveTo(x + 50, y)
        pyautogui.moveTo(x, y)
        
        # Wait for a random interval between 1 and 3 minutes (60 to 180 seconds)
        sleep_time = random.randint(60, 180)
        print(f"Sleeping for {sleep_time} seconds")
        time.sleep(sleep_time)

if __name__ == "__main__":
    
    try:
        print("Preventing sleep. Press Ctrl+C to exit.")
        while True:
            prevent_sleep()
            move_mouse()
            time.sleep(60)  # Sleep for 60 seconds before refreshing the state
    except KeyboardInterrupt:
        print("Exiting and restoring settings.")
        restore_settings()