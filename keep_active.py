import pyautogui
import time
import random
import ctypes
from threading import Thread, Event

ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001

def prevent_sleep():
    # Call SetThreadExecutionState with ES_CONTINUOUS | ES_SYSTEM_REQUIRED
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)

def restore_settings():
    # Restore original power settings
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

def move_mouse(stop_event):
    pyautogui.FAILSAFE = False  # Correctly disable the fail-safe
    while not stop_event.is_set():
        # Get the current position of the mouse
        x, y = pyautogui.position()
        
        # Move the mouse to the right by 50 pixels
        pyautogui.moveTo(x + 50, y)
        
        # Sleep for 3 seconds
        for _ in range(3):
            if stop_event.is_set():
                return
            time.sleep(1)
        
        # Move the mouse back to the original position
        pyautogui.moveTo(x, y)

        # Wait for a random interval between 60 and 180 seconds
        sleep_time = random.randint(60, 180)
        print(f"Sleeping for {sleep_time} seconds")
        for _ in range(sleep_time):
            if stop_event.is_set():
                return
            time.sleep(1)

def press_space_randomly(stop_event):
    while not stop_event.is_set():
        # Wait for a random interval between 0 and 60 seconds
        sleep_time = random.randint(0, 60)
        for _ in range(sleep_time):
            if stop_event.is_set():
                return
            time.sleep(1)
        if stop_event.is_set():
            return
        pyautogui.press('space')
        print(f"Pressed space bar at {time.strftime('%Y-%m-%d %H:%M:%S')}")
        # Wait for the remaining time to make it 60 seconds in total
        for _ in range(60 - sleep_time):
            if stop_event.is_set():
                return
            time.sleep(1)

def scroll_mouse(stop_event):
    while not stop_event.is_set():
        # Wait for 30 seconds
        for _ in range(30):
            if stop_event.is_set():
                return
            time.sleep(1)
        if stop_event.is_set():
            return
        pyautogui.scroll(-3)
        print(f"Scrolled down 3 lines at {time.strftime('%Y-%m-%d %H:%M:%S')}")

if __name__ == "__main__":
    try:
        print("Preventing sleep. Press Ctrl+C to exit.")
        prevent_sleep()
        
        # Create an event to stop the threads
        stop_event = Event()
        
        # Start the move_mouse, press_space_randomly, and scroll_mouse functions in parallel
        mouse_thread = Thread(target=move_mouse, args=(stop_event,))
        space_thread = Thread(target=press_space_randomly, args=(stop_event,))
        scroll_thread = Thread(target=scroll_mouse, args=(stop_event,))

        mouse_thread.start()
        space_thread.start()
        scroll_thread.start()

        while True:
            time.sleep(1)
            if not (mouse_thread.is_alive() and space_thread.is_alive() and scroll_thread.is_alive()):
                break

    except KeyboardInterrupt:
        print("Exiting and restoring settings.")
        stop_event.set()
        mouse_thread.join()
        space_thread.join()
        scroll_thread.join()
        restore_settings()
 