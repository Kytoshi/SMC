import pyautogui
import time
from pynput import keyboard

# Global flag to stop the program
stop_flag = False

def clear_boxes(data):
    """Deleting Count in Input."""
    # Type the data as a string
    pyautogui.typewrite(str(data))
    time.sleep(0.1)
    pyautogui.press("enter")  # Press Enter to press confirm button
    time.sleep(0.1)
    pyautogui.press("esc")  # Press esc to dismiss confirm dialog box

def on_press(key):
    """Handle key press events."""
    global stop_flag
    try:
        if key.char == 'q':  # Stop the program when 'q' is pressed
            pyautogui.press("esc")
            stop_flag = True
            print("Cancel signal received. Stopping program...")
            return False  # Stop the listener
    except AttributeError:
        pass

def main():

    global stop_flag

    print("IMPORTANT: Please have the Physical Inventory App OPEN and READY with the correct Location loaded in.")
    print("...")
    print("To stop the program, press 'q' at any time.")
    print("===================================")
    shutdown = int(input("Please Enter how many boxes you are clearing: "))
    time.sleep(1)
    try:
        # Start listening for key presses
        listener = keyboard.Listener(on_press=on_press)
        listener.start()

        print("Please select the starting box in the Physical Inventory Application...")
        time.sleep(.2)
        print("Loading...")
        time.sleep(2)
        print("...")
        print("Beginning clearing in 3 seconds.")
        time.sleep(1)
        print("3")
        time.sleep(1)
        print("2")
        time.sleep(1)
        print("1")

        print("===================================")
        print("AMOUNT OF BOXES CLEARED")
        print("===================================")
        for count in range(shutdown):
            if stop_flag:
                print("Process stopped by user.")
                break
            clear_boxes(-1)
            print(f"{count + 1} cleared")
            
            time.sleep(0.1)  # Optional: Adjust time delay between inputs if needed
            

        # Stop the listener
        listener.stop()
        print("===================================")
        print(f"{count + 1} boxes have been cleared.")
        if not stop_flag:
            print("===================================")
            print("All data removed successfully.")
        print("===================================")
        while True:
            user_input = input("Enter 'exit' to close the program: ")
            if user_input.lower() == 'exit':
                print("Exiting program...")
                break

    except Exception as e:
        print(f"Error: {e}")
    

if __name__ == "__main__":
    main()
