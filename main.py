import time
import win32com.client

def create_and_save_notepad_file():
    # Create an instance of the Notepad application
    notepad = win32com.client.Dispatch("WScript.Shell")

    # Open Notepad
    notepad.Run("chrome.exe")
    
    # # Wait for Notepad to open (adjust this delay as needed)
    notepad.AppActivate("Notepad")
    time.sleep(2)
    
    # # Write some text
    notepad.SendKeys("Hello, this is a sample text.")

    # # # Save the file with the name "moshe.txt"
    # time.sleep(2)
    # notepad.SendKeys("^s")  # Press Ctrl+S for save
    # time.sleep(2)
    # notepad.SendKeys("moshe.txt")
    # time.sleep(2)
    notepad.SendKeys("~")   # Press Enter to confirm the filename

if __name__ == "__main__":
    create_and_save_notepad_file()
