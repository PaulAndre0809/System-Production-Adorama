import win32gui
import win32process

def callback(hwnd, extra):
    if win32gui.IsWindowVisible(hwnd):
        title = win32gui.GetWindowText(hwnd)
        classname = win32gui.GetClassName(hwnd)
        # Check if it's a dialog or contains excel
        if "excel" in title.lower() or "excel" in classname.lower() or classname == "#32770":
            _, win_pid = win32process.GetWindowThreadProcessId(hwnd)
            print(f"HWND: {hwnd}, PID: {win_pid}, Class: {classname}, Title: '{title}'")
    return True

def main():
    print("Listing visible Excel windows or Dialogs:")
    win32gui.EnumWindows(callback, None)

if __name__ == "__main__":
    main()
