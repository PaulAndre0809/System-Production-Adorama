import win32gui
import win32process

def get_child_windows(hwnd):
    def callback(child_hwnd, child_hwnds):
        child_hwnds.append(child_hwnd)
        return True
    child_hwnds = []
    win32gui.EnumChildWindows(hwnd, callback, child_hwnds)
    return child_hwnds

def find_dialog_hwnd():
    dialogs = []
    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd):
            classname = win32gui.GetClassName(hwnd)
            title = win32gui.GetWindowText(hwnd)
            if classname == "#32770" and "Visual Basic" in title:
                dialogs.append(hwnd)
        return True
    win32gui.EnumWindows(callback, None)
    return dialogs[0] if dialogs else None

def main():
    target_hwnd = find_dialog_hwnd()
    if not target_hwnd:
        print("No Visual Basic dialog found.")
        return
        
    try:
        title = win32gui.GetWindowText(target_hwnd)
        classname = win32gui.GetClassName(target_hwnd)
        print(f"Parent window HWND: {target_hwnd}, Class: {classname}, Title: '{title}'")
        
        children = get_child_windows(target_hwnd)
        print("Child controls text:")
        for child in children:
            child_text = win32gui.GetWindowText(child)
            child_class = win32gui.GetClassName(child)
            if child_text.strip():
                print(f"  Control HWND: {child}, Class: {child_class}, Text: '{child_text}'")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
