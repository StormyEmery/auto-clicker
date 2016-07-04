import win32gui, win32api, win32con
import ctypes, ctypes.wintypes
import time
import threading

class AutoClicker:

    def __init__ (self, cps, window_name):
        self.cps = cps;
        self.window_name = window_name;
        self.running = False;

    def WindowExists(self):
        try:
            window = win32gui.FindWindow(None, self.window_name)

        except win32gui.error:
            self.window = None
        else:
            self.window = window

    def start(self):

        self.WindowExists()

        if(self.window):
            while True:          
                while self.running:
                    self.left_click(self.window)
                    time.sleep(self.cps)

                time.sleep(0.2) #to keep CPU usage down while not auto-clicking
        else:
             print("Window NOT Found")
             quit()

    def left_click(self,window):
        rect = win32gui.GetWindowRect(window)

        x = rect[0]
        y = rect[1]
        w = rect[2] - x
        h = rect[3] - y
    
        pos = win32gui.ScreenToClient(window, (x + int((w) * .75), y + int((h) * .5)))
        lparam = win32api.MAKELONG(pos[0], pos[1])

        win32gui.PostMessage(window, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        win32gui.PostMessage(window, win32con.WM_LBUTTONUP, 0, lparam)

    #http://stackoverflow.com/questions/15777719/how-to-detect-key-press-when-the-console-window-has-lost-focus
    def check_for_stop(self):
        ctypes.windll.user32.RegisterHotKey(None, 1, 0, win32con.VK_F2)
        try:
            msg = ctypes.wintypes.MSG()
            while ctypes.windll.user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
                if msg.message == win32con.WM_HOTKEY:
                    self.running = not self.running
                ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
                ctypes.windll.user32.DispatchMessageA(ctypes.byref(msg))
        finally:
            ctypes.windll.user32.UnregisterHotKey(None, 1)


class CheckThread(threading.Thread):
    
    def __init__(self, auto_clicker):
        threading.Thread.__init__(self)
        self.auto_clicker = auto_clicker
        
    def run(self):
        self.auto_clicker.check_for_stop()


def main():
    print("Welcome to Storm's Auto-Clicker!")
    window_name = input("Name of window you want to auto-click: ")
    print("Window name: " + window_name)
    click_speed = input("How fast should I click? (in seconds): ")
    print("Click Speed: " + click_speed + "\n")
    print("Preparing...\n")
    time.sleep(3)
    print("To start/stop ultimate clickage, tap the 'F2' key!")              
    print("If you need to change the click speed or typed in the wrong window name, just close this application and restart it.")
    auto_clicker = AutoClicker(float(click_speed), window_name)
    check_thread = CheckThread(auto_clicker)

    check_thread.start()
    auto_clicker.start()
    

if __name__ == "__main__":
    main()