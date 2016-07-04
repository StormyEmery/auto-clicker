import win32gui, win32api, win32con
import ctypes, ctypes.wintypes
import time
import threading
import msvcrt

from tkinter import *

class AutoClicker(threading.Thread):

    def __init__ (self):
        threading.Thread.__init__(self)
        self.setDaemon(True)
        self.cps = .1;
        self.window_name = ""
        self.old_window_name = self.window_name
        self.keybind = "F2"
        self.running = False
        self.stop = False
        self.check_thread = None

    def WindowExists(self):
        try:
            window = win32gui.FindWindow(None, self.window_name)

        except win32gui.error:
            self.window = None
        else:
            self.window = window

    def my_start(self):

        self.WindowExists()

        if(self.window):
            while True:          
                while self.running:
                    if(self.old_window_name != self.window_name):
                        self.WindowExists()
                    if(self.window):
                        self.left_click(self.window)
                    time.sleep(self.cps)

                    if(self.stop): #break inner loop
                        break;

                if(self.stop): #break outer loop
                        break;

                time.sleep(0.2) #to keep CPU usage down while not auto-clicking
                #mouse_pos = win32gui.GetCursorPos()
                #print("Mouse Pos: (%d, %d)" % (mouse_pos[0], mouse_pos[1]))

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

        msg = ctypes.wintypes.MSG()
        while not self.check_thread.stop:
            if ctypes.windll.user32.PeekMessageA(ctypes.byref(msg), None, 0, 0, 1) != 0:
                if msg.message == win32con.WM_HOTKEY:
                    self.running = not self.running
                ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
                ctypes.windll.user32.DispatchMessageA(ctypes.byref(msg))
            time.sleep(.2) #to keep CPU usage down

        ctypes.windll.user32.UnregisterHotKey(None, 1)
        return
     
    def run(self):
        self.my_start()


class CheckThread(threading.Thread):
    
    def __init__(self, auto_clicker):
        threading.Thread.__init__(self)
        self.setDaemon(True)
        self.auto_clicker = auto_clicker
        self.stop = False
        self.auto_clicker.check_thread = self
        
    def run(self):
        self.auto_clicker.check_for_stop()

class GUI(threading.Thread):

    def __init__(self, auto_clicker, check_thread):
        threading.Thread.__init__(self)
        self.setDaemon(True)
        self.auto_clicker = auto_clicker;
        self.check_thread = check_thread;
        self.fields = 'Window Name', 'Click Speed (in seconds)', 'On/Off Keybind (Default is F2)'

    def callback(self):
        self.check_thread.stop = True
        self.auto_clicker.stop = True
        self.root.quit()

    def run(self):
        self.root = Tk()
        self.root.protocol("WM_DELETE_WINDOW", self.callback)
        ents = self.makeform(self.root)
        self.root.bind('<Return>', (lambda event, e=ents: self.fetch(e)))   
        b1 = Button(self.root, text='Save',
                command=(lambda e=ents: self.fetch(e)))
        b1.pack(side=LEFT, padx=5, pady=5)
        b2 = Button(self.root, text='Quit', command=self.callback)
        b2.pack(side=LEFT, padx=5, pady=5)
        self.root.mainloop()

    def fetch(self, entries):
        count = 0 
        for entry in entries:
            field = entry[0]
            text  = entry[1].get()
            if(count == 0):
                self.auto_clicker.old_window_name = self.auto_clicker.window_name
                self.auto_clicker.window_name = text
            elif(count == 1):
                self.auto_clicker.cps = float(text)

            count += 1

        if(not self.check_thread.is_alive() and not self.auto_clicker.is_alive()):
            self.check_thread.start()
            self.auto_clicker.start()

    def makeform(self, root):
       entries = []
       for field in self.fields:
          row = Frame(root)
          lab = Label(row, width=0, text=field, anchor='w')
          ent = Entry(row)
          row.pack(side=TOP, fill=X, padx=5, pady=5)
          lab.pack(side=LEFT)
          ent.pack(side=RIGHT, expand=YES, fill=X)
          entries.append((field, ent))
       return entries

def main():
    auto_clicker = AutoClicker()
    check_thread = CheckThread(auto_clicker)
    gui = GUI(auto_clicker, check_thread)
    gui.start()
    
    gui.join()
    check_thread.join()
    auto_clicker.join()
    
    

if __name__ == "__main__":
    main()