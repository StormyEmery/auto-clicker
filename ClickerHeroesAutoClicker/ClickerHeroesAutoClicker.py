import win32gui, win32api, win32con
import ctypes, ctypes.wintypes
import time
import threading
import msvcrt
import pyHook, pythoncom

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
        self.min_cps = .05
        self.max_cps = 86400
        self.click_x = 0
        self.click_y = 0
        self.percentage_x = 0
        self.percentage_y = 0
        self.window = None

    def WindowExists(self):
        try:
            window = win32gui.FindWindow(None, self.window_name)

        except win32gui.error:
            self.window = None
        else:
            self.window = window
            self.calculate_percentages()

    def calculate_percentages(self):
        rect = win32gui.GetWindowRect(self.window)

        x = rect[0]
        y = rect[1]
        w = rect[2] - x
        h = rect[3] - y

        self.percentage_x = abs(x - self.click_x) / float(w)
        self.percentage_y = abs(y - self.click_y) / float(h)


    def my_start(self):

        self.WindowExists()
        self.old_window_name = self.window_name

        if(self.window):
            while True:          
                while self.running:
                    print("Old: %s\n New: %s\n\n" % (self.old_window_name, self.window_name))
                    if(self.old_window_name != self.window_name):
                        self.WindowExists()
                    if(self.window):
                        self.left_click()
                    time.sleep(self.cps)

                    if(self.stop): #break inner loop
                        break;

                if(self.stop): #break outer loop
                        break;

                time.sleep(0.2) #to keep CPU usage down while not auto-clicking

        else:
             print("Window NOT Found")


    def left_click(self):
        rect = win32gui.GetWindowRect(self.window)

        x = rect[0]
        y = rect[1]
        w = rect[2] - x
        h = rect[3] - y
    
        pos = win32gui.ScreenToClient(self.window, (x + int(w * self.percentage_x), y + int(h * self.percentage_y)))
        lparam = win32api.MAKELONG(pos[0], pos[1])

        win32gui.PostMessage(self.window, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        win32gui.PostMessage(self.window, win32con.WM_LBUTTONUP, 0, lparam)

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

    def __init__(self, auto_clicker, check_thread, mouse_input):
        threading.Thread.__init__(self)
        self.setDaemon(True)
        self.auto_clicker = auto_clicker
        self.check_thread = check_thread
        self.mouse_input = mouse_input
        self.fields = 'Window Name', 'Click Speed (in seconds)', 'On/Off Keybind (Default is F2)'
        self.start()
        self.windows = []

    def stop_callback(self):
        self.check_thread.stop = True
        self.auto_clicker.stop = True
        self.root.quit()

    def run(self):
        self.root = Tk()
        self.root.title("Storm's Auto-Clicker")
        self.root.iconbitmap('../Icons/clicker.ico')
        self.root.protocol("WM_DELETE_WINDOW", self.stop_callback)
        ents = self.makeform(self.root)
        self.root.bind('<Return>', (lambda event, e=ents: self.fetch(e)))   
        b1 = Button(self.root, text='Save',
                command=(lambda e=ents: self.fetch(e)))
        b1.pack(side=LEFT, padx=5, pady=5)
        b2 = Button(self.root, text='Quit', command=self.stop_callback)
        b2.pack(side=LEFT, padx=5, pady=5)
        self.root.mainloop()

    def validate_cps_entry(self, cps):
        if(cps < self.auto_clicker.min_cps or cps > self.auto_clicker.max_cps):
            return False
        return True

    def fetch(self, entries):
        count = 0 
        valid_window_name = False
        for entry in entries:
            field = entry[0]
            text  = entry[1].get()
            if(count == 0):
                self.auto_clicker.old_window_name = self.auto_clicker.window_name
                self.auto_clicker.window_name = text
                for window in self.windows:
                    if(text.lower() == window.lower()):
                        valid_window_name = True
                        if(entry[1]['bg'] == 'red'):
                            entry[1]['bg'] = 'white'
                        break
                if(not valid_window_name):
                    entry[1].delete(0,'end')
                    entry[1]['bg'] = 'red'
                    entry[1].insert(END, "Cannot Find Window")
            elif(count == 1):
                if(self.validate_cps_entry(float(text))):
                    if(entry[1]['bg'] == 'red'):
                            entry[1]['bg'] = 'white'
                    self.auto_clicker.cps = float(text)
                else:
                    entry[1].delete(0, 'end')
                    entry[1]['bg'] = 'red'
                    entry[1].insert(END, "Value has to be between .05 and 86400")

            count += 1

        if(not self.check_thread.is_alive() and not self.auto_clicker.is_alive()):
            if(valid_window_name):
                self.check_thread.start()
                self.auto_clicker.start()

    def get_all_window_titles(self, hwnd, l_param):
        window_title = win32gui.GetWindowText(hwnd)
        if(win32gui.IsWindowVisible(hwnd) and window_title != ""):
            self.windows.append(window_title)

    def makeform(self, root):
        win32gui.EnumWindows(self.get_all_window_titles, None)
        
        entries = []
        for field in self.fields:
            row = Frame(root)
            lab = Label(row, width=0, text=field, anchor='w')
            ent = Entry(row)
            row.pack(side=TOP, fill=X, padx=5, pady=5)
            lab.pack(side=LEFT)
            ent.pack(side=RIGHT, expand=YES, fill=X)
            if(field == "On/Off Keybind (Default is F2)"):
                ent.insert(END, "Currently Unavailable")
                ent.config(state=DISABLED)
            entries.append((field, ent))

        row = Frame(root)
        lab = Label(row, width=0, text='Click Position (x, y)', anchor='w')
        self.ent_x = Entry(row, width=5)
        self.ent_y = Entry(row, width=5)
        row.pack(side=TOP, anchor='w', pady=5)
        lab.pack(side=LEFT)
        self.ent_x.pack(side=LEFT, expand=YES, padx=2)
        self.ent_y.pack(side=LEFT, expand=YES, padx=2)
        entries.append(('x', self.ent_x))
        entries.append(('y', self.ent_y))
        self.b3 = Button(row, text='New Position', command=self.unlock_callback)
        self.b3.pack(side=LEFT, padx=5, pady=5)
        return entries

    def unlock_callback(self):
        print(self.ent_x.get())
        self.ent_x.delete(0, 'end')
        self.ent_y.delete(0, 'end')
        self.ent_x['bg'] = 'yellow'
        self.ent_y['bg'] = 'yellow'
        self.b3.config(state=DISABLED)
        self.root.update()
        click_pos = self.mouse_input.get_mouse_positon()
        self.auto_clicker.click_x = click_pos[0]
        self.auto_clicker.click_y = click_pos[1]
        self.ent_x.insert(END, click_pos[0])
        self.ent_y.insert(END, click_pos[1])
        self.ent_x['bg'] = 'white'
        self.ent_y['bg'] = 'white'
        self.b3.config(state=ACTIVE)
        if(self.auto_clicker.window):
            self.auto_clicker.calculate_percentages()


class MouseInput():

    def __init__(self):
        self.position = None
        self.active = True

    def get_mouse_positon(self):
        self.active = True
        self.position = None
        def onClick(event):
            if(self.active):
                self.position = event.Position
            return True
        
        hm = pyHook.HookManager()
        hm.SubscribeMouseAllButtonsDown(onClick)
        hm.HookMouse()
        hm.HookKeyboard() 
        while self.position == None:
            pythoncom.PumpWaitingMessages()
            time.sleep(.01)
        
        hm.UnhookMouse()
        hm.UnhookKeyboard()
        self.active = False
        return self.position


def main():
    mouse_input = MouseInput()
    auto_clicker = AutoClicker()
    check_thread = CheckThread(auto_clicker)
    gui = GUI(auto_clicker, check_thread, mouse_input)

    
    gui.join()
    if(check_thread.is_alive() and auto_clicker.is_alive()):
        check_thread.join()
        auto_clicker.join()
    

if __name__ == "__main__":
    main()