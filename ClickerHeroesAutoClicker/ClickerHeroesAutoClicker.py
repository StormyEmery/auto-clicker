import win32gui, win32api, win32con
import ctypes, ctypes.wintypes
import time
import threading
import msvcrt
import pyHook, pythoncom
import warnings

from tkinter import *

class AutoClicker(threading.Thread):

    def __init__ (self):
        threading.Thread.__init__(self)
        self.setDaemon(True)
        self.cps = .1
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
                    if(self.old_window_name != self.window_name):
                        self.WindowExists()
                    if(self.window):
                        self.left_click()
                    time.sleep(self.cps)

                    if(self.stop): #break inner loop
                        break

                if(self.stop): #break outer loop
                        break

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
        self.windows = []

    def stop_callback(self):
        self.check_thread.stop = True
        self.auto_clicker.stop = True
        self.root.quit()

    def setup_gui(self):
        self.root = Tk()
        self.root.title("Storm's Auto-Clicker")
        self.root.iconbitmap('../Icons/clicker.ico') #../Icons/clicker.ico
        self.root.protocol("WM_DELETE_WINDOW", self.stop_callback)
        
        self.makeform()
         
    def run(self):
        self.setup_gui()
        self.root.mainloop()

    def validate_cps_entry(self, cps):
        valid_cps = True
        numeric_cps = .1
        try:
            numeric = float(cps)
            if(float(cps) < self.auto_clicker.min_cps or float(cps) > self.auto_clicker.max_cps):
                valid_cps = False
        except ValueError:
            valid_cps = False
       
        return valid_cps

    def validate_window_name(self, entry):
        valid_window_name = False
        for window in self.windows:
            text = entry[1].get()
            if(text.lower() == window.lower()):
                valid_window_name = True
                if(entry[1]['bg'] == 'red'):
                    entry[1]['bg'] = 'white'
                break
        return valid_window_name

    def handle_input_error(self, entry, msg):
        entry[1].delete(0,'end')
        entry[1]['bg'] = 'red'
        entry[1].insert(END, msg)

    def fetch(self, entries):
        count = 0 

        for entry in entries:
            field = entry[0]
            text  = entry[1].get()

            #window name entry
            if(count == 0):
                self.auto_clicker.old_window_name = self.auto_clicker.window_name
                self.auto_clicker.window_name = text
                valid_window_name = self.validate_window_name(entry)

                if(not valid_window_name):
                    self.handle_input_error(entry, "Window Cannot be Found")

            #click speed entry
            elif(count == 1):
                if(self.validate_cps_entry(text)):
                    if(entry[1]['bg'] == 'red'):
                            entry[1]['bg'] = 'white'
                    self.auto_clicker.cps = float(text)
                else:
                    self.handle_input_error(entry, "Value has to be between .05 and 86400")

            count += 1

        #since fetch() can be called multiple times, we don't want to try 
        #and start threads that are already alive.
        #so if they are already running, do nothing
        #else, if given a valid window name, start the other two threads
        if(not self.check_thread.is_alive() and not self.auto_clicker.is_alive()):
            if(valid_window_name):
                self.check_thread.start()
                self.auto_clicker.start()

    def get_all_window_titles(self, hwnd, l_param):
        window_title = win32gui.GetWindowText(hwnd)
        if(win32gui.IsWindowVisible(hwnd) and window_title != ""):
            self.windows.append(window_title)
            print(window_title)

    def make_entries(self):
        entries = []

        for field in self.fields:
            row = Frame(self.root)
            lab = Label(row, width=0, text=field, anchor='w')
            ent = Entry(row)
            row.pack(side=TOP, fill=X, padx=5, pady=5)
            lab.pack(side=LEFT)
            ent.pack(side=RIGHT, expand=YES, fill=X)
            if(field == "On/Off Keybind (Default is F2)"):
                ent.insert(END, "Currently Unavailable")
                ent.config(state=DISABLED)
            entries.append((field, ent))

        #Special entries for click position
        row = Frame(self.root)
        lab = Label(row, width=0, text='Click Position (x, y)', anchor='w')
        self.ent_x = Entry(row, width=5)
        self.ent_y = Entry(row, width=5)
        row.pack(side=TOP, anchor='w', pady=5)
        lab.pack(side=LEFT)
        self.ent_x.pack(side=LEFT, expand=YES, padx=2)
        self.ent_y.pack(side=LEFT, expand=YES, padx=2)
        entries.append(('x', self.ent_x))
        entries.append(('y', self.ent_y))

        return (entries, row)

    def make_buttons(self, entries):
        self.click_position_button = Button(entries[1], text='New Position', command=self.unlock_callback)
        self.click_position_button.pack(side=LEFT, padx=5, pady=5)

        save_button = Button(self.root, text='Save',
                command=(lambda e=entries[0]: self.fetch(e)))
        save_button.pack(side=LEFT, padx=5, pady=5)

        exit_button = Button(self.root, text='Quit', command=self.stop_callback)
        exit_button.pack(side=LEFT, padx=5, pady=5)

    def makeform(self):
        win32gui.EnumWindows(self.get_all_window_titles, None)
        
        entries = self.make_entries()
        self.make_buttons(entries)

        self.root.bind('<Return>', (lambda event, e=entries[0]: self.fetch(e)))  

    def unlock_callback(self):
        #sets background of these two entries to yellow
        #and disables the new click position button
        self.ent_x.delete(0, 'end')
        self.ent_y.delete(0, 'end')
        self.ent_x['bg'] = 'yellow'
        self.ent_y['bg'] = 'yellow'
        self.click_position_button.config(state=DISABLED)
        self.root.update()

        click_pos = self.mouse_input.get_mouse_positon()

        #inserts the detected click position into the fields
        self.auto_clicker.click_x = click_pos[0]
        self.auto_clicker.click_y = click_pos[1]
        self.ent_x.insert(END, click_pos[0])
        self.ent_y.insert(END, click_pos[1])
        self.ent_x['bg'] = 'white'
        self.ent_y['bg'] = 'white'
        self.click_position_button.config(state=ACTIVE)

        #If the window to be clicked has been found, 
        #calculate window offsets
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
    gui.start()
    
    gui.join()
    if(check_thread.is_alive() and auto_clicker.is_alive()):
        check_thread.join()
        auto_clicker.join()

    ctypes.windll.user32.PostQuitMessage(0)


if __name__ == "__main__":
    warnings.simplefilter('ignore')
    main()