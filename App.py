from tkinter import *
import tkinter as tk
from tkinter.ttk import *
import ctypes
import sys,os

from Views.MainMenu import *

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # self.title('MC Planing System Version 0.2.0') #26/05/2022
        # self.title('MC Planing System Version 0.2.1') #31/05/2022
        self.title('MC Planing System Version 0.2.2') #01/06/2022
        # self.state('zoomed')

        try:
            user32 = ctypes.windll.user32
            self.geometry(f'{int(user32.GetSystemMetrics(0)*0.985)}x{int(user32.GetSystemMetrics(1)*0.9)}+10+10')
        except:
            self.geometry(f'{int(self.winfo_screenwidth()*0.985)}x{int(self.winfo_screenheight()*0.9)}+10+10')
        
        # create a view and place it on the root window
        view = MainMenu(self)
        view.grid(row=0, column=0, padx=10, pady=10, sticky=tk.NW)

if __name__ == '__main__':
    app = App()
    app.mainloop()