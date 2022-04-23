from tkinter import *
import tkinter as tk
from tkinter.ttk import *

from Views.MainMenu import *

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title('MC Planing System Version 0.1.0')
        self.state('zoomed')
        self.geometry('1440x810+10+10')

        # create a view and place it on the root window
        view = MainMenu(self)
        view.grid(row=0, column=0, padx=10, pady=10)

if __name__ == '__main__':
    app = App()
    app.mainloop()