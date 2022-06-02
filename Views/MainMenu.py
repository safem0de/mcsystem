from threading import Thread
from tkinter import *
import tkinter as tk
from tkinter.ttk import *
from tkinter import filedialog as fd, font, ttk
from tkinter.messagebox import showinfo

import os, sys
import importlib
from turtle import width

from Controllers.ExcelController import *
import datetime

class MainMenu(Frame):

    excel = ExcelData()

    def __init__(self, parent):
        super().__init__(parent)

        if '_PYIBoot_SPLASH' in os.environ and importlib.util.find_spec("pyi_splash"):
            import pyi_splash
            pyi_splash.update_text('UI Loaded ...')
            pyi_splash.close()

        #create widgets
        self.labelheader = Label(self, text = 'MC Mecha2', font=("Impact", 25))
        self.labelheader.grid(row=0, column=0, sticky=tk.NW)

        self.style = ttk.Style()
        self.style.configure('TNotebook.Tab', font=('Bahnschrift SemiLight Condensed', 14))
        self.style.configure('big.TButton', font=('Bahnschrift SemiBold', 12), foreground = "blue4")

        # open file button
        self.open_button = Button(self,text='Open a File', command=lambda:select_file(), style='big.TButton')
        self.open_button.grid(row=0, column=1, sticky=tk.EW)

        # print(self.winfo_screenwidth(),self.winfo_screenheight())

        self.lf = LabelFrame(self, text='Datas prepared by Safem0de ')
        self.lf.grid(row=1, column=0, columnspan=100 ,sticky=tk.EW)

        self.alignments = ('Raw Data','On Hand', 'Daily Issue', 'Credits')
        self.nb = Notebook(self.lf)
        self.nb.grid(column=0, row=0, ipadx=10, ipady=10, sticky=tk.NSEW)

        self.f0 = Frame(self.nb, name=self.alignments[0].replace(" ","_").lower(), width=int(self.winfo_screenwidth()*0.7),height=int(self.winfo_screenheight()*0.6))
        self.f1 = Frame(self.nb, name=self.alignments[1].replace(" ","_").lower(), width=int(self.winfo_screenwidth()*0.7),height=int(self.winfo_screenheight()*0.6))
        self.f2 = Frame(self.nb, name=self.alignments[2].replace(" ","_").lower(), width=int(self.winfo_screenwidth()*0.7),height=int(self.winfo_screenheight()*0.6))
        self.f3 = Frame(self.nb, name=self.alignments[3].replace(" ","_").lower(), width=int(self.winfo_screenwidth()*0.7),height=int(self.winfo_screenheight()*0.6))

        self.nb.add(self.f0, text=self.alignments[0])
        self.nb.add(self.f1, text=self.alignments[1])
        self.nb.add(self.f2, text=self.alignments[2])
        self.nb.add(self.f3, text=self.alignments[3])

        ##################### ==== Credits ==== ######################
        # Update programing details and                              #
        # infomation of project maker                                #
        ############################################################## 

        try:
            self.f = open(file=os.path.join(sys._MEIPASS, "Credits.txt"), encoding="utf8")
            self.lblCredits = ttk.Label(self.f3, text=self.f.read(), font=(None, 9))
            self.lblCredits.grid(row=0, column=0, padx=3, pady=3, sticky=tk.NE)
        except:
            self.f = open("Credits.txt", encoding="utf8")
            self.lblCredits = ttk.Label(self.f3, text=self.f.read(), font=(None, 9))
            self.lblCredits.grid(row=0, column=0, padx=3, pady=3, sticky=tk.NE)    

        def add_OnHand_File():

            filetypes = (
                ('Excel file', '*.xlsx'),
                ('All files', '*.*')
            )

            filename = fd.askopenfilename(
                title = 'Open a file',
                initialdir = '/',
                filetypes = filetypes)

            if not filename == "":
                showinfo(
                    title = 'Selected File',
                    message = filename
                )
            else:
                showinfo(
                    title = 'Selected File',
                    message = 'File Not Found!!!'
                )
                return

            Thread(self.excel.createOnHandData()).start()
            Thread(self.excel.readExcelStock(filename)).start()

            self.excel.create_After_Issue()

            OnHand()

            self.canvas_set = Canvas(self.f2, width=int(self.nb.winfo_width()), height=int(self.nb.winfo_height()*0.8), highlightthickness=0)

            scroll_x_set = Scrollbar(self.f2, orient="horizontal", command=self.canvas_set.xview)
            scroll_y_set = Scrollbar(self.f2, orient="vertical", command=self.canvas_set.yview)

            scrollable_frame_set = ttk.Frame(self.canvas_set)
            scrollable_frame_set.bind(
                "<Configure>",
                lambda e: self.canvas_set.configure(
                    scrollregion=self.canvas_set.bbox("all")
                )
            )

            self.canvas_set.create_window((0, 0), window=scrollable_frame_set, anchor=tk.NW)
            self.canvas_set.configure(yscrollcommand=scroll_y_set.set, xscrollcommand=scroll_x_set.set)

            #### ======= Download Excel ===== ####
            self.Download_excel_btn = Button(self.f2, text='Download Excel File', command=lambda:selectFolder(), style='big.TButton')
            self.Download_excel_btn.grid(row=0, column=0, sticky=tk.NW)

            self.lf_RT_available = LabelFrame(scrollable_frame_set, text='Rotor จ่ายได้')
            self.lf_RT_available.grid(row=1, column=0, sticky=tk.W)

            self.lf_RT_notavailable = LabelFrame(scrollable_frame_set, text='Rotor จ่ายไม่ได้')
            self.lf_RT_notavailable.grid(row=2, column=0, sticky=tk.W)

            self.lf_ST_available = LabelFrame(scrollable_frame_set, text='Stator จ่ายได้')
            self.lf_ST_available.grid(row=1, column=1, sticky=tk.W)

            self.lf_ST_notavailable = LabelFrame(scrollable_frame_set, text='Stator จ่ายไม่ได้')
            self.lf_ST_notavailable.grid(row=2, column=1, sticky=tk.W)

            ### ======= Daily Stock Rotor Available ===== ####
            h1 = self.excel.createDailyHeader()
            self.tree_Daily_Rotor = Treeview(self.lf_RT_available, columns=h1, show='headings')

            for col in h1:
                self.tree_Daily_Rotor.heading(col, text = col)
                self.tree_Daily_Rotor.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createDailyIssue('rotor').values.tolist():
                self.tree_Daily_Rotor.insert('', tk.END, values=data)

            self.tree_Daily_Rotor.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            scrollbar_Daily_Rotor = Scrollbar(self.lf_RT_available, orient=tk.VERTICAL, command=self.tree_Daily_Rotor.yview)
            self.tree_Daily_Rotor.configure(yscroll=scrollbar_Daily_Rotor.set)
            scrollbar_Daily_Rotor.grid(row=0, column=1, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Daily Stock Stator Available ===== ####
            h2 = self.excel.createDailyHeader()
            self.tree_Daily_Stator = Treeview(self.lf_ST_available, columns=h2, show='headings')

            for col in h2:
                self.tree_Daily_Stator.heading(col, text = col)
                self.tree_Daily_Stator.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createDailyIssue('stator').values.tolist():
                self.tree_Daily_Stator.insert('', tk.END, values=data)

            self.tree_Daily_Stator.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            scrollbar_Daily_Stator = Scrollbar(self.lf_ST_available, orient=tk.VERTICAL, command=self.tree_Daily_Stator.yview)
            self.tree_Daily_Stator.configure(yscroll=scrollbar_Daily_Stator.set)
            scrollbar_Daily_Stator.grid(row=0, column=1, rowspan=20, pady=3, sticky=tk.NS)

            ### ======= Daily Stock Rotor Not Available ===== ####
            h3 = self.excel.createDailyHeader()
            self.tree_Shortage_Rotor = Treeview(self.lf_RT_notavailable, columns=h3, show='headings')

            for col in h3:
                self.tree_Shortage_Rotor.heading(col, text = col)
                self.tree_Shortage_Rotor.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createShortage('rotor').values.tolist():
                self.tree_Shortage_Rotor.insert('', tk.END, values=data)

            self.tree_Shortage_Rotor.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            scrollbar_Shortage_Rotor = Scrollbar(self.lf_RT_notavailable, orient=tk.VERTICAL, command=self.tree_Shortage_Rotor.yview)
            self.tree_Shortage_Rotor.configure(yscroll=scrollbar_Shortage_Rotor.set)
            scrollbar_Shortage_Rotor.grid(row=0, column=1, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Daily Stock Stator Not Available ===== ####
            h4 = self.excel.createDailyHeader()
            self.tree_Shortage_Stator = Treeview(self.lf_ST_notavailable, columns=h4, show='headings')

            for col in h4:
                self.tree_Shortage_Stator.heading(col, text = col)
                self.tree_Shortage_Stator.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createShortage('stator').values.tolist():
                self.tree_Shortage_Stator.insert('', tk.END, values=data)

            self.tree_Shortage_Stator.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            scrollbar_Shortage_Stator = Scrollbar(self.lf_ST_notavailable, orient=tk.VERTICAL, command=self.tree_Shortage_Stator.yview)
            self.tree_Shortage_Stator.configure(yscroll=scrollbar_Shortage_Stator.set)
            scrollbar_Shortage_Stator.grid(row=0, column=1, rowspan=20, pady=3, sticky=tk.NS)

            ################################################ CANVAS ################################################
            self.canvas_set.grid(row=1,column=0, sticky=tk.EW)
            scroll_x_set.grid(row=2, column=0, sticky=tk.EW+tk.S)
            scroll_y_set.grid(row=1, column=1, sticky=tk.NS+tk.E)
            ############################################ END OF CANVAS ############################################

            ##################### ==== Stock Page ==== ###################
            # Update programing details and                              #
            # infomation of project maker                                #
            ##############################################################

            #### ======= Shaft Stock ===== ####
            self.tree_Shaft_stock = Treeview(self.lf_Shaft_stock, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Shaft_stock.heading(col, text = col)
                self.tree_Shaft_stock.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createRequestPartData('shaft'):
                self.tree_Shaft_stock.insert('', tk.END, values=data)

            self.tree_Shaft_stock.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Shaft Need to Order ===== ####
            self.tree_Shaft_not_enough = Treeview(self.lf_Shaft_not_enough, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Shaft_not_enough.heading(col, text = col)
                self.tree_Shaft_not_enough.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createNeedToOrder('shaft'):
                self.tree_Shaft_not_enough.insert('', tk.END, values=data)

            self.tree_Shaft_not_enough.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Rotor Stack Stock ===== ####
            self.tree_Rotor_stock = Treeview(self.lf_Rotor_stock, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Rotor_stock.heading(col, text = col)
                self.tree_Rotor_stock.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createRequestPartData('rotor'):
                self.tree_Rotor_stock.insert('', tk.END, values=data)

            self.tree_Rotor_stock.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Rotor Stack Need to Order ===== ####
            self.tree_Rotor_not_enough = Treeview(self.lf_Rotor_not_enough, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Rotor_not_enough.heading(col, text = col)
                self.tree_Rotor_not_enough.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createNeedToOrder('rotor'):
                self.tree_Rotor_not_enough.insert('', tk.END, values=data)

            self.tree_Rotor_not_enough.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Magnet Stock ===== ####
            self.tree_Magnet_stock = Treeview(self.lf_Magnet_stock, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Magnet_stock.heading(col, text = col)
                self.tree_Magnet_stock.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createRequestPartData('magnet'):
                self.tree_Magnet_stock.insert('', tk.END, values=data)

            self.tree_Magnet_stock.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Magnet Need to Order ===== ####
            self.tree_Magnet_not_enough = Treeview(self.lf_Magnet_not_enough, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Magnet_not_enough.heading(col, text = col)
                self.tree_Magnet_not_enough.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createNeedToOrder('magnet'):
                self.tree_Magnet_not_enough.insert('', tk.END, values=data)

            self.tree_Magnet_not_enough.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Spacer Need to Order ===== ####
            self.tree_Spacer_stock = Treeview(self.lf_Spacer_stock, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Spacer_stock.heading(col, text = col)
                self.tree_Spacer_stock.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createRequestPartData('spacer'):
                self.tree_Spacer_stock.insert('', tk.END, values=data)

            self.tree_Spacer_stock.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Spacer Need to Order ===== ####
            self.tree_Spacer_not_enough = Treeview(self.lf_Spacer_not_enough, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Spacer_not_enough.heading(col, text = col)
                self.tree_Spacer_not_enough.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createNeedToOrder('spacer'):
                self.tree_Spacer_not_enough.insert('', tk.END, values=data)

            self.tree_Spacer_not_enough.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Stator Stock ===== ####
            self.tree_Stator_stock = Treeview(self.lf_Stator_stock, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Stator_stock.heading(col, text = col)
                self.tree_Stator_stock.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createRequestPartData('stator'):
                self.tree_Stator_stock.insert('', tk.END, values=data)

            self.tree_Stator_stock.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Stator Need to Order ===== ####
            self.tree_Stator_not_enough = Treeview(self.lf_Stator_not_enough, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Stator_not_enough.heading(col, text = col)
                self.tree_Stator_not_enough.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createNeedToOrder('stator'):
                self.tree_Stator_not_enough.insert('', tk.END, values=data)

            self.tree_Stator_not_enough.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Sap Stock ===== ####
            self.tree_Sap_stock = Treeview(self.lf_Sap_stock, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Sap_stock.heading(col, text = col)
                self.tree_Sap_stock.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createRequestPartData('sap'):
                self.tree_Sap_stock.insert('', tk.END, values=data)

            self.tree_Sap_stock.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Sap Need to Order ===== ####
            self.tree_Sap_not_enough = Treeview(self.lf_Sap_not_enough, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Sap_not_enough.heading(col, text = col)
                self.tree_Sap_not_enough.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createNeedToOrder('sap'):
                self.tree_Sap_not_enough.insert('', tk.END, values=data)

            self.tree_Sap_not_enough.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            ############################################ SEARCH ############################################

            self.lblSearch = Label(self.f1, text = 'Search Parts: ', font=('Bahnschrift SemiBold', 12), foreground = "blue4")
            self.lblSearch.grid(row=0, column=1, sticky=tk.NS+tk.E)

            t_list = [self.tree_Shaft, self.tree_Shaft_stock, self.tree_Shaft_not_enough,
                    self.tree_Rotor,self.tree_Rotor_stock,self.tree_Rotor_not_enough,
                    self.tree_Magnet,self.tree_Magnet_stock,self.tree_Magnet_not_enough,
                    self.tree_Spacer,self.tree_Spacer_stock,self.tree_Spacer_not_enough,
                    self.tree_Stator,self.tree_Stator_stock,self.tree_Stator_not_enough,
                    self.tree_Sap,self.tree_Sap_stock,self.tree_Sap_not_enough
                    ]

            self.search = tk.StringVar()
            self.txtSearch = Entry(self.f1, textvariable=self.search)
            self.txtSearch.grid(row=0, column=2, sticky=tk.NSEW)
            self.txtSearch.bind("<Return>", lambda e: search(e, self.search, t_list))

            self.Search_btn = Button(self.f1, text='Search', command=lambda e:search(e, self.search, t_list), style='big.TButton')
            self.Search_btn.grid(row=0, column=3, sticky=tk.NS+tk.W)

            ############################################ SEARCH ############################################

        def select_file():

            filetypes = (
                ('Excel file', '*.xlsx'),
                ('All files', '*.*')
            )

            filename = fd.askopenfilename(
                title = 'Open a file',
                initialdir = '/',
                filetypes = filetypes)

            if not filename == "":
                showinfo(
                    title = 'Selected File',
                    message = filename
                )
            else:
                showinfo(
                    title = 'Selected File',
                    message = 'File Not Found!!!'
                )
                return

            self.excel.readExcel(filename)
            self.columns = self.excel.createRawDataHeader()
            self.on_hand_columns = ('Item No.','Qty')
            datas = self.excel.createRawData()

            self.tree_rawData = Treeview(self.f0, columns=self.columns, show='headings',height=20)

            for col in self.columns:
                self.tree_rawData.heading(col, text = col)
                self.tree_rawData.column(col, minwidth=0, width=90, anchor=tk.E)

            for data in datas:
                self.tree_rawData.insert('', tk.END, values=data)

            self.tree_rawData.grid(row=0, column=0, pady=3, sticky=tk.NSEW)

            self.scrollbar = Scrollbar(self.f0, orient=tk.VERTICAL, command=self.tree_rawData.yview)
            self.tree_rawData.configure(yscroll=self.scrollbar.set)
            self.scrollbar.grid(row=0, column=1, pady=3, sticky=tk.NS)

            self.Add_Stock_btn = Button(self.f1, text='Add On Hand from File', command=lambda:add_OnHand_File(), style='big.TButton')
            self.Add_Stock_btn.grid(row=0, column=0, sticky=tk.NW)

        def OnHand():

            self.canvas = Canvas(self.f1, width=int(self.nb.winfo_width()), height=int(self.nb.winfo_height()*0.8), highlightthickness=0)

            scroll_x = Scrollbar(self.f1, orient="horizontal", command=self.canvas.xview)
            scroll_y = Scrollbar(self.f1, orient="vertical", command=self.canvas.yview)

            scrollable_frame = ttk.Frame(self.canvas)
            scrollable_frame.bind(
                "<Configure>",
                lambda e: self.canvas.configure(
                    scrollregion=self.canvas.bbox("all")
                )
            )

            self.canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
            self.canvas.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

            self.lf_Shaft = LabelFrame(scrollable_frame, text='Shaft')
            self.lf_Shaft.grid(row=1, column=0, sticky=tk.W)

            self.lf_Shaft_stock = LabelFrame(scrollable_frame, text='Shaft Stock Balance !!')
            self.lf_Shaft_stock.grid(row=2, column=0, sticky=tk.W)

            self.lf_Shaft_not_enough = LabelFrame(scrollable_frame, text='Shaft need2order !!')
            self.lf_Shaft_not_enough.grid(row=3, column=0, sticky=tk.W)

            self.lf_Rotor = LabelFrame(scrollable_frame, text='Rotor Stack')
            self.lf_Rotor.grid(row=1, column=1, sticky=tk.W)

            self.lf_Rotor_stock = LabelFrame(scrollable_frame, text='Rotor Stack Stock Balance !!')
            self.lf_Rotor_stock.grid(row=2, column=1, sticky=tk.W)

            self.lf_Rotor_not_enough = LabelFrame(scrollable_frame, text='Rotor Stack need2order !!')
            self.lf_Rotor_not_enough.grid(row=3, column=1, sticky=tk.W)

            self.lf_Magnet = LabelFrame(scrollable_frame, text='Magnet')
            self.lf_Magnet.grid(row=1, column=2, sticky=tk.W)

            self.lf_Magnet_stock = LabelFrame(scrollable_frame, text='Magnet Stock Balance !!')
            self.lf_Magnet_stock.grid(row=2, column=2, sticky=tk.W)

            self.lf_Magnet_not_enough = LabelFrame(scrollable_frame, text='Magnet need2order !!')
            self.lf_Magnet_not_enough.grid(row=3, column=2, sticky=tk.W)
            
            self.lf_Spacer = LabelFrame(scrollable_frame, text='Spacer')
            self.lf_Spacer.grid(row=1, column=3, sticky=tk.W)

            self.lf_Spacer_stock = LabelFrame(scrollable_frame, text='Spacer Stock Balance !!')
            self.lf_Spacer_stock.grid(row=2, column=3, sticky=tk.W)

            self.lf_Spacer_not_enough = LabelFrame(scrollable_frame, text='Spacer need2order !!')
            self.lf_Spacer_not_enough.grid(row=3, column=3, sticky=tk.W)

            self.lf_Stator = LabelFrame(scrollable_frame, text='Stator Stack')
            self.lf_Stator.grid(row=1, column=4, sticky=tk.W)

            self.lf_Stator_stock = LabelFrame(scrollable_frame, text='Stator Stack Stock Balance !!')
            self.lf_Stator_stock.grid(row=2, column=4, sticky=tk.W)

            self.lf_Stator_not_enough = LabelFrame(scrollable_frame, text='Stator Stack need2order !!')
            self.lf_Stator_not_enough.grid(row=3, column=4, sticky=tk.W)

            self.lf_Sap = LabelFrame(scrollable_frame, text='SAP No.')
            self.lf_Sap.grid(row=1, column=5, sticky=tk.W)

            self.lf_Sap_stock = LabelFrame(scrollable_frame, text='SAP No. Stock Balance !!')
            self.lf_Sap_stock.grid(row=2, column=5, sticky=tk.W)

            self.lf_Sap_not_enough = LabelFrame(scrollable_frame, text='SAP No. need2order !!')
            self.lf_Sap_not_enough.grid(row=3, column=5, sticky=tk.W)

            #### ======= Shaft OnHand Stock ===== ####

            self.tree_Shaft = Treeview(self.lf_Shaft, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Shaft.heading(col, text = col)
                self.tree_Shaft.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createOnHandData_type('shaft'):
                self.tree_Shaft.insert('', tk.END, values=data)

            self.tree_Shaft.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Rotor Stack OnHand Stock ===== ####

            self.tree_Rotor = Treeview(self.lf_Rotor, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Rotor.heading(col, text = col)
                self.tree_Rotor.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createOnHandData_type('rotor'):
                self.tree_Rotor.insert('', tk.END, values=data)

            self.tree_Rotor.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Magnet OnHand Stock ===== ####

            self.tree_Magnet = Treeview(self.lf_Magnet, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Magnet.heading(col, text = col)
                self.tree_Magnet.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createOnHandData_type('magnet'):
                self.tree_Magnet.insert('', tk.END, values=data)

            self.tree_Magnet.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Spacer OnHand Stock ===== ####

            self.tree_Spacer = Treeview(self.lf_Spacer, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Spacer.heading(col, text = col)
                self.tree_Spacer.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createOnHandData_type('spacer'):
                self.tree_Spacer.insert('', tk.END, values=data)

            self.tree_Spacer.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Stator Stack OnHand Stock ===== ####

            self.tree_Stator = Treeview(self.lf_Stator, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Stator.heading(col, text = col)
                self.tree_Stator.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createOnHandData_type('stator'):
                self.tree_Stator.insert('', tk.END, values=data)

            self.tree_Stator.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            #### ======= Sap OnHand Stock ===== ####

            self.tree_Sap = Treeview(self.lf_Sap, columns=self.on_hand_columns, show='headings')

            for col in self.on_hand_columns:
                self.tree_Sap.heading(col, text = col)
                self.tree_Sap.column(col, minwidth=0, width=90, stretch=False, anchor=tk.E)

            for data in self.excel.createOnHandData_type('sap'):
                self.tree_Sap.insert('', tk.END, values=data)

            self.tree_Sap.grid(row=0, column=0, rowspan=20, pady=3, sticky=tk.NS)

            ################################################ CANVAS ################################################
            self.canvas.grid(row=1,column=0, columnspan=20, sticky=tk.EW)
            scroll_x.grid(row=2, column=0, columnspan=20, sticky=tk.EW+tk.S)
            scroll_y.grid(row=1, column=1, columnspan=20, sticky=tk.NS+tk.E)
            ############################################ END OF CANVAS_ ############################################

        def selectFolder():
            filename =fd.askdirectory(
                title = 'Select Folder for Save Data',
                )

            if filename != "":
                showinfo(
                        title = 'File will save in this Folder',
                        message = filename
                    )
            else:
                showinfo(
                        title = 'File will save in this Folder',
                        message = 'No Folder Directory selected'
                    )
                return

            x = datetime.datetime.now()
            y = x.strftime("%d-%b-%y %H%M%S")

            data_frame1 = self.excel.createDailyIssue('rotor')
            data_frame2 = self.excel.createShortage('rotor')
            data_frame3 = self.excel.createDailyIssue('stator')
            data_frame4 = self.excel.createShortage('stator')

            with pd.ExcelWriter(f"{filename}\\MC_{y}.xlsx") as writer:
                ### use to_excel function and specify the sheet_name and index
                ### to store the dataframe in specified sheet
                Thread(data_frame1.to_excel(writer, sheet_name="Rotor จ่ายได้", index=False)).start()
                Thread(data_frame2.to_excel(writer, sheet_name="Rotor จ่ายไม่ได้", index=False)).start()
                Thread(data_frame3.to_excel(writer, sheet_name="Stator จ่ายได้", index=False)).start()
                Thread(data_frame4.to_excel(writer, sheet_name="Stator จ่ายไม่ได้", index=False)).start()

            showinfo(
                title = 'File Created',
                message = 'Happy working!! by Safem0de'
                )

        def search(event, search_entry : StringVar, treelist : list):
            query = search_entry.get()
            # print(query)
            for i in treelist:
                selections = []
                for child in i.get_children():
                    # print(tree.item(child)['values'][0])
                    if query.lower() in str(i.item(child)['values'][0]).lower():   # compare strings in  lower cases.
                        # print(tree.item(child)['values'][0])
                        selections.append(child)
                # print('search completed')
                i.selection_set(selections)

