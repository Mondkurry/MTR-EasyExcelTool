import sys
sys.path.append('Backend')
from ProcessExcel import populate_array_from_excel

import tkinter as tk
import customtkinter
from tkinter import *
from tkinter import filedialog as fd
import numpy as np

backgroundColor = '#32373B'     # Dark Grey
containerColor = '#4A5859'       # Light Grey
textColor = '#F5C396'           # Light Orange


class mainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry("200x100")
        self.resizable(False, False)
        customtkinter.set_appearance_mode("dark")

        self.grid_columnconfigure(
            (0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16), weight=0)
        self.grid_rowconfigure(
            (0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16), weight=0)
        
        self.fileSelection = FileSelection()