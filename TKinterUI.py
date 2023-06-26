# Tkinter code that creates the UI for the program and handles the user input
#  

import tkinter as tk
import customtkinter
from tkinter import *
from tkinter import filedialog
import numpy as np
import sys

sys.path.append('../Backend')

backgroundColor = '#32373B'     # Dark Grey
containerColor = '#4A5859'       # Light Grey
textColor = '#F5C396'           # Light Orange

class MTR_EasyExcel_Tool(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MTR EasyExcel Tool")
        self.geometry("800x800")
        #self.resizable(False, False)
        customtkinter.set_appearance_mode("dark")

        self.grid_columnconfigure((0,1,2,4,5,6,7,8,9,10,11,12,13,14,15,16), weight=0)
        self.grid_rowconfigure((0,1,2,4,5,6,7,8,9,10,11,12,13,14,15,16), weight=0)
        self.grid_columnconfigure(0, weight=1, uniform="fred")

        self.Title = customtkinter.CTkLabel(self, text="MTR EasyExcel Tool", 
                                            font=("Arial", 36, "bold"), 
                                            fg_color=containerColor,
                                            text_color=textColor,
                                            corner_radius=10,
                                            height=100-20,
                                            width=600-20,
                                            padx=10,
                                            pady=10,)
        self.Title.grid(row=0, column=5, pady=10, padx=10, columnspan=12, rowspan=4)

        self.importButtonInstructions = customtkinter.CTkLabel(self, 
                                                               text="Import File:", 
                                                               font=("Arial", 20, "bold"), 
                                                               corner_radius=10, 
                                                               fg_color=containerColor,
                                                               text_color=textColor,
                                                               height=250-20,
                                                               width = 200-20,
                                                               padx = 10,
                                                               pady = 10,
                                                               anchor="n")
        self.importButtonInstructions.grid(row=0, column=0,  pady=10, padx=10, rowspan=12, columnspan=4)

        self.Pick_col_section = customtkinter.CTkLabel(self,
                                                        text="Pick Columns",
                                                        font=("Arial", 16, "bold"),
                                                        corner_radius=10,
                                                        fg_color=containerColor,
                                                        text_color=textColor,
                                                        height = 150-20,
                                                        width = 600-20,
                                                        padx=10,
                                                        pady=10,
                                                        anchor="n")
        self.Pick_col_section.grid(row=4, column=4,  pady=10, padx=10, rowspan=8, columnspan=12)

        self.File_Preview = customtkinter.CTkLabel(self,
                                                        text="File Preview",
                                                        font=("Arial", 16, "bold"),
                                                        corner_radius=10,
                                                        fg_color=containerColor,
                                                        text_color=textColor,
                                                        height = 350-20,
                                                        width = 800-20,
                                                        padx=10,
                                                        pady=10,
                                                        anchor="n")
        self.File_Preview.grid(row=12, column=0,  pady=10, padx=10, columnspan=16, rowspan=4)

        self.File_Preview = customtkinter.CTkLabel(self,
                                                        text="Populate Fields Buttons",
                                                        font=("Arial", 16, "bold"),
                                                        corner_radius=10,
                                                        fg_color=containerColor,
                                                        text_color=textColor,
                                                        height = 200-20,
                                                        width = 800-20,
                                                        padx=10,
                                                        pady=10,
                                                        anchor="n")
        self.File_Preview.grid(row=16, column=0,  pady=10, padx=10, columnspan=16, rowspan=4)


def main():
    window = MTR_EasyExcel_Tool()
    window.mainloop()

if __name__ == "__main__":
    main()


