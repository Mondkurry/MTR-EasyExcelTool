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
        self.resizable(False, False)
        customtkinter.set_appearance_mode("dark")
        self.grid_columnconfigure((0,1,2,4), weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=2)
        self.grid_rowconfigure(2, weight=2)
        self.grid_rowconfigure(3, weight=3)
        self.grid_rowconfigure(4, weight=2)
        self.grid_columnconfigure(0, weight=1, uniform="fred")

        self.Title = customtkinter.CTkLabel(self, text="MTR EasyExcel Tool", 
                                            font=("Arial", 28, "bold"), 
                                            fg_color=containerColor,
                                            text_color=textColor,
                                            corner_radius=10,
                                            height=((1/8) *800),
                                            width=(3/4)*800,
                                            padx=10,
                                            pady=10,)
        self.Title.grid(row=0, column=1, pady=10, padx=10, rowspan=1, columnspan=3)
        


        self.importButtonInstructions = customtkinter.CTkLabel(self, 
                                                               text="Import File:", 
                                                               font=("Arial", 20, "bold"), 
                                                               corner_radius=10, 
                                                               fg_color=containerColor,
                                                               text_color=textColor,
                                                               height=(3/8)*800,
                                                               width = 1/4*800,
                                                               padx = 10,
                                                               pady = 10,
                                                               anchor="n")
        self.importButtonInstructions.grid(row=0, column=0,  pady=10, padx=10, rowspan=3, columnspan=1)

        self.Pick_col_section = customtkinter.CTkLabel(self,
                                                        text="Pick Columns",
                                                        font=("Arial", 16, "bold"),
                                                        corner_radius=10,
                                                        fg_color=containerColor,
                                                        text_color=textColor,
                                                        height = (1/4)*800,
                                                        width = (3/4)*800,
                                                        padx=10,
                                                        pady=10,
                                                        anchor="n")
        self.Pick_col_section.grid(row=1, column=1,  pady=10, padx=10, rowspan=2, columnspan=3)

        self.File_Preview = customtkinter.CTkLabel(self,
                                                        text="File Preview",
                                                        font=("Arial", 16, "bold"),
                                                        corner_radius=10,
                                                        fg_color=containerColor,
                                                        text_color=textColor,
                                                        height = 6/16*800,
                                                        width = 800,
                                                        padx=10,
                                                        pady=10,
                                                        anchor="n")
        self.File_Preview.grid(row=3, column=0,  pady=10, padx=10, columnspan=4)

        self.File_Preview = customtkinter.CTkLabel(self,
                                                        text="Populate Fields Buttons",
                                                        font=("Arial", 16, "bold"),
                                                        corner_radius=10,
                                                        fg_color=containerColor,
                                                        text_color=textColor,
                                                        height = 1/4*800,
                                                        width = 800,
                                                        padx=10,
                                                        pady=10,
                                                        anchor="n")
        self.File_Preview.grid(row=4, column=0,  pady=10, padx=10, columnspan=4)
        
        # # code to import a file using a button in tkinter labeled "import" the name of the current imported file should be printed right next to the button
        # self.import_button = customtkinter.CTkButton(self, 
        #                             width = 80,
        #                             height = 30,
        #                             text="Import",
        #                             font=("Arial", 16, "bold"), 
        #                             corner_radius=5,
        #                             fg_color = textColor,
        #                             text_color=backgroundColor,
                                    
        #                             )
        # self.import_button.grid(row=1, column=0,  pady=50, padx=10)

        # self.sheetNameText = tk.Label(self, text="Current File:", font=("Arial", 8))
        # self.sheetNameText.grid(row=2, column=0, pady=10, padx=10)

        # self.sheetNameText = tk.Label(self, text="Current File:", font=("Arial", 8))
        # self.sheetNameText.grid(row=3, column=0, pady=10, padx=10)


def main():
    window = MTR_EasyExcel_Tool()
    window.mainloop()

if __name__ == "__main__":
    main()


