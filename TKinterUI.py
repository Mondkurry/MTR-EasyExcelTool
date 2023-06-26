# Tkinter code that creates the UI for the program and handles the user input
#  

import tkinter as tk
import customtkinter
from tkinter import *
from tkinter import filedialog as fd
import numpy as np
import sys

sys.path.append('../Backend')

backgroundColor = '#32373B'     # Dark Grey
containerColor = '#4A5859'       # Light Grey
textColor = '#F5C396'           # Light Orange

class MTR_EasyExcel_Tool(tk.Tk):

    
    def __init__(self):
        super().__init__()
        self.file_path = ""
        self.title("MTR EasyExcel Tool")
        self.geometry("800x800")
        #self.resizable(False, False)
        customtkinter.set_appearance_mode("dark")

        self.grid_columnconfigure((0,1,2,4,5,6,7,8,9,10,11,12,13,14,15,16), weight=0)
        self.grid_rowconfigure((0,1,2,4,5,6,7,8,9,10,11,12,13,14,15,16), weight=0)
        self.grid_columnconfigure(0, weight=1, uniform="fred")

        ################### CREATE BOUNDING BOXES FOR ALL APP ICONS ###################

        self.Title = customtkinter.CTkLabel(self, text="MTR EasyExcel Tool", 
                                            font=("Arial", 36, "bold"), 
                                            fg_color=containerColor,
                                            text_color=textColor,
                                            corner_radius=10,
                                            height=100-20,
                                            width=600-20,
                                            padx=10,
                                            pady=10,)
        self.Title.grid(row=0, column=5, pady=10, padx=10, rowspan=4, columnspan=12)

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
        self.File_Preview.grid(row=12, column=0,  pady=10, padx=10,  rowspan=4, columnspan=16)

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
        self.File_Preview.grid(row=16, column=0,  pady=10, padx=10, rowspan=4, columnspan=16)

        ################### IMPORT BUTTON STUFF ###################

        # Create the button that will open the file dialog
        self.import_button = customtkinter.CTkButton(self, 
                                    height = 30,
                                    text="Import",
                                    font=("Arial", 16, "bold"), 
                                    corner_radius=5,
                                    fg_color = textColor,
                                    text_color=backgroundColor,
                                    hover = DISABLED,
                                    command=self.open_file
                                    )
        self.import_button.grid(row=3, column=0,  pady=10, padx=10)

        self.CurrentFileText = customtkinter.CTkLabel(self,
                                                    text="Current File: ",
                                                    width = 150,
                                                    font=("Arial", 12, "bold"),
                                                    fg_color=containerColor,
                                                    text_color=textColor,
                                                    anchor="nw")
        self.CurrentFileText.grid(row=4, column=0, pady=10, padx=10)

        self.currentFile = customtkinter.CTkLabel(self,
                                                    text=self.readable_file_path,
                                                    width = 150,
                                                    font=("Arial", 16, "bold"),
                                                    fg_color=containerColor,
                                                    text_color=textColor,
                                                    anchor="nw")
        self.currentFile.grid(row=5, column=0, pady=10, padx=10)

        # Create the open_file function that will open the file dialog
      

        # Create the open_file function that will open the file dialog

    def set_file_path(self, file_path):
        self.file_path = file_path
    
    def get_file_path(self):
        return self.file_path
    
    def readable_file_path(self):
        return self.file_path.split(".xlsx")[-1]

    def open_file(self):
        file = fd.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.set_file_path(file)





def main():
    window = MTR_EasyExcel_Tool()
    window.mainloop()

if __name__ == "__main__":
    main()


