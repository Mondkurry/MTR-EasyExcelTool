# Tkinter code that creates the UI for the program and handles the user input
#  

import tkinter as tk
import customtkinter
from tkinter import *
from tkinter import filedialog as fd
import numpy as np
import sys
sys.path.append('Backend')
from ProcessExcel import populate_array_from_excel


backgroundColor = '#32373B'     # Dark Grey
containerColor = '#4A5859'       # Light Grey
textColor = '#F5C396'           # Light Orange


class MTR_EasyExcel_Tool(tk.Tk):

    
    def __init__(self):
        super().__init__()
        self.default_file_path = ""
        self.set_file_path(self.default_file_path)
        self.title("MTR EasyExcel Tool")
        self.geometry("800x800")
        #self.resizable(False, False)
        customtkinter.set_appearance_mode("dark")

        self.grid_columnconfigure((0,1,2,4,5,6,7,8,9,10,11,12,13,14,15,16), weight=0)
        self.grid_rowconfigure((0,1,2,4,5,6,7,8,9,10,11,12,13,14,15,16), weight=0)

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


    #--------------------------------------------------------------------------#
    #------------------------- FILE SELECTION MODULE --------------------------#
    #--------------------------------------------------------------------------#

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
        self.importButtonInstructions.grid(row=0, column=0,  pady=10, padx=10, rowspan=6, columnspan=4)

        self.import_button = customtkinter.CTkButton(self, 
                                    height = 30,
                                    width=150,
                                    text="Import",
                                    font=("Arial", 16, "bold"), 
                                    corner_radius=5,
                                    fg_color = textColor,
                                    text_color=backgroundColor,
                                    hover = DISABLED,
                                    command=self.open_file
                                    )
        self.import_button.grid(row=3, column=0,  pady=0, padx=25)

        self.CurrentFileText = customtkinter.CTkLabel(self,
                                                    text="Current File: ",
                                                    width = 150,
                                                    font=("Arial", 12, "bold"),
                                                    fg_color=containerColor,
                                                    text_color=textColor,
                                                    anchor="nw")
        self.CurrentFileText.grid(row=4, column=0, pady=5, padx=10)

        self.number_col = []
        for i in range(12):
            self.checkBox = customtkinter.CTkCheckBox(self, 
                                                        width=5,
                                                        height=5,
                                                        bg_color=containerColor,
                                                        text="",
                                                        # offvalue=0,
                                                        offvalue=self.set_col(i),
                                                        )
            self.checkBox.grid(row=5, column=(5+i), pady=10, padx=10)
        print(self.number_col)  
    
    def set_col(self, col):
        self.number_col.append(col)
        print(self.number_col)

    def set_file_path(self, set_file_path):
        self.file_path = set_file_path
    
    def get_file_path(self):
        return self.file_path
    
    def get_simple_file_path(self):
        return self.file_path.split("/")[-1]
    
    def set_sheet_name(self):
        self.default_sheet_name = self.sheet_name_input.get()
        print(self.get_sheet_name())
    
    def get_sheet_name(self):
        return self.default_sheet_name

    def open_file(self):
        file = fd.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.set_file_path(file)

        self.currentFile = customtkinter.CTkLabel(self,
                                                        text=self.get_simple_file_path(),
                                                        width = 150,
                                                        font=("Arial", 12, "bold"),
                                                        fg_color=containerColor,
                                                        text_color=textColor,
                                                        wraplength=150,
                                                        anchor="n")
        self.currentFile.grid(row=5, column=0, pady=5, padx=10)

        self.sheet_name_input = customtkinter.CTkEntry(self,
                                                        placeholder_text="Sheet Name",
                                                        width = 75,
                                                        height = 15,
                                                        font=("Arial", 10, "bold"),
                                                        fg_color=containerColor,
                                                        text_color=textColor,
                                                        )
        self.sheet_name_input.grid(row=6, column=0, pady=10, padx=10)

        self.confirm_sheet_name_button = customtkinter.CTkButton(self, 
                                    height = 20,
                                    width=75,
                                    text="Confirm Sheet Name",
                                    font=("Arial", 12, "bold"), 
                                    fg_color = textColor,
                                    text_color=backgroundColor,
                                    hover = DISABLED,
                                    command=self.set_sheet_name
                                    )
        self.confirm_sheet_name_button.grid(row=7, column=0,  pady=0, padx=25, sticky="n")

    #--------------------------------------------------------------------------#
    #------------------------ COLUMN SELECTION MODULE -------------------------#
    #--------------------------------------------------------------------------#

    def make_columns(self):
        num_columns_in_file = len(populate_array_from_excel(self.get_file_path(), self.get_sheet_name())[0])
        for i in range(num_columns_in_file):
            print(i)

                                                         
    
    














def main():
    window = MTR_EasyExcel_Tool()
    window.mainloop()

if __name__ == "__main__":
    main()


