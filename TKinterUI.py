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

        self.Title = tk.Label(self, text="MTR EasyExcel Tool: Tool for stream lining Valve Lists!", font=("Arial", 20))
        self.Title.grid(row=0, column=0, pady=10, padx=10)

        self.importButtonInstructions = customtkinter.CTkLabel(self, 
                                                               text="Enter the file path of the Excel file you want to process:", 
                                                               font=("Arial", 12), 
                                                               corner_radius=10, 
                                                               fg_color=containerColor,
                                                               text_color=textColor,
                                                               height=35,
                                                               width = 500,
                                                               padx = 10,
                                                               anchor = "w")
        
        self.importButtonInstructions.grid(row=1, column=0, pady=10, padx=10, sticky = "W", columnspan=2)
        
        # code to import a file using a button in tkinter labeled "import" the name of the current imported file should be printed right next to the button
        # self.import_button = tk.Button(self, text="Import", command=self.openFile)
        self.import_button = customtkinter.CTkButton(self, 
                                    width = 80,
                                    height = 25,
                                    text="Import",
                                    corner_radius=5,
                                    fg_color = textColor,
                                    text_color=backgroundColor,
                                    )
        self.import_button.grid(row=1, column=1, pady=10, padx=10, sticky = "E")

        self.sheetNameText = tk.Label(self, text="Enter the name of the sheet you want to process:", font=("Arial", 8))
        self.sheetNameText.grid(row=1, column=3, pady=10, padx=10)


def main():
    window = MTR_EasyExcel_Tool()
    window.mainloop()

if __name__ == "__main__":
    main()


