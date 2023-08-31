# -*- coding: utf-8 -*-

import numpy as np
import sys

sys.path.append('Backend')

from openpyxl.styles import PatternFill, Alignment
from openpyxl import load_workbook
from UpdateInputSheet import process_and_update_userinput_sheet
from Populate import populate_output_andReformat
from array import *

file_path = 'workbook1.xlsx'
default_sheet_name = 'Sheet1'
input_sheet_name = 'Input Sheet'
output_sheet_name = 'Output Sheet'

input_columns_to_check = [2,3,4,9]	# Specify the columns to concatenate (e.g., A=1, B=2, D=4


def print_array(input_array):
  	for row in input_array:
    		print(row)
		
def main():
	process_and_update_userinput_sheet(file_path, default_sheet_name, input_sheet_name, input_columns_to_check)
	
	while True:
		user_input = input("Enter 'generate spreadsheet': ")
		if user_input.lower() == "generate spreadsheet":
			populate_output_andReformat(file_path, default_sheet_name, input_sheet_name, output_sheet_name, input_columns_to_check)
			break
		else:
			print("Invalid input. Please try again.")


    

if __name__ == "__main__":
    main()
 