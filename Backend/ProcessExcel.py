import openpyxl
from openpyxl import Workbook
import numpy as np

def print_array(input_array):
  	for row in input_array:
    		print(row)
                

def populate_array_from_excel(file_path, sheet_name):	# Function to copy Sheet1 into an array
    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]							

    result_array = []

    for row in sheet.iter_rows(values_only=True):		# Iterate over the rows and columns and append values to the array
        result_array.append(list(row))

    # Remove cells from the bottom of the array that only contain None values   
    while True:
        if result_array[-1] == [None] * len(result_array[-1]):
            result_array.pop()
        else:
            break

    print_array(result_array)        
    return result_array     #Return Array Values


def extract_columns(input_array, check_columns):		# Takes an input array and removes all non-check columns
    output_array = []
    
    for row in input_array:
        output_row = [row[index] for index in check_columns]
        output_array.append(output_row)

    return output_array									# Returns an array that only contains columns to check. Shifts columns to left so cannot just repopulate with this

def compare_arrays(array1, array2): 					# Compares two arrays and returns the values for indices in array 1 that are present in array2
    indexes = []										# Use this to find the location of array2 values in the original array1
    for value in array2:
        if value in array1:
            index = array1.index(value)
            indexes.append(index)
    return indexes


def get_unique_rows(input_array, check_columns):		# Removes all non-unique rows from an array, also records indices of unique rows in the original array using compare_arrays function
    Array_OnlyCheckColumns = extract_columns(input_array, check_columns)
    unique_rows_onlyCheckColumns = []
    unique_rows = []
    index_of_unique = []
    index = 0

    for row in Array_OnlyCheckColumns:
        if row not in unique_rows_onlyCheckColumns:
            unique_rows_onlyCheckColumns.append(row)
            unique_rows.append(input_array[index])
        index += 1

    index_of_unique = compare_arrays(unique_rows_onlyCheckColumns, Array_OnlyCheckColumns)
    
    return index_of_unique, unique_rows


def process_array(file_path, sheet_name, columns_to_check): # Just a general compilation of all of the prior functions, Takes in args and returns a processed sheet as well as indices of where the values need to go back
    DefaultState_Sheet = populate_array_from_excel(file_path, sheet_name)
    ProcessedState_UniqueIndices, ProcessedState_Sheet = get_unique_rows(DefaultState_Sheet, columns_to_check)
    
    return ProcessedState_Sheet, ProcessedState_UniqueIndices