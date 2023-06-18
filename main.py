import openpyxl
from openpyxl.utils import get_column_letter

def populate_array_from_excel(file_path, sheet_name):
    # Load the workbook
    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]

    # Create an empty array
    result_array = []

    # Iterate over the rows and columns and append values to the array
    for row in sheet.iter_rows(values_only=True):
        result_array.append(list(row))

    return result_array


def extract_columns(input_array, column_indices):
    output_array = []
    
    for row in input_array:
        output_row = [row[index] for index in column_indices]
        output_array.append(output_row)
    
    return output_array


def print_array(input_array):
    for row in input_array:
        print(row)


def main():
    input_file_path = "worbook1.xlsx"
    input_sheet_name = "Sheet1"
    columns_to_check = [1, 2, 4]  # Specify the columns to concatenate (e.g., A=1, B=2, D=4)

    DefaultState_InputSheet = populate_array_from_excel(input_file_path, input_sheet_name)
    ProcessedState_InputSheet = extract_columns(DefaultState_InputSheet, columns_to_check)

    print_array(ProcessedState_InputSheet)


if __name__ == "__main__":
    main()
