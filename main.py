import openpyxl
from openpyxl.utils import get_column_letter
from ProcessExcel import process_array


def print_array(input_array):
    for row in input_array:
        print(row)


def main():
    input_file_path = "workbook1.xlsx"
    input_sheet_name = "Sheet1"
    input_columns_to_check = [3, 4, 5, 6]  # Specify the columns to concatenate (e.g., A=1, B=2, D=4)

    ProcessedState_InputSheet = process_array(input_file_path, input_sheet_name, input_columns_to_check)

    print_array(ProcessedState_InputSheet)


if __name__ == "__main__":
    main()
