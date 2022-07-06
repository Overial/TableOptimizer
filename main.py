# Overial
# https://github.com/Overial

# importing necessary modules
import sys
import openpyxl
from openpyxl import Workbook
from pathlib import Path


def create_excel_file() -> Workbook:
    # create blank Workbook object
    workbook = openpyxl.Workbook()

    # get workbook active worksheet
    worksheet = workbook.active

    # set worksheet title name
    worksheet.title = 'current data'

    return workbook


def parse_excel_file(excel_file_name):
    print('initiated table optimization')

    # get input workbook
    input_workbook = openpyxl.load_workbook(f'input_data/{excel_file_name}')

    # get active input worksheet from input workbook
    input_worksheet = input_workbook.active

    # get output workbook
    output_workbook = openpyxl.load_workbook('output_data/optimized_table.xlsx')

    # get active output worksheet from output workbook
    output_worksheet = output_workbook.active

    # get active sheet title frmo input workbook
    input_worksheet_title = input_worksheet.title
    print(f'active sheet input worksheet title: {input_worksheet_title}')

    for row in range(3, input_worksheet.max_row):
        for col in range(2, input_worksheet.max_column + 1):
            input_cell = input_worksheet.cell(row, col)
            print(input_cell.value, end='\t\t\t')

            output_cell = output_worksheet.cell(row, col)
            output_cell.value = input_cell.value

        print('')

    output_workbook.save('output_data/optimized_table.xlsx')


# main function
def main() -> int:
    path = Path('output_data/optimized_table.xlsx')

    if path.is_file():
        print('optimized table already exists')
        excel_file_name = input('enter excel file name: ')
        parse_excel_file(excel_file_name)
    else:
        print('creating optimized table...')
        workbook = create_excel_file()
        workbook.save('output_data/optimized_table.xlsx')

    return 0


if __name__ == '__main__':
    sys.exit(main())
