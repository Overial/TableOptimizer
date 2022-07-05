# Overial
# https://github.com/Overial

# importing necessary modules
import sys
import openpyxl


# main function
def main() -> int:
    print('initiated table optimization')

    # get workbook
    filename = input('enter filename:')
    workbook = openpyxl.load_workbook(f'input_data/{filename}')

    # get active sheet
    worksheet = workbook.active

    # get active sheet title
    sheet_title = worksheet.title
    print(f'active sheet title: {sheet_title}')

    # iterate through rows and cols to print values
    for row in range(2, worksheet.max_row):
        for col in worksheet.iter_cols(2, worksheet.max_column):
            print(col[row].value, end="\t\t")
        print('')

    return 0


if __name__ == '__main__':
    sys.exit(main())
