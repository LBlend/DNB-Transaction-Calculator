import xlrd
from os import listdir
from datetime import datetime


def main():
    sortword = input('Type the word that you want to sort the data by\n')
    print_transactions = enter_transactions()
    calculate(sortword, print_transactions)


def enter_transactions():
    """Gives the user the option to print transaction details"""

    transcations = input('\n\nPrint all transactions?\ny/n\n')
    if transcations.lower() == 'n':
        return False
    elif transcations.lower() == 'y':
        return True
    else:
        print('\nINVALID INPUT\n\n')
        return enter_transactions()


def calculate(sortword: str, print_transactions: bool):
    """Calculates the total amount"""

    total = 0

    if print_transactions is True:
        print('\nDate | Transaction | Amount')

    for i in listdir('input'):
        excel_file = xlrd.open_workbook(filename=f'input/{i}')
        excel_sheet = excel_file.sheet_by_index(0)
        for row in range(excel_sheet.nrows):
            date = excel_sheet.cell_value(row, 0)
            name = excel_sheet.cell_value(row, 1)
            amount = excel_sheet.cell_value(row, 3)
            if sortword.lower() not in name.lower() or amount == '':
                continue
            if print_transactions is True:
                date = datetime(*xlrd.xldate_as_tuple(date, excel_file.datemode)).strftime('%B %d %Y')
                print(f'{date} | {name} | {amount}')

            total += amount

    print(f'\n\nTOTAL: {round(total, 2)}')


main()  # Runs the script
