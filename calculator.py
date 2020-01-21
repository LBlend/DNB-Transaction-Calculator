import xlrd
from os import listdir
from datetime import datetime


def main():
    sortword = input('Type the word that you want to sort the data by\n')
    income_or_expenses = received_or_paid()
    print_transactions = enter_transactions()
    calculate(sortword, income_or_expenses, print_transactions)


def received_or_paid():
    """Choose betweeen calculating income or expenses"""

    user_input = input('\n\nCalculate income or expenses?\ntype "i" for income / type "e" expenses\n')
    if user_input.lower() == 'i':
        return 4
    elif user_input.lower() == 'e':
        return 3
    else:
        print('\nINVALID INPUT\n\n')
        return received_or_paid()


def enter_transactions():
    """Gives the user the option to print transaction details"""

    user_input = input('\n\nPrint all transactions?\ntype "y" for yes / type "n" for no\n')
    if user_input.lower() == 'n':
        return False
    elif user_input.lower() == 'y':
        return True
    else:
        print('\nINVALID INPUT\n\n')
        return enter_transactions()


def calculate(sortword: str, income_or_expenses: int, print_transactions: bool):
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
            amount = excel_sheet.cell_value(row, income_or_expenses)
            if sortword.lower() not in name.lower() or amount == '':
                continue
            if print_transactions is True:
                date = datetime(*xlrd.xldate_as_tuple(date, excel_file.datemode)).strftime('%B %d %Y')
                print(f'{date} | {name} | {amount}')

            total += amount

    print(f'\n\nTOTAL: {round(total, 2)}')


main()  # Runs the script
