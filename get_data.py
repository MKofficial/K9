import openpyxl as opx
from openpyxl.utils.exceptions import InvalidFileException


def get_file_arr(excel_file: str = None) -> list:
    """
    :param excel_file: Excel file with the newest race's results
    :return: List that contains persons from the race
    """
    try:
        # try to open new workbook
        wb = opx.load_workbook(excel_file)
        ws = wb.active
    except InvalidFileException:
        print('Error!\n'.upper() + f'You try to open \"{excel_file}\"\n\n' + 'There is an error in opening your file\n'
                                                                             'Please rerun the program again with the '
                                                                             'right file format')
        input('\nPress any key to continue')
        wb = opx.Workbook()

    except FileNotFoundError:
        print('Error!\n'.upper() + f'You try to open \"{excel_file}\"\n\n' + 'There is an error in finding your file\n'
                                                                             'Please rerun the program again with the '
                                                                             'right file name')
        input('\nPress any key to continue')
        quit()
    else:
        pass

    wb.save(excel_file)
