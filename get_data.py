import openpyxl as opx
from openpyxl.utils.exceptions import InvalidFileException
from my_utils import problem_message, set_cols_num


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
        quit()
    except FileNotFoundError:
        print('Error!\n'.upper() + f'You try to open \"{excel_file}\"\n\n' + 'There is an error in finding your file\n'
                                                                             'Please rerun the program again with the '
                                                                             'right file name or path')
        input('\nPress any key to continue')
        quit()
    else:
        with open('setup_cols.txt', 'r', encoding='utf-8') as file:
            if file.read() == '':
                file.close()
                with open('setup_cols.txt', 'w', encoding='utf-8') as file_2:
                    for param in set_cols_num:
                        file_2.write(f'{param}:' + input('Set column index ("A" -> 1) for '
                                                         + param.upper() + ': ') + '\n')
            elif file.read() is not None:
                rewrite = input('Do you want to open settings for column\'s parameters (Y/N): ').lower()
                file.close()
                if rewrite in ['y', 'yes', 'ano', 'a']:
                    with open('setup_cols.txt', 'w', encoding='utf-8') as file_2:
                        for param in set_cols_num:
                            file_2.write(f'{param}:' + input('Set column index ("A" -> 1) for '
                                                             + param.upper() + ': ') + '\n')
            else:
                print('System error!\n'.upper())
                input('Press any key to quit')
                quit()