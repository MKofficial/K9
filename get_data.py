import openpyxl as opx
from openpyxl.utils.exceptions import InvalidFileException
from my_utils import problem_message, set_cols_num, raise_error


def setup_columns(file_txt: str = 'setup_cols.txt') -> list:
    """
    :param file_txt: File which will contain specific index for each column
    :return: List with
    """
    with open(file_txt, 'r', encoding='utf-8') as file:
        if file.read() == '':
            file.close()
            with open(file_txt, 'w', encoding='utf-8') as file_2:
                for param in set_cols_num:
                    file_2.write(f'{param}:' + input('Set column index ("A" -> 1) for '
                                                     + param.upper() + ': ') + '\n')
                file_2.close()
        elif file.read() is not None:
            rewrite = input('Do you want to open settings for column\'s parameters (Y/N): ').lower()
            file.close()
            if rewrite in ['y', 'yes', 'ano', 'a']:
                with open('setup_cols.txt', 'w', encoding='utf-8') as file_2:
                    for param in set_cols_num:
                        file_2.write(f'{param}:' + input('Set column index ("A" -> 0) for '
                                                         + param.upper() + ': ') + '\n')
                    file_2.close()
        else:
            raise_error(file='', message=problem_message[3])

    with open(file_txt, 'r', encoding='utf-8') as file:
        arr = [int((i.strip('\n')).split(':')[-1]) for i in file.readlines()]
        file.close()

    return arr


def get_file_arr(excel_file: str = 'kuneticka_devitka.xlsx', distance: str = 'Dlouhý okruh - 9 km',
                 category: str = 'Muži 18 - 29 let') -> list:
    """
    :param excel_file: Excel file with the newest race's results
    :param category: Person's category
    :param distance: Person's distance
    :return: List that contains persons from specific category and distance
    """
    try:
        # try to open new workbook
        wb = opx.load_workbook(excel_file)
        ws = wb.active
    except InvalidFileException:
        raise_error(excel_file, problem_message[2])
    except FileNotFoundError:
        raise_error(excel_file, problem_message[0])
    else:
        index = [i for i in setup_columns()]
        print(index)

    people = [[ws[person.row][index[0]].value, ws[person.row][index[1]].value,
               ws[person.row][index[2]].value, ws[person.row][index[3]].value,
               ws[person.row][index[4]].value]
              for i in ws.iter_rows(min_col=index[3] + 1, max_col=index[3] + 1) for person in i
              if person.value == distance and ws[person.row][index[4]].value == category]

    return people


people = get_file_arr()
print(people)
