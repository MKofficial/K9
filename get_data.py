import openpyxl as opx
from openpyxl.utils.exceptions import InvalidFileException
from my_utils import problem_message, set_cols_num, raise_error, dist_and_cat

unfinished = ['DNF', 'DSQ', 'DNS']


def setup_columns(file_txt: str = 'setup_cols.txt') -> list:
    """
    :param file_txt: File which will contain specific index for each column
    :return: List with number of each column
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


def add_points(people: list = [], multiply: bool = False) -> list:
    """
    :param people: List with people to be added points as their last element
    :param multiply: This multiply the points. Usage only if the K9 is on the sixth race
    :return: List with people containing points now
    """
    points = 11
    if multiply:
        points *= 1.5

    for person in people:
        if person[2] in unfinished:
            for i in unfinished:
                if person[2] == i:
                    person.append(i)
                    del person[2]
        if points == 10:
            points -= 1
        if points <= 0:
            person.append(0)
            del person[2]
        else:
            person.append(points)
            points -= 1
            del person[2]

    return people


def get_file_arr(excel_file: str = 'kuneticka_devitka.xlsx', distance: str = 'Dlouhý okruh - 9 km',
                 category: str = 'Muži 18 - 29 let', sixth_race: bool = False) -> list:
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

    people = [[ws[person.row][index[0]].value, ws[person.row][index[1]].value,
               ws[person.row][index[2]].value, ws[person.row][index[3]].value,
               ws[person.row][index[4]].value]
              for i in ws.iter_rows(min_col=index[3] + 1, max_col=index[3] + 1) for person in i
              if person.value == distance and ws[person.row][index[4]].value == category]

    # add points to participates in each category
    add_points(people, sixth_race)
    return people


def get_complete_file_arr(excel_file: str = 'kuneticka_devitka.xlsx') -> list:
    """
    :param excel_file: File with the race results
    :return: All participants of that race
    """
    people = [cat_people for param in dist_and_cat for cat_people in get_file_arr(excel_file, param[0], param[1])]
    return people


def get_final_arr(excel_file: str = 'final_table.xlsx') -> list:
    """
    :param excel_file: Excel file with the complete results
    :return: List of all elements in there
    """
    try:
        wb = opx.load_workbook(excel_file)
        ws = wb.active
    except InvalidFileException:
        raise_error(excel_file, problem_message[2])
    except FileNotFoundError:
        raise_error(excel_file, problem_message[0])
    else:
        if len(ws.column_dimensions) == 0:
            people = []
        else:
            people = [[person.value for person in element] for element in ws]

    return people