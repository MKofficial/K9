from my_utils import *
from openpyxl.utils.exceptions import InvalidFileException
import openpyxl as opx

heading = ['Pořadí', 'Jméno a příjmení', 'Klub', 'Distance', 'Kategorie', '1. závod', '2. závod', '3. závod',
           '4. závod', '5. závod', '6. závod', '7. závod', '8. závod', '9. závod', '10. závod', '11. závod',
           '12. závod', 'Body']


def style_and_save(excel_workbook, excel_workbook_name, excel_workbook_sheet, styling_function):
    """
    :param excel_workbook: Workbook from openpyxl library
    :param excel_workbook_name: Name of the the certain workbook
    :param styling_function: Function that will style the table
    :param excel_workbook_sheet: Workbook sheet from openpyxl library
    :return: None

    This function apply the certain styling function to the current table and save it
    """

    styling_function(excel_workbook_sheet)
    excel_workbook.save(excel_workbook_name)


# there has to be defined workbook and sheet which is passed as an argument to excel_file parameter
# WARNING you have to save the sheet manually in the code
def preposition(excel_sheet) -> None:
    """
    :param excel_sheet: Excel sheet to be styled
    :return: None

    These change will be visible after saving that workbook with openpyxl
    WARNING this function does not support saving
    """
    try:
        excel_sheet.insert_rows(0)
        excel_sheet.insert_cols(0)
        excel_sheet.append(heading)
        excel_sheet.move_range('A' + str(excel_sheet.max_row) + ':R' + str(excel_sheet.max_row),
                               rows=-excel_sheet.max_row + 1)
    except FileNotFoundError:
        raise_error(excel_sheet, problem_message[0])
    except InvalidFileException:
        raise_error(excel_sheet, problem_message[2])


def category_and_position(workbook_name, array: list, race_number: int) -> None:
    """
    :param workbook_name: Excel workbook name
    :param array: Array containing people and their points
    :param race_number: An integer defining the race number to make "Body" in right position
    :return: None
    """
    workbook = opx.Workbook()
    sheet = workbook.active

    headline = ['Pořadí', 'Jméno a příjmení', 'Klub', '1. závod', '2. závod', '3. závod', '4. závod',
                '5. závod', '6. závod', '7. závod', '8. závod', '9. závod', '10. závod', '11. závod', '12. závod',
                'Body']
    for param in dist_and_cat:
        position = 0
        temp = []

        for person in array:
            if person[2] == param[0] and person[3] == param[1]:
                # warning...deleting position will change position of other values!!!
                # deleting category and route
                del person[3]
                del person[2]

                # append the person to temporary array
                temp.append(person)

        # sort persons in temporary array with the last value of each field - sum of points
        temp.sort(key=lambda x: x[-1], reverse=True)

        # to check if there are some people having same sum of points
        last_pos = None
        for i in temp:
            if i[-1] == 0:
                i.insert(0, "-")
            elif last_pos != i[-1]:
                position += 1
                last_pos = i[-1]
                i.insert(0, position)
            else:
                i.insert(0, position)
            for j in range(12 - race_number):
                i.insert(-1, None)

        temp.insert(0, headline)
        temp.insert(0, (param[0], param[1]))

        for i in temp:
            sheet.append(i)

        # styling headline of each category
        for i in sheet['A']:
            ...

        workbook.save(workbook_name)
