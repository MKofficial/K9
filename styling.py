from my_utils import *
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.styles import PatternFill, Alignment
from openpyxl.styles.fonts import Font
from openpyxl.styles.borders import Border, Side
from openpyxl.drawing.image import Image
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
            # insert None to manage the sum of points into the right column
            for j in range(12 - race_number):
                i.insert(-1, None)

        temp.insert(0, headline)
        temp.insert(0, (param[0], param[1]))

        for i in temp:
            sheet.append(i)

        # styling headline of each category
        for i in sheet['A']:
            if i.value == 'Dlouhý okruh - 9 km' or i.value == 'Krátký okruh - 4,5 km':
                for j in sheet[i.row]:
                    j.fill = PatternFill(start_color='c0c0c0', end_color='c0c0c0', fill_type='solid')
                    if j == "<Cell 'Sheet'.P" + str(sheet[i.row]) + ">":
                        break

        for i in sheet['B']:
            for c in dist_and_cat:
                if i.value == c[1]:
                    sheet.cell(row=i.row, column=2, value=dist_and_cat_modified[dist_and_cat.index(c)][1])

            if i.value == 'Mladší žáci (6 - 10 let)':
                row = i.row + 1
                sheet.insert_rows(i.row)
                sheet.cell(row=i.row - 1, column=1, value='Malý okruh - 4 500 m')
                sheet.insert_rows(i.row - 1)
                # ws.merge_cells(f'A{row + 11}:C{row + 11}')
                # color
                for j in sheet[row]:
                    j.fill = PatternFill(start_color='ccffcc', end_color='ccffcc', fill_type='solid')
                    j.font = Font(name='Arial', size=15, color='FF0000', bold=True)
                    if j == "<Cell 'Sheet'.P" + str(sheet[row]) + ">":
                        break

            if i.value == 'Muži A (18 - 29 let)':
                sheet.insert_rows(i.row)
                sheet.cell(row=i.row - 1, column=1, value='Hlavní závod - 8 800 m')
                # ws.merge_cells(f'A{i.row - 1}:C{i.row - 1}')
                # color
                for j in sheet.iter_rows(min_row=i.row - 1, max_row=i.row - 1):
                    for k in j:
                        k.fill = PatternFill(start_color='ccffcc', end_color='ccffcc', fill_type='solid')
                        k.font = Font(name='Arial', size=15, color='FF0000', bold=True)
                        if k == "<Cell 'Sheet'.P" + str(sheet[i.row]) + ">":
                            break

        workbook.save(workbook_name)
