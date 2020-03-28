from my_utils import *
from openpyxl.utils.exceptions import InvalidFileException
heading = ['Pořadí', 'Jméno a příjmení', 'Klub', 'Distance', 'Kategorie', '1. závod', '2. závod', '3. závod',
           '4. závod', '5. závod', '6. závod', '7. závod', '8. závod', '9. závod', '10. závod', '11. závod',
           '12. závod', 'Body']


def style_and_save(excel_workbook, excel_workbook_name, styling_function):

    styling_function
    excel_workbook.save(excel_workbook_name)


# there has to be defined workbook and sheet which is passed as an argument to excel_file parameter
# WARNING you have to save the sheet manually in the code
def style_final_table(excel_sheet) -> None:
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
