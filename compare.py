
from my_utils import *


def compare_and_save(excel_workbook, excel_workbook_name, excel_workbook_sheet, comparing_function,
                     param1: list, param2: list, param3: int) -> None:
    """
    :param param3: Race_num
    :param param2: Complete_data_set
    :param param1: Race_data_set
    :param excel_workbook: Workbook from openpyxl library
    :param excel_workbook_name: Name of the the certain workbook
    :param excel_workbook_sheet: Workbook sheet form openpyxl library
    :param comparing_function: Function that will style the table
    :return: None

    This function apply the certain styling function to the current table and save it
    """

    complete_data = comparing_function(param1, param2, param3)
    for i in complete_data:
        excel_workbook_sheet.append(i)

    excel_workbook.save(excel_workbook_name)


def compare(race_data_set: list = None, complete_data_set: list = None, race_num: int = 1) -> list:
    """
    :param race_data_set: Data from the race
    :param complete_data_set: Data collected through all races
    :param race_num: Number defining race number
    :return: List containing
    """
    for i in complete_data_set:
        # in complete_data_set, but not in the race_data_set
        if i[0] not in [i[0] for i in race_data_set]:
            i.append('-')
        # in complete_data_set and in the race_data_set
        for j in race_data_set:
            if i[0] == j[0]:
                i.append(j[4])
    # in the race_data_set, but not in the complete_data_set
    for i in race_data_set:
        if i[0] not in [x[0] for x in complete_data_set]:
            if race_num == 1:
                complete_data_set.append(i)
            else:
                for j in range(race_num - 1):
                    i.insert(4, '-')
                    complete_data_set.append(i)

    return complete_data_set
