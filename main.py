import get_data as gd
import openpyxl as opx
import styling as stl
from my_utils import *
import os

if __name__ == '__main__':
    print(version)
    race_path = input('Type an absolute path to your file: ')
    race_file = os.path.basename(race_path)

    print('\nWarning!!!'.upper() + '\nSave and close all your excel workbooks before running program\n')

    # ask user to run the program
    while True:
        run = input('Run? (Y/N): ').lower()
        if run in ['n', 'no', 'ne']:
            quit()
        elif run in ['y', 'yes', 'ano', 'a']:
            # run the program via breaking this loop
            break
        else:
            print(problem_message[1])
            input('Press any key to continue')
            print()

    final_arr = gd.get_final_arr('final_table.xlsx')
    if race_file[0] == '6':
        complete_file_arr = gd.get_complete_file_arr(race_path, True)
    else:
        complete_file_arr = gd.get_complete_file_arr(race_path, False)

    final_table_wb = opx.load_workbook('final_table.xlsx')
    final_table_ws = final_table_wb.active

    # TODO: compare and save some data to final_table

    # styling
    stl.style_and_save(final_table_wb, 'final_table.xlsx', stl.style_final_table, final_table_ws)
