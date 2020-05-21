import get_data as gd
import openpyxl as opx
import styling as stl
import compare as comp
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

    complete_data = gd.get_final_arr('final_table.xlsx')  # final arr
    if race_file[0] == '6':
        race_data = gd.get_complete_file_arr(race_path, True)  # complete_file_arr
    else:
        race_data = gd.get_complete_file_arr(race_path, False)

    final_table_wb = opx.load_workbook('final_table.xlsx')
    final_table_ws = final_table_wb.active

    # compare and save to final_table.xlsx
    comp.compare_and_save('final_table.xlsx',comp.compare, race_data, complete_data,
                          int(race_file[0]))

    # get complete data from races before and the race now
    complete_data = gd.get_final_arr('final_table.xlsx')

    # add points
    complete_data = gd.points_sum(complete_data)

    # save first styling to total results
    # stl.style_and_save(final_table_wb, "final_table.xlsx", total_results, stl.style_final_table)

    # apply category headline to total results
    stl.category_and_position("Total results.xlsx", complete_data, int(race_file[0]))
