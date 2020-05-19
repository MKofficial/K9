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
    final_table_ws = final_table_wb["Raw data"]
    final_table_ws_styled = final_table_wb["Styled"]
    total_results = final_table_wb["Total Results"]

    # compare and save to final_table.xlsx
    comp.compare_and_save(final_table_wb, 'final_table.xlsx', final_table_ws, comp.compare, race_data, complete_data,
                          int(race_file[0]))

    # get complete data from races before and the race now
    complete_data = gd.get_final_arr('final_table.xlsx')

    # styling final_table.xlsx
    # stl.style_and_save(final_table_wb, 'final_table.xlsx', final_table_ws, stl.style_final_table)

    # add points
    complete_data = gd.points_sum(complete_data)

    # save to styled final table
    points = [[i[-1]] for i in complete_data]
    for i in points:
        final_table_ws_styled.append(i)

    final_table_wb.save('final_table.xlsx')