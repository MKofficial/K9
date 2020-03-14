import get_data as gd
from my_utils import *
import os

if __name__ == '__main__':
    print(version)
    g_race_path = input('Type an absolute path to your file: ')

    # search for slashes or backslashes in the path
    # if there is not any, raise error
    if g_race_path.find('\\') != -1:
        g_race_file = g_race_path.split('\\')[-1]
    elif g_race_path.find('/') != -1:
        g_race_file = g_race_path.split('/')[-1]
    else:
        print('error\n'.upper() + f'You try to open \"{g_race_path}\"\n\n' + f'Problem: {problem_message[0]}')
        input('Press any key to quit')
        quit()

    print('Save and close all your excel workbooks before running program\n')

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

    gd.get_file_arr(g_race_path)