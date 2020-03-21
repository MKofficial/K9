version_control = 'v0.1.0'
version = f'K9 transform program.exe\nVersion: {version_control}\n'
problem_message = ['We could not find your file, because the file does not exist or the path is wrong',
                   'I do not understand your command', 'We could not open your file, because the file can\'t be '
                   'open as an excel sheet', 'System error']
set_cols_num = ['name', 'team', 'time', 'distance', 'category']


def raise_error(file: str = '?unknown file?', message: str = 'No description of the error') -> None:
    if file != '':
        print('Error!\n'.upper() + f'You try to open \"{file}\"\n' + f'Problem: {message}')
        input('\nPress any key to quit')
        quit()
    elif file == '':
        print('Error!\n'.upper() + f'Problem: {message}')
