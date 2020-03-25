version_control = 'v0.1.0'
version = f'K9 transform program.exe\nVersion: {version_control}\n'
problem_message = ['We could not find your file, because the file does not exist or the path is wrong',
                   'I do not understand your command', 'We could not open your file, because the file can\'t be '
                   'open as an excel sheet', 'System error']
set_cols_num = ['name', 'team', 'time', 'distance', 'category']
dist_and_cat = [['Dlouhý okruh - 9 km', 'Muži 18 - 29 let'], ['Dlouhý okruh - 9 km', 'Muži 30 - 39 let'],
            ['Dlouhý okruh - 9 km', 'Muži 40 - 49 let'], ['Dlouhý okruh - 9 km', 'Muži 50 - 59 let'],
            ['Dlouhý okruh - 9 km', 'Muži 60 - 69 let'], ['Dlouhý okruh - 9 km', 'Muži 70 +'],
            ['Dlouhý okruh - 9 km', 'Ženy 18 - 29 let'], ['Dlouhý okruh - 9 km', 'Ženy B (30 - 39 let)'],
            ['Dlouhý okruh - 9 km', 'Ženy 40 - 49 let'], ['Dlouhý okruh - 9 km', 'Ženy 50 +'],
            ['Krátký okruh - 4,5 km', 'Žáci 6 - 10 let'], ['Krátký okruh - 4,5 km', 'Žáci 11 - 14 let'],
            ['Krátký okruh - 4,5 km', 'Junioři 15 - 17 let'], ['Krátký okruh - 4,5 km', 'Muži 18 +'],
            ['Krátký okruh - 4,5 km', 'Žákyně 6 - 10 let'], ['Krátký okruh - 4,5 km', 'Žákyně 11 - 14 let'],
            ['Krátký okruh - 4,5 km', 'Juniorky 15 - 17 let'], ['Krátký okruh - 4,5 km', 'Ženy 18 +']]


def raise_error(file: str = '?unknown file?', message: str = 'No description of the error') -> None:
    if file != '':
        print('Error!\n'.upper() + f'You try to open \"{file}\"\n' + f'Problem: {message}')
        input('\nPress any key to quit')
        quit()
    elif file == '':
        print('Error!\n'.upper() + f'Problem: {message}')
