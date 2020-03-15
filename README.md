# Kuneticka devitka
This is a code for modifing and styling results with **Python** usign *openpyxl* library.

---
## Main.py file
This is the main program which import other modules.

---
## Get_data.py file
This module handle functions used in main.py program.
### Function: get_file_arr(excel_file)
    :param excel_file: Excel file with the newest race's results
    :return: List that contains persons from the race

Function open new workbook using library *openpyxl*.
Handle some exceptions if the *excel_file* is not find or if there is another problem.
Then it opens a file to fetch a data from. These data contain information about columns searchnig algorithm. 

#### setup_cols.txt
> name: 1
> team: 4
time: 11
distance: 5
category: 6