from my_utils import *


def compare(race_data_set: list = None, complete_data_set: list = None, race_num: int = 1) -> list:
    for i in complete_data_set:
        if i[0] not in [i[0] for i in race_data_set]:
            i.append('-')
        for j in race_data_set:
            if i[0] == j[0]:
                i.append(j[4])
    for i in race_data_set:
        if i[0] not in [x[0] for x in complete_data_set]:
            ...

    return complete_data_set
