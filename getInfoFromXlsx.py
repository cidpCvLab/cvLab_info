# -*- coding: utf-8 -*-
"""
Created on Sat Jul 10 09:01:06 2021

@author: 11200
"""

import xlrd  # ==1.2.0
from collections import defaultdict
import pprint


def readdk(file):
    data = xlrd.open_workbook(file)
    table = data.sheet_by_index(0)
    info = defaultdict(dict)

    for i in range(1, table.nrows):
        row = table.row_values(i)
        info[str(int(row[1]))] = {
            'name': row[0],
            'phone': str(row[2]),
            'e-mail': row[3]
        }

    return info


def read_number(file):
    data = xlrd.open_workbook(file)
    table = data.sheet_by_index(0)
    # numbers = table.col_values(1)[1:]
    # numbers = map(int, numbers)
    # numbers = map(str, numbers)
    # return list(numbers)
    return [str(int(i)) for i in table.col_values(1)[1:]]


path = "打卡信息录入 (1).xlsx"
cvLab = readdk(path)
pprint.pprint(cvLab)

numbers = read_number(path)
print(numbers)
