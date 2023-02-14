# -*- coding: utf-8 -*-
# !/usr/bin/env python
# @Time    : 2023/2/13 16:28 
# @Author  : mtl
# @Desc    : excel 增加自定义方法执行
# @File    : util.py
# @Software: PyCharm
import collections

from openpyxl.worksheet.worksheet import Worksheet
from utils.tool import char_to_num

class dynamic_method():

    def __call__(self, *args, **kwargs):
        funcs = args[0]
        for func in funcs:
            if func and func != "":
                call_function = getattr(self, func)
                call_function(*args[1:])

    """
    合并单元格向下合并
    """
    def merge_bottom(self, coordinate: str, sheet: Worksheet, pos_mapping: collections.defaultdict):
        col, row = char_to_num(coordinate)
        col, row = col - 1, row - 1
        pos = pos_mapping[row][col]
        start_row, end_row = pos[0] + 1, pos[-1] + 1
        sheet.merge_cells(start_row=start_row, start_column=col + 1,
                        end_column=col + 1, end_row=end_row)