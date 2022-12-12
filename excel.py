# -*- coding: utf-8 -*-
# !/usr/bin/env python
# @Time    : 2022/11/9 15:03 
# @Author  : mtl
# @Desc    : excel 根据 模板导出工具
# @File    : index.py
# @Software: PyCharm
import collections
import copy
import traceback
from openpyxl.workbook import Workbook
from openpyxl.worksheet import worksheet
from openpyxl.cell import Cell, MergedCell
import re

from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet
from utils.tool import to_flat, char_to_num, pos_char_to_num, num_to_pos_char


class work_sheet_tool():

    def copy_cell(self, source_cell: Cell, target_cell: Cell) -> None:
        target_cell._style = copy.copy(source_cell._style)
        target_cell.font = copy.copy(source_cell.font)
        target_cell.border = copy.copy(source_cell.border)
        target_cell.fill = copy.copy(source_cell.fill)
        target_cell.number_format = copy.copy(source_cell.number_format)
        target_cell.protection = copy.copy(source_cell.protection)
        target_cell.alignment = copy.copy(source_cell.alignment)
        target_cell.value = source_cell.value

    def write_sheet(self, source: Worksheet, obj_value: dict, target: Worksheet):
        pos: list = [list(row) for row in source.iter_rows()]
        pos_mapping: collections.defaultdict = collections.defaultdict(list)
        rlen: int = len(pos)
        clen: int = len(pos[0])
        i: int = 0
        w: int = 0
        start_row: int = 0
        end_row: int = 0
        merge_dict: dict = {}
        while i < rlen:
            j: int = 0
            maxn: int = 0
            if i - w not in pos_mapping.keys():
                pos_mapping[i - w].append(i)
            while j < clen:
                try:
                    if pos[i][j] == 0:
                        pos[i][j] = copy.copy(pos[i - 1][j])
                        if f"{start_row}-{j}" not in merge_dict and pos[start_row][j].value is not None and \
                                pos[start_row][j].value != "":
                            merge_dict[f"{start_row}-{j}"] = [start_row + 1, j + 1, end_row + 1, j + 1]
                    if isinstance(pos[i][j].value, str):
                        col: Cell = pos[i][j]
                        match_value: list = re.findall('\${(.+)}', col.value)
                        if len(match_value) > 0:
                            match_value_str: str = match_value[0]
                            index: int = match_value_str.find("*")
                            n: int = 0
                            if index > 0:
                                while True:
                                    key: str = f"{match_value_str[:index]}{n}{match_value_str[index + 1:]}"
                                    if key in obj_value.keys():
                                        if n == 0:
                                            col.value = obj_value[key]
                                        elif n + i > rlen or pos[n + i][j] != 0:
                                            ls_col = [0] * clen
                                            ccol = copy.copy(col)
                                            ccol.value = obj_value[key]
                                            ls_col[j] = ccol
                                            pos.insert(n + i, ls_col)
                                            if n + i not in pos_mapping.keys() and n + i not in pos_mapping[i - w]:
                                                pos_mapping[i - w].append(n + i)
                                        else:
                                            ccol = copy.copy(col)
                                            ccol.value = obj_value[key]
                                            pos[n + i][j] = ccol
                                        n += 1
                                    else:
                                        break
                                maxn = max(n - 1, maxn)
                            else:
                                col.value = obj_value.get(match_value_str, "")
                except:
                    traceback.print_exc()
                j += 1
            if maxn > 0:
                rlen += maxn
                w += maxn
                start_row, end_row = i, i + maxn
            i += 1
        target.insert_rows(len(pos))
        for i, r in enumerate(pos):
            for j, c in enumerate(r):
                source_cell: [Cell, MergedCell] = pos[i][j]
                target_cell: Cell = target.cell(i + 1, j + 1)
                self.copy_cell(source_cell, target_cell)
        for cell in source.merged_cells:
            cell_min_num = pos_mapping[cell.min_row - 1][0] + 1
            cell_max_num = pos_mapping[cell.max_row - 1][0] + 1
            target.merge_cells(start_row=cell_min_num, start_column=cell.min_col,
                            end_column=cell.max_col, end_row=cell_max_num)
        for key in merge_dict:
            value = merge_dict[key]
            target.merge_cells(start_row=value[0], start_column=value[1],
                            end_column=value[3], end_row=value[2])
        target.column_dimensions = copy.deepcopy(source.column_dimensions)
        for index in source.row_dimensions:
            for pos in pos_mapping.get(index, []):
                target.row_dimensions[pos].height = source.row_dimensions[index].height
        self.write_table(source, target, pos_mapping)

    def write_table(self, source: Worksheet, target: Worksheet, pos_mapping: dict):
        print(source.tables)
        tab: tuple
        for tab in source.tables.items():
            ...
            left_top, right_bottom = pos_char_to_num(tab[1])
            left_top = (left_top[0], pos_mapping[left_top[1]][0])
            # bug后续修改
            right_bottom = (right_bottom[0], pos_mapping.get(right_bottom[1], [right_bottom[1] + 1])[0])
            ctab = copy.deepcopy(source.tables[tab[0]])
            ctab.ref = f"{num_to_pos_char(left_top)}:{num_to_pos_char(right_bottom)}"
            ctab.autoFilter.ref = ctab.ref
            target.add_table(ctab)

def write_excel_for_template(value: dict, wb_tmp: Workbook) -> None:
    sheet_name: str
    wst: work_sheet_tool = work_sheet_tool()
    for sheet_name, obj in value.items():
        obj_value: dict = to_flat(obj)
        if sheet_name in wb_tmp.sheetnames:
            source: worksheet = wb_tmp[sheet_name]
            target: Worksheet = wb_tmp.create_sheet("new_" + sheet_name)
            wb_tmp.remove(source)
            wst.write_sheet(source, obj_value, target)
            target.views.sheetView = source.views.sheetView
            target.title = sheet_name
