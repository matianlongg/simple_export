#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2022/12/10 0:33
# @Author  : mtl
# @File    : example.py
# @Description : *****


from openpyxl import load_workbook

from excel import write_excel_for_template

wb_tmp = load_workbook(f'./111.xlsx')
value = {
    "家庭财产库存清单": {
        "room": [
            {
                "name": "客厅",
                "goods": "物品1"
            },
            {
                "name": "客厅",
                "goods": "物品1"
            }
        ],
        "test": "1"
    },
    "sheet1": {
        "data": {
            "name": "1",
            "data_list": [
                {"a": 3},
                {"a": 4}
            ],
            "b": "cec测试"
        },
        "data_list": [
            {
                "a": 1,
                "b": 2
            },
            {
                "a": 2,
                "b": 1
            }
        ],
        "name": 123
    },
    "sheet2": {
        "data": {
            "name": "1",
            "data_list": [
                {"a": 3}
            ],
            "b": "cec测试"
        }
    }
}
write_excel_for_template(value=value, wb_tmp=wb_tmp)
wb_tmp.save("./val.xlsx")
wb_tmp = load_workbook(f'./val.xlsx', data_only=True, read_only=True)