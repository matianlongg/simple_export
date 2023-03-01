#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2022/12/9 22:56
# @Author  : mtl
# @File    : tool.py
# @Description : *****
import math
import re
import time
import re
from queue import Queue
from typing import List

def to_flat(x: dict) -> dict:
    q: Queue = Queue()
    q.put(('', x))
    result: dict = {}
    prefix: str
    o: dict
    while not q.empty():
        prefix, o = q.get()
        if isinstance(o, dict):
            for k, v in o.items():
                if prefix == "":
                    q.put((f"{k}", v))
                else:
                    q.put((f"{prefix}_{k}", v))
        elif isinstance(o, list):
            for i, v in enumerate(o):
                q.put((f"{prefix}[{i}]", v))
        else:
            result[prefix] = o
    return result

def is_number(s: str) -> bool:
    return bool(re.match(r'^[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?$', s))

def char_to_num(char: str) -> tuple:
    match = re.match(r'^([A-Za-z]+)(\d+)?$', char)
    if not match:
        return 0, 0
    letters, digits = match.groups()
    num = titleToNumber(letters)
    digit = int(digits) if digits else 0
    return num, digit

def pos_char_to_num(char: str) -> List[tuple]:
    return [char_to_num(c) for c in char.split(":")]

def titleToNumber(columnTitle: str) -> int:
    ans = 0
    for s in columnTitle:
        num = ord(s) - ord('A') + 1
        ans = ans * 26 + num
    return ans

def convertToTitle(columnNumber: int) -> str:
    OFFSET = ord('A') - 1  # A对应的数字
    title = ""
    columnNumber += 1
    while columnNumber > 0:
        columnNumber -= 1
        num = columnNumber % 26
        title = chr(num + OFFSET) + title
        columnNumber //= 26
    return title

def num_to_pos_char(t: tuple) -> str:
    col, row = t
    return f"{convertToTitle(col)}{row}"

def getNowTime(format="%Y-%m-%d %H:%M:%S"):
    """
        获取现在的时间
    """
    return time.strftime(format, time.localtime())


def pixels_to_points(value, dpi=96):
    """96 dpi, 72i"""
    return value * 72 / dpi

def points_to_pixels(value, dpi=96):
    return int(math.ceil(value * dpi / 72))

if __name__ == '__main__':
    print(ord("c"), ord("C") - 64, ord("A") - 64)
    print(char_to_num("AA77"))