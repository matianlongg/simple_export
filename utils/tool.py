#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2022/12/9 22:56
# @Author  : mtl
# @File    : tool.py
# @Description : *****
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
    try:
        float(s)
        return True
    except:
        return False
def char_to_num(char: str) -> tuple:
    for i, s in enumerate(char):
        if is_number(s):
            return titleToNumber(char[:i]), int(char[i:])
    return (0, 0)

def pos_char_to_num(char: str) -> List[tuple]:
    res: list = []
    for c in char.split(":"):
        res.append(char_to_num(c))
    return res

def titleToNumber(columnTitle: str) -> int:
    ans = 0
    for s in columnTitle:
        num = ord(s) - ord('A') + 1
        ans = ans * 26 + num
    return ans

def convertToTitle(columnNumber: int) -> str:
    ans = list()
    while columnNumber > 0:
        columnNumber -= 1
        ans.append(chr(columnNumber % 26 + ord("A")))
        columnNumber //= 26
    return "".join(ans[::-1])

def num_to_pos_char(t: tuple) -> str:
    return convertToTitle(t[0]) + str(t[1])

if __name__ == '__main__':
    print(ord("c"), ord("C") - 64, ord("A") - 64)
    print(char_to_num("AA77"))

