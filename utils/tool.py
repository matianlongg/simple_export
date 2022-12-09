#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2022/12/9 22:56
# @Author  : mtl
# @File    : tool.py
# @Description : *****
from queue import Queue

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