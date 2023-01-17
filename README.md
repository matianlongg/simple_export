# simple_export

simple_export是一款导出工具包，目标是根据模板快速导出，基于openpyxl

## 一、安装

```
pip install simple-export
```

## 二、使用

> 使用默认模板快速导出

执行以下代码会在当前文件夹生成一个val1.xlsx

```
#!/usr/bin/env python
# -*- coding: utf-8 -*-
from simple_export.example import test1
test1()
```

> example

    方法名：write_excel_for_template

    入参：value  # {"sheet页名称": {}} 一级key需要跟sheet页相同

    入参：wb_tmp # openpyxl的workbook对象

    入参：convert_pic # 是否需要转换value中的图片数据

    出参：None

```
#!/usr/bin/env python
# -*- coding: utf-8 -*-
import traceback
from pathlib import Path
from openpyxl import load_workbook
from excel import write_excel_for_template

def test1():
    wb_tmp = load_workbook(Path(__file__).parent / 'template/excel1.xlsx')
    value = {
        "家庭财产库存清单": {
            "thing": [
                {
                    "id": "=ROW($A1)",
                    "room": "客厅",
                    "goods": "物品1",
                    "structure": "制造商1",
                    "serial_number": "33XCBH3",
                    "time": "=TODAY()-120",
                    "source": "联机",
                    "price": 2000,
                    "eval_price": 2000,
                    "remark": "",
                    "has_photo": "是",
                    "photo_url": "https://t13.baidu.com/it/u=3512963775,847648221&fm=224&app=112&f=JPEG?w=500&h=500"
                },
                {
                    "id": "=ROW($A1)",
                    "room": "客厅2",
                    "goods": "物品2",
                    "structure": "制造商2",
                    "serial_number": "33XCBH4",
                    "time": "=TODAY()-90",
                    "source": "联机",
                    "price": 1000,
                    "eval_price": 1000,
                    "remark": "",
                    "has_photo": "是",
                    "photo_url": r"C:\Users\e9\Pictures\微信图片_20220616151258.png"
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
    write_excel_for_template(value=value, wb_tmp=wb_tmp, convert_pic=True)
    wb_tmp.save("./val1.xlsx")
    wb_tmp.close()
    # wb_tmp = load_workbook(f'./val1.xlsx', data_only=True)

if __name__ == '__main__':
    test1()
```

 在需要渲染数据得单元格使用${key} 替换key 多级需要用 _ 连接， 数组需要加入[*]

例：

> 数组 key_a[*]
> 
> 值 key_b

![](http://github.com/mtl940610/simple_export/blob/main/static/2023-01-17-17-07-05-image.png?raw=true)

处理后

![](https://github.com/mtl940610/simple_export/blob/main/static/2023-01-17-17-19-50-image.png?raw=true)
