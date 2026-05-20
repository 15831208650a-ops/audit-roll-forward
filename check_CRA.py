#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查C货币资金模板中CRA的位置
"""

import openpyxl

def check_cra_location(template_path):
    wb = openpyxl.load_workbook(template_path, data_only=False)
    ws = wb["C.00 Lead"]

    print("查找CRA相关行：")
    for row in range(1, 30):
        for col in range(1, 10):
            val = ws.cell(row=row, column=col).value
            if val and any(x in str(val) for x in ["CRA", "风险等级", "层级", "RP", "风险系数"]):
                print(f"  R{row}C{col}: {val}")

    wb.close()

check_cra_location(r"F:\AI专用\AI工作区\底稿roll forward\templates\C SWP 货币资金 202YMMDD XYZ公司.xlsx")
