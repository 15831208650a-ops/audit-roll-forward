#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证C货币资金输出文件
"""

import openpyxl

output_path = r"F:\AI专用\AI工作区\底稿roll forward\output_test\C SWP 货币资金 20261231 东风鸿远工程咨询有限公司.xlsx"
wb = openpyxl.load_workbook(output_path, data_only=False)
ws = wb["C.00 Lead"]

print("=== 日期 ===")
print(f"R3C3: {ws.cell(row=3, column=3).value}")

print("\n=== 主表数据（R37-R45）===")
for row in range(37, 50):
    row_data = []
    for col in range(1, 16):
        val = ws.cell(row=row, column=col).value
        if val is not None:
            row_data.append(f"C{col}={val}")
    if row_data:
        print(f"  R{row}: {', '.join(row_data)}")

print("\n=== For Disclosure（R60-R75）===")
for row in range(60, 76):
    row_data = []
    for col in range(1, 8):
        val = ws.cell(row=row, column=col).value
        if val is not None:
            row_data.append(f"C{col}={val}")
    if row_data:
        print(f"  R{row}: {', '.join(row_data)}")

wb.close()
