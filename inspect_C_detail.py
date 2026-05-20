#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细检查C货币资金输出文件的期初数和表2
"""

import openpyxl

def inspect_C_detail(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=False)
    sheet_name = "C.00 Lead"
    ws = wb[sheet_name]

    print(f"检查文件: {filepath}\n")

    # 检查第37-50行的数据（主表）
    print("=== 主表数据（第37-50行）===")
    for row in range(37, min(55, ws.max_row + 1)):
        for col in range(9, 15):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                print(f"  R{row}C{col}: {val}")

    # 检查第60行以下的For Disclosure表
    print("\n=== For Disclosure表（第60-75行）===")
    for row in range(60, min(76, ws.max_row + 1)):
        has_data = False
        row_data = []
        for col in range(1, 12):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                has_data = True
                row_data.append(f"C{col}={val}")
        if has_data:
            print(f"  R{row}: {', '.join(row_data)}")

    wb.close()

if __name__ == "__main__":
    inspect_C_detail(r"F:\AI专用\AI工作区\底稿roll forward\output_final\C SWP 货币资金 20261231 东风鸿远工程咨询有限公司.xlsx")
