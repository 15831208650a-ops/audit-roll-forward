#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细检查模板和上年底稿的结构差异
"""

import openpyxl

def check_file(filepath, name, data_only=True):
    wb = openpyxl.load_workbook(filepath, data_only=data_only)
    ws = wb["C.00 Lead"]

    print(f"\n=== {name} ===")

    # 1. 表头区域（R37附近）
    print("\n--- 表头区域(R37) ---")
    for col in range(1, 16):
        val = ws.cell(row=37, column=col).value
        if val:
            print(f"  C{col}: {val}")

    # 2. 数据行（R38-R45）
    print("\n--- 数据行 ---")
    for row in range(38, 50):
        row_data = []
        for col in range(1, 16):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                row_data.append(f"C{col}={val}")
        if row_data:
            print(f"  R{row}: {', '.join(row_data)}")

    # 3. 合计行
    print("\n--- 合计行 ---")
    for row in [42, 43, 44, 45, 46, 47]:
        row_data = []
        for col in range(1, 16):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                row_data.append(f"C{col}={val}")
        if row_data:
            print(f"  R{row}: {', '.join(row_data)}")

    # 4. 日期
    print("\n--- 日期 ---")
    for row in range(1, 10):
        for col in range(1, 5):
            val = ws.cell(row=row, column=col).value
            if val and "202" in str(val):
                print(f"  R{row}C{col}: {val}")

    wb.close()

check_file(r"F:\AI专用\AI工作区\底稿roll forward\templates\C SWP 货币资金 202YMMDD XYZ公司.xlsx", "模板", data_only=False)
check_file(r"F:\AI专用\AI工作区\底稿roll forward\上年底稿文件夹\C SWP 货币资金 20251231 东风鸿远工程咨询有限公司.xlsx", "上年底稿", data_only=True)
