#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查CRA等级表结构
"""

import openpyxl

def check_cra_file(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)

    print(f"文件: {filepath}")
    print(f"Sheets: {wb.sheetnames}\n")

    # 遍历所有sheet
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(f"\n=== Sheet: {sheet_name} ===")
        print(f"尺寸: {ws.max_row} x {ws.max_column}")

        # 显示前50行数据
        for row in range(1, min(50, ws.max_row + 1)):
            row_data = []
            for col in range(1, min(10, ws.max_column + 1)):
                val = ws.cell(row=row, column=col).value
                if val is not None:
                    row_data.append(f"C{col}={val}")
            if row_data:
                print(f"  R{row}: {', '.join(row_data)}")

    wb.close()

check_cra_file(r"F:\AI专用\AI工作区\底稿roll forward\输入；参考资料.xlsx")
