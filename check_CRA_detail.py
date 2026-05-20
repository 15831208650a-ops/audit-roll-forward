#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细检查CRA区域
"""

import openpyxl

def check_cra_detail(template_path):
    wb = openpyxl.load_workbook(template_path, data_only=False)
    ws = wb["C.00 Lead"]

    print("=== CRA区域 (R12-R18) ===")
    for row in range(12, 20):
        row_data = []
        for col in range(1, 8):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                row_data.append(f"C{col}={val}")
        if row_data:
            print(f"  R{row}: {', '.join(row_data)}")

    wb.close()

check_cra_detail(r"F:\AI专用\AI工作区\底稿roll forward\templates\C SWP 货币资金 202YMMDD XYZ公司.xlsx")
