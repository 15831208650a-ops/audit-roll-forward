#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细检查U_exp财务费用子表的期初数位置
"""

import openpyxl

def inspect_expense_rows(filepath, sheet_name):
    """检查财务费用子表的详细行结构"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb[sheet_name]

    print(f"Sheet: {sheet_name}\n")

    # 读取R28-R35的关键区域
    for row in range(27, min(40, ws.max_row + 1)):
        print(f"\nRow {row}:")
        for col in range(1, min(20, ws.max_column + 1)):
            val = ws.cell(row=row, column=col).value
            if val and str(val).strip():
                text = str(val).strip()
                if len(text) > 50:
                    text = text[:50] + "..."
                print(f"  C{col}: {text}")

    wb.close()


if __name__ == "__main__":
    filepath = r"F:\AI专用\AI工作区\底稿roll forward\templates\U_exp SWP other 202YMMDD XYZ公司.xlsx"
    inspect_expense_rows(filepath, "Uexp财务费用BKD")
