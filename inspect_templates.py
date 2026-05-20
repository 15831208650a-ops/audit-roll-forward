#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量检查所有底稿模板的结构差异
"""

import os
import openpyxl

def inspect_excel(filepath):
    """检查单个Excel文件的结构"""
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheets = wb.sheetnames
        info = {
            "filename": os.path.basename(filepath),
            "sheets": sheets,
            "details": {}
        }

        for sheet_name in sheets:
            ws = wb[sheet_name]
            info["details"][sheet_name] = {
                "max_row": ws.max_row,
                "max_col": ws.max_column
            }

            # 检查前20行是否有"期末审定数"
            if "Lead" in sheet_name or "K.00" in sheet_name:
                for row in range(1, min(80, ws.max_row + 1)):
                    for col in range(1, 15):
                        val = ws.cell(row=row, column=col).value
                        if val and "期末审定数" in str(val):
                            info["details"][sheet_name]["期末审定数_row"] = row
                            info["details"][sheet_name]["期末审定数_col"] = col
                            break
                    if "期末审定数_row" in info["details"][sheet_name]:
                        break

            # 检查是否有"年初余额"
            if "K.01" in sheet_name or "Agree" in sheet_name:
                for row in range(1, min(80, ws.max_row + 1)):
                    for col in range(1, 5):
                        val = ws.cell(row=row, column=col).value
                        if val and "年初余额" in str(val):
                            info["details"][sheet_name]["年初余额_row"] = row
                            info["details"][sheet_name]["年初余额_col"] = col
                            break
                    if "年初余额_row" in info["details"][sheet_name]:
                        break

                # 检查第10行左右的列头
                header_data = {}
                for row in [10, 11, 12]:
                    row_data = []
                    for col in range(1, min(25, ws.max_column + 1)):
                        val = ws.cell(row=row, column=col).value
                        if val:
                            row_data.append(f"Col{col}={val}")
                    if row_data:
                        header_data[f"Row{row}"] = row_data
                info["details"][sheet_name]["headers"] = header_data

        wb.close()
        return info
    except Exception as e:
        return {"filename": os.path.basename(filepath), "error": str(e)}


def main():
    # 检查根目录下的所有Excel文件
    base_dir = r"F:\AI专用\AI工作区\底稿roll forward"

    files = []
    for f in os.listdir(base_dir):
        if f.endswith('.xlsx') and not f.startswith('~') and not f.startswith('输入'):
            files.append(os.path.join(base_dir, f))

    print(f"找到 {len(files)} 个文件:\n")

    for filepath in sorted(files):
        info = inspect_excel(filepath)
        print(f"{'='*80}")
        print(f"文件: {info['filename']}")
        if 'error' in info:
            print(f"  错误: {info['error']}")
            continue

        print(f"  Sheet数量: {len(info['sheets'])}")
        print(f"  Sheet列表: {info['sheets']}")

        for sheet_name, details in info['details'].items():
            print(f"\n  [{sheet_name}] 行x列: {details['max_row']}x{details['max_col']}")
            if "期末审定数_row" in details:
                print(f"    '期末审定数' 位于: 第{details['期末审定数_row']}行, 第{details['期末审定数_col']}列")
            if "年初余额_row" in details:
                print(f"    '年初余额' 位于: 第{details['年初余额_row']}行, 第{details['年初余额_col']}列")
            if "headers" in details:
                for row_key, row_vals in details["headers"].items():
                    print(f"    {row_key}: {', '.join(row_vals[:5])}...")

        print()

if __name__ == "__main__":
    main()
