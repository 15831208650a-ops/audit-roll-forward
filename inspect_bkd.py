#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查BKD格式科目的Sheet结构
"""

import os
import openpyxl

def inspect_bkd_sheets(filepath):
    """检查BKD表的详细结构"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    result = {"filename": os.path.basename(filepath), "sheets": {}}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # 只检查包含BKD或Agree的Sheet
        if any(x in sheet_name for x in ["BKD", "Agree", "01", "明细"]):
            sheet_info = {"max_row": ws.max_row, "max_col": ws.max_column}

            # 查找"期初余额"或"年初余额"
            for row in range(1, min(30, ws.max_row + 1)):
                for col in range(1, 10):
                    val = ws.cell(row=row, column=col).value
                    if val and any(x in str(val) for x in ["期初余额", "年初余额", "期初数", "上期审定"]):
                        if "balance_rows" not in sheet_info:
                            sheet_info["balance_rows"] = []
                        sheet_info["balance_rows"].append({"row": row, "col": col, "text": str(val)[:50]})

            # 读取Row 8-20的内容预览
            preview = {}
            for row in range(8, min(25, ws.max_row + 1)):
                row_data = []
                for col in range(1, min(12, ws.max_column + 1)):
                    val = ws.cell(row=row, column=col).value
                    if val and str(val).strip():
                        row_data.append(f"C{col}={str(val)[:40]}")
                if row_data:
                    preview[f"R{row}"] = row_data
            sheet_info["preview"] = preview

            result["sheets"][sheet_name] = sheet_info

    wb.close()
    return result


def main():
    base_dir = r"F:\AI专用\AI工作区\底稿roll forward"

    files = [
        "M SWP 应付票据 202YMMDD XYZ公司.xlsx",
        "N SWP  应付账款 202YMMDD XYZ公司.xlsx",
        "L2 SWP 长期待摊费用 202YMMDD XYZ公司.xlsx",
        "Q1 SWP 银行借款 202YMMDD XYZ公司 - 副本.xlsx",
        "C SWP 货币资金 202YMMDD XYZ公司.xlsx",
        "U_exp SWP other 202YMMDD XYZ公司.xlsx",
        "J1 SWP 在建工程 202YMMDD XYZ公司.xlsx",
        "L1 SWP 无形资产 202YMMDD XYZ公司.xlsx",
    ]

    for f in files:
        filepath = os.path.join(base_dir, f)
        if not os.path.exists(filepath):
            print(f"文件不存在: {f}")
            continue

        try:
            result = inspect_bkd_sheets(filepath)
            print(f"\n{'='*80}")
            print(f"文件: {result['filename']}")
            for sheet, info in result['sheets'].items():
                print(f"\n  [{sheet}] {info['max_row']}x{info['max_col']}")
                if "balance_rows" in info:
                    print(f"    找到余额行: {info['balance_rows']}")
                if "preview" in info:
                    for rkey, rvals in list(info['preview'].items())[:8]:
                        print(f"    {rkey}: {rvals}")
        except Exception as e:
            print(f"\n文件 {f} 检查失败: {e}")


if __name__ == "__main__":
    main()
