import openpyxl
import json

# 检查参考资料表中的CRA信息
pmte_path = "输入；参考资料.xlsx"
wb = openpyxl.load_workbook(pmte_path, data_only=True)

# PMTE Sheet
ws_pmte = wb["PMTE"]
print("=== PMTE 表 ===")
for row in range(1, 14):
    for col in range(1, 8):
        val = ws_pmte.cell(row=row, column=col).value
        if val:
            print(f"  ({row},{col}): {val}")

# CRA Sheet
ws_cra = wb["CRA"]
print("\n=== CRA 表前15行 ===")
for row in range(1, 15):
    for col in range(1, 12):
        val = ws_cra.cell(row=row, column=col).value
        if val:
            print(f"  ({row},{col}): {str(val)[:80]}")

wb.close()
