import openpyxl

pmte_path = "输入；参考资料.xlsx"
wb = openpyxl.load_workbook(pmte_path, data_only=True)
ws = wb["CRA"]

print("=== 固定资产CRA详细信息（第140-155行） ===")
for row in range(140, 156):
    print(f"\n第{row}行:")
    for col in range(1, 12):
        val = ws.cell(row=row, column=col).value
        if val:
            print(f"  ({col}): {val}")

wb.close()
