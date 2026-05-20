import openpyxl

# 检查标准模板的K.01结构
template_path = "K1 SWP 固定资产 202YMMDD XYZ公司.xlsx"
wb = openpyxl.load_workbook(template_path, data_only=False)
ws = wb["K.01 Agree SL to GL"]

print("=== 标准模板 K.01 结构 ===")
for row in range(9, 16):
    print(f"\n第{row}行:")
    for col in range(1, 22):
        val = ws.cell(row=row, column=col).value
        if val:
            print(f"  ({col}): {str(val)[:80]}")

wb.close()
