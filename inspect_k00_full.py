import openpyxl
import json

# 检查 K.00 Lead Sheet 完整结构
prior_path = "输入：K1 固定资产 20251231六六六有限公司.xlsx"
wb = openpyxl.load_workbook(prior_path, data_only=False)
ws = wb["K.00 Lead Sheet"]

result = {}
result['k00'] = []
for row in range(1, min(30, ws.max_row + 1)):
    for col in range(1, min(15, ws.max_column + 1)):
        val = ws.cell(row=row, column=col).value
        if val:
            result['k00'].append({'row': row, 'col': col, 'value': str(val)[:100]})

result['k01'] = []
ws2 = wb["K.01 Agree SL to GL"]
for row in range(9, 15):
    for col in range(1, min(25, ws2.max_column + 1)):
        val = ws2.cell(row=row, column=col).value
        if val:
            result['k01'].append({'row': row, 'col': col, 'value': str(val)[:100]})

with open('inspect_k00_result.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

wb.close()
print("Done")
