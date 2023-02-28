from openpyxl import Workbook, load_workbook

wb = load_workbook("services.xlsx")
ws = wb.active

for row in ws.values:
    print(row)
