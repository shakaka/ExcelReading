import openpyxl
from openpyxl import load_workbook

wb = load_workbook("testFiles/basicTest.xlsx")

name_list = wb.get_sheet_names()
print(name_list)


table = wb.active
print(table['B1'].font)

table['B1'] = 'name change'

wb.save("testFiles/basicTest.xlsx")
