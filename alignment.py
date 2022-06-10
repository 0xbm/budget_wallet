from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

wb = load_workbook("budget.xlsx")
ws = wb["JANUARY"]

for cell in ws['1:1']: #for cell in ws['B']:
    cell.alignment = Alignment(horizontal='center')
'''
ws1.column_dimensions["B"].width = 13
ws1.column_dimensions["C"].width = 13
ws1.column_dimensions["D"].width = 8
'''