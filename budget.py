from openpyxl import Workbook

class CreateFile:
    wb = Workbook()
    ws = wb['Sheet']
    ws['A1'] = 'POJAZDY'
    ws['B1'] = 'ILOSC PALIWA'
    ws['C1'] = 'WAGA PALIWA'
    ws['D1'] = 'CENA'
    wb.save("budget.xlsx")


CreateFile()

