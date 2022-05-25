from openpyxl import Workbook
import datetime

class ExcelTemplateCreate:

    def create_sheets(self):
        wb = Workbook()
        ws1 = wb.create_sheet("ANALISYS", 0)
        ws2 = wb.create_sheet("JANUARY", 1)
        ws3 = wb.create_sheet("FEBRUARY", 2)
        ws4 = wb.create_sheet("MARCH", 3)
        ws5 = wb.create_sheet("APRIL", 4)
        ws6 = wb.create_sheet("MAY", 5)
        ws7 = wb.create_sheet("JUNE", 6)
        ws8 = wb.create_sheet("JULY", 7)
        ws9 = wb.create_sheet("AUGUST", 8)
        ws10 = wb.create_sheet("SEPTEMBER", 9)
        ws11 = wb.create_sheet("OCTOBER", 10)
        ws12 = wb.create_sheet("NOVEMBER", 11)
        ws13 = wb.create_sheet("DECEMBER", 12)

        wb.save("budget.xlsx")

    def create_templates_for_months(self):
        wb = Workbook()
        ws1 = wb.create_sheet("ANALISYS", 0)
        ws1['A3'] = 'DATE'
        ws1['B3'] = 'BILLS'
        ws1['C3'] = 'CAR'
        ws1['D3'] = 'SHOPPING'
        ws1['E3'] = 'EATING OUT'
        ws1['F3'] = 'CLOTHES'
        ws1['G3'] = 'HOME SHOPPING'
        ws1['H3'] = 'EVENTS'
        ws1['I3'] = 'GIFTS'
        ws1['J3'] = 'AGD RTV'
        ws1['K3'] = 'CULTURE AND ENTERTAINMENT'
        ws1['L3'] = 'HOLIDAYS'
        ws1['M3'] = 'REPAIRS'
        ws1['N3'] = 'SPORT'
        ws1['O3'] = 'ANIMALS'
        ws1['P3'] = 'MEDICINE'
        ws1['Q3'] = 'INSURANCE'
        ws1['R3'] = 'BEAUTY'
        ws1['S3'] = 'OTHER STUFF'
        ws1['T3'] = 'CHARITY'

        ws1.column_dimensions["B"].width = 13
        ws1.column_dimensions["C"].width = 13
        ws1.column_dimensions["D"].width = 8

        for row in range(1):
            ws1.append(range(0, 13))

        for cell in ws1['C']:
            cell.number_format = "DD/MM/YYYY"


        wb.save("budget.xlsx")

    #def copy_cells(self):


t = ExcelTemplateCreate()
t.create_templates_for_months()
t.create_sheets()
