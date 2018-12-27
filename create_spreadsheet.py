from os import listdir
import PyPDF2
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side, Alignment, Protection, Font

drawing_names = listdir('drawings')

def create_spreadhseet():
    wb = load_workbook('Blank MTO r1.xlsx')
    ws = wb.active
    current_column = 23
    for i in range(len(drawing_names)):
        ws.cell(row=5, column=current_column, value=drawing_names[i][:-4])
        current_column += 1

    wb.save('MTO.xlsx')
    
    ## CREATES NEW WORKBOOK
    # wb = Workbook()
    # ws = wb.active
    # ws.title = "MTO (MM)"

    # ws['E3'] = "MATERIALS"
    # ws['F3'] = "MARKUP"
    # ws['F4'] = 1
    # ws['F4'].font = Font(size=18)
    # ws['G3'] = "FAB"
    # ws['H3'] = "MARKUP"
    # ws['H4'] = 1
    # ws['H4'].font = Font(size=18)
    # ws['I3'] = "RAW RUBBER"
    # ws['J3'] = "MARKUP"
    # ws['J4'] = 1
    # ws['J4'].font = Font(size=18)
    # ws['K3'] = "RL LABOUR"
    # ws['L3'] = "MARKUP"
    # ws['L4'] = 1
    # ws['L4'].font = Font(size=18)
    # ws['M3'] = "PAINT"
    # ws['N3'] = "MARKUP"
    # ws['N4'] = 1
    # ws['N4'].font = Font(size=18)
    # ws['O3'] = "EXTRAS"
    # ws['P3'] = "MARKUP"
    # ws['P4'] = 1
    # ws['P4'].font = Font(size=18)
    # ws['Q3'] = "EXTRAS"
    # ws['R3'] = "MARKUP"
    # ws['R4'] = 1
    # ws['R4'].font = Font(size=18)

    # ws['A6'] = "DESCRIPTION"
    # ws['B6'] = "SIZE"
    # ws['C6'] = "RL THK"
    # ws['D6'] = "COMMODITY TYPE"
    # ws['E6'] = "MATERIAL COST (PER M/EA)"
    # ws['F6'] = "MATERIAL SELL (PER MM/EA)"
    # ws['G6'] = "FAB COST"
    # ws['H6'] = "FAB SELL"
    # ws['I6'] = "RUBBER COST (PER M/EA)"
    # ws['J6'] = "RUBBER SELL (PER MM/EA)"
    # ws['K6'] = "RUBBER LABOUR COST (PER M/EA)"
    # ws['L6'] = "RUBBER LABOUR SELL (PER MM/EA)"
    # ws['M6'] = "PAINT COST (PER M/EA)"
    # ws['N6'] = "PAINT SELL (PER MM/EA)"
    # ws['O6'] = "EXTRAS COST (NDE, SHIPPING, ETC)"
    # ws['P6'] = "EXTRAS SELL"
    # ws['Q6'] = "EXTRAS"
    # ws['R6'] = "EXTRAS SELL(2)"
    # ws['S6'] = "UNIT RATE"
    # ws['T6'] = "UOM"
    # ws['U6'] = "TOTAL"
    # ws['V6'] = "EXTENDED TOTALS"

    # current_column = 23
    # for i in range(len(drawing_names)):
    #     c = ws.cell(row=6, column=current_column, value=drawing_names[i][:-4])
    #     c.alignment = Alignment(text_rotation=45)
    #     current_column += 1

    # wb.save("MTO.xlsx")

create_spreadhseet()