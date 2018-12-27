from os import listdir
import tabula
from PyPDF2 import PdfFileReader, PdfFileWriter
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side, Alignment, Protection, Font

drawing_names = listdir('drawings')
for i in range(len(drawing_names)):
    inputpdf = PdfFileReader(open('drawings/' + drawing_names[i], 'rb'))
    for j in range(inputpdf.numPages):
        output = PdfFileWriter()
        output.addPage(inputpdf.getPage(j))
        with open('drawings_seperated/' + drawing_names[i][:-4] + "_sheet_" + str(j + 1) + '.pdf', 'wb') as outputStream:
            output.write(outputStream)

def create_spreadhseet():
    wb = load_workbook('Blank MTO r1.xlsx')
    ws = wb.active
    current_column = 23
    for i in range(len(drawing_names)):
        ws.cell(row=5, column=current_column, value=drawing_names[i][:-4])
        current_column += 1

    wb.save('MTO.xlsx')

# create_spreadhseet()