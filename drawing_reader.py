import tabula
from os import listdir
from PyPDF2 import PdfFileWriter, PdfFileReader

tabula.convert_into_by_batch('drawings_seperated', output_format="csv", lattice=True, pages='all', area=(57.88,1154.442,788.22,1674.309))
