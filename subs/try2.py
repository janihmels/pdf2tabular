from Pdf_To_Text import *
from PdfAdult import pdfAudit
import os
import tabula
import math
import re

path = "C:\\Users\\Gad\\Documents\\GitHub\\pdf2tabular\\exempleAudit\\Newlyannotated\\BUnique\\To Be Anotated\\SAM PRESTON 12 2015.pdf"
pdf_text = pdf_To_text(path, [0])

print(pdf_text)
'''
df = tabula.read_pdf(path,pages=1, area = (0,410,800,539))
print(df)
# area = (620.79874016,929.83181102,79.838740157,336.88818898))
    
'''
