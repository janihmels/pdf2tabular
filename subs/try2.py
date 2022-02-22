from Pdf_To_Text import *
from PdfAudit import pdfAudit
import os
import tabula
import math
import re

path = "C:\\Users\\Gad\\Documents\\GitHub\\pdf2tabular\\exempleAudit\\Newlyannotated\\GVL\\2012\\2012_0001_detailed_report_artist_claims_53982_eng.pdf"
pdf_text = pdf_To_text(path, [1])

print(pdf_text)

# text = pdf_To_textPypdf(path,0)
# print(text)


'''
#pdf_text = pdf_To_textPypdf(path, 0)

df = tabula.read_pdf(path,pages=1, area = (0,410,800,539))
print(df)
# area = (620.79874016,929.83181102,79.838740157,336.88818898))
    
'''