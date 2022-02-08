from Pdf_To_Text import pdf_To_text
from PdfAdult import pdfAudit
import os

path = "C:\\Users\\Gad\\Documents\\GitHub\\pdf2tabular\\exempleAudit\\Publishers\\BMG\\BMG\\2016\\1H16 statement-080280_201606Z.PDF"

#path = "C:\\Users\\Gad\\Documents\\GitHub\\pdf2tabular\\exempleAudit\\Publishers\\BMG\\BMG\\2017\\1H17 statement-080280_201706Z.PDF"

pdf_text = pdf_To_text(path, pages=[0])

print(pdf_text)
