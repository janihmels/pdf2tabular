from Pdf_To_Text import pdf_To_text
from PdfAdult import pdfAudit
import os

path = "C:\\Users\\Gad\\Documents\\GitHub\\pdf2tabular\\exempleAudit\\PROs\\SoundExchange\\SoundExchange\\Statements\\2017\\75043709_APR2017_A_Summary.pdf"
pdf_text = pdf_To_text(path, pages=[0])

print(pdf_text)
