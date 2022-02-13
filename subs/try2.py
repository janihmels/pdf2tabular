from Pdf_To_Text import pdf_To_text
from PdfAdult import pdfAudit
import os


path = "C:\\Users\\Gad\\Documents\\GitHub\\pdf2tabular\\exempleAudit\\PROs\\Koda\\September 2021 - Ristorp.pdf"
pdf_text = pdf_To_text(path, [0])
print(pdf_text)
