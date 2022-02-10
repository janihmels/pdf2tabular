from Pdf_To_Text import pdf_To_text
from PdfAdult import pdfAudit
import os
from pdfminer.pdfpage import PDFPage

path = "/exempleAudit/PROs/SoundExchange/SoundExchange/Statements/2017/75043709_JUN2017_A_Summary.pdf"


pdf_text = pdf_To_text(path, [0])


print(pdf_text)
