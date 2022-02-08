from Pdf_To_Text import pdf_To_text
from PdfAdult import pdfAudit
import os


path = "C:\\Users\\Gad\\Documents\\GitHub\\pdf2tabular\\exempleAudit\\PROs\\SoundExchange\\SoundExchange"
for root, dirs, files in os.walk(path):
    for file in files:
        if file[-4:].lower() == ".pdf" and "distribution letter" not in file.lower():
            pathFile = os.path.join(root, file)
            print(pdfAudit(pathFile, "SoundExchange",0))
