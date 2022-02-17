from Pdf_To_Text import pdf_To_text
from PdfAdult import pdfAudit
import os


path = "C:\\Users\\Gad\\Documents\\GitHub\\pdf2tabular\\exempleAudit\\Newlyannotated\\BUnique"
for root, dirs, files in os.walk(path):
    for file in files:
        if file[-4:].lower() == ".pdf":

            pathFile = os.path.join(root, file)
            audit = pdfAudit(pathFile, "BUnique",0)
            print(audit)
