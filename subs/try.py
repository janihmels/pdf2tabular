from Pdf_To_Text import pdf_To_text
from PdfAudit import pdfAudit
import os

path = "C:\\Users\\Gad\\Documents\\GitHub\\pdf2tabular\\exempleAudit\\Newlyannotated"
for root, dirs, files in os.walk(path):
    for file in files:
        if file[-4:].lower() == ".pdf":

            pathFile = os.path.join(root, file)
            try:
                audit = pdfAudit(pathFile)
                if audit is None:
                    print(pathFile,"None")
                elif "result" in audit.keys():
                    print(pathFile, audit)
            except:
                print(pathFile,"Error id")
