from Pdf_To_Text import pdf_To_text
from PdfAdult import pdfAudit
import os


path = "/exempleAudit/PROs/SoundExchange"
for root, dirs, files in os.walk(path):
    for file in files:
        if file[-4:].lower() == ".pdf":

            pathFile = os.path.join(root, file)
            audit = pdfAudit(pathFile, "SoundExchange",0)
            print(audit)
