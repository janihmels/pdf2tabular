import os
import re
from subs.Pdf_To_Text import pdf_To_text

from subs.PdfAdult import pdfAdultBMI

pdflist = []
path = "C:\\Users\\Gad\\Documents\\GitHub\\pdf2tabular\\exemple\\BMI\\BlackstarMusic\\Statement"
filelist = []
for root, dirs, files in os.walk(path):
    for file in files:
        pathFile = os.path.join(root, file)
        pdf_text = pdf_To_text(pathFile, pages=[0])
        print(pdfAdultBMI(pdf_text))

