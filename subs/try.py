from subs.Pdf_To_Text import pdf_To_text
import os
import re


pdflist = []
path = "C:/Users/Gad/Documents/GitHub/pdf2tabular/exemple/BMI/BlackstarMusic/Statement"

pdf_text = pdf_To_text(path+"/2018/1.pdf", pages=[0])

print(pdf_text)
