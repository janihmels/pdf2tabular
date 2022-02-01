from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import re

def pdf_To_text(path, pages):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()

    laparams = LAParams(all_texts=True, detect_vertical=True, line_overlap=0.5, char_margin=2000.0, line_margin=0.5, word_margin=2, boxes_flow=1)
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    for page in PDFPage.get_pages(fp, set(pages), maxpages=0, password="", caching=True,check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()
    fp.close()
    device.close()
    retstr.close()
    return text


CMG = "exemple/CMG/24d7ad18b5054c039d165c480012e6a6.pdf"
PRS = "exemple/PRS/2016/Essex David_070172307_201602B_052_089780_26722.PDF"
Wixen = "exemple/Wixen/006245/Statements/2019/Q1 2019 Wixen Music 006245 Stmt.pdf"

pdf_text = pdf_To_text(CMG, pages=[0])

print(pdf_text)