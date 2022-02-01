from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import re
import os


def pdf_To_text(path, pages):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()

    laparams = LAParams(all_texts=True, detect_vertical=True, line_overlap=0.5, char_margin=2000.0, line_margin=0.5,
                        word_margin=2, boxes_flow=1)
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    for page in PDFPage.get_pages(fp, set(pages), maxpages=0, password="", caching=True, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()
    fp.close()
    device.close()
    retstr.close()
    return text


def CheckPdf(text):
    textSplited = text.split("\n")
    if len(text) > 96 and text[:96] == "Publishing Summary Statement\nWriter ID\nAccount Name\nVendor ID\nStatement Date\nStatement Frequency":
        return "CMG"
    elif len(textSplited) > 7 and re.findall("Distribution \d\d\d\d\w+", textSplited[7]) != []:
        return "PRS"
    else:
        return "None"


path = "exemple"
filelist = []

for root, dirs, files in os.walk(path):
    for file in files:
        pathFile = os.path.join(root, file)
        publisher = pathFile.split("\\")[1]
        if publisher != "NotNow":
            pdf_text = pdf_To_text(pathFile, pages=[0])
            print(CheckPdf(pdf_text), publisher,pathFile)


