from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import resolve1
from io import StringIO


def pdf_To_text(path, pages, isLastpage=False):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()

    laparams = LAParams(all_texts=True, detect_vertical=True, line_overlap=0.5, char_margin=4000.0, line_margin=0.5,
                        word_margin=2, boxes_flow=1)
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    if isLastpage:
        parser = PDFParser(fp)
        document = PDFDocument(parser)
        pages = [int(resolve1(document.catalog['Pages'])['Count']) - pages[0] - 1]

    for page in PDFPage.get_pages(fp, set(pages), maxpages=0, password="", caching=True, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()
    fp.close()
    device.close()
    retstr.close()
    return text
