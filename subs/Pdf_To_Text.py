from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import resolve1
from io import StringIO
from pdfminer.pdfparser import PDFSyntaxError
import PyPDF2


def pdf_To_textPypdf(path, pages):
    try:
        pdfFileObj = open(path, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        pageObj = pdfReader.getPage(pages)
        text = pageObj.extractText()
        pdfFileObj.close()
        return text
    except:
        return "None"


def pdf_To_text(path, pages, isLastpage=False):
    try:
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
    except PDFSyntaxError:
        return path+" No /Root object! - Is this really a PDF?"
