
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import re


def buildcsv():
    x = []
    file = open("dataFinish.csv", "w")  # open file
    file.write("title,artist,source,reference,product,income_type,income_period,rate,quantity,amount_received,percent_payable,amount_payable\n")  # write the first line

    nlTitle = False
    nlSource = False
    title = "20 Questions" # first artist because there is no Composition Total yet
    artist = "Amy Rigby" # first artist because there is no Composition Total yet
    source = ""

    for i in range(0, 40):  # number of pages
        pdf_text = pdf_To_text("data.pdf", pages=[i])
        x = "".join(re.findall("[\n].*\w* - \w*[\n]", pdf_text))  # ary for source
        pdf_text = pdf_text.split("\n\n")
        for line in pdf_text:
            if "Composition Total" in line:  # find title and artist
                nlTitle = True
            elif nlTitle:  # apply tiitle and artist
                titleSplit = line.split(' - ')

                if len(titleSplit) >= 2:  # if title Invalid
                    title = titleSplit[0]
                    artist = titleSplit[1]
                nlTitle = False

            if line in x:  # Find Source
                nlSource = True
                source = line
            elif nlSource: #apply source
                allContext = line.split("\n")

                for con in allContext:  # lines in source
                    if "Composition Total" not in con != "":
                        refrence = re.findall("\d{5} ", con)[0]
                        product = re.findall("\d[- .a-zA-z]*\d{5} ", con)[0][1:-6]
                        income_type = re.findall("\d{5} [a-zA-z]*", con)[0][6:]
                        income_period = re.findall("\d{2}/.\d - \d{2}/\d{2}", con)[0]
                        rate = str(float(re.findall("\d.\d{8}", con)[0]))
                        quantity = str(re.findall("  \d*,*\d+  ", con)[0][2:-2])
                        amount_received = str(float(re.findall("  \d+.\d\d  ", con)[0][2:-2]))
                        percent_payable = str(float(re.findall(" \d\d.\d\d ", con)[0][1:-1]))
                        amount_payable = str(float(re.findall("[$] \d+.\d\d ", con)[0][2:-1]))

                        file.write("\"" + title + "\",\"" + artist + "\",\"" + source + "\",\"" + refrence + "\",\"" + product + "\",\"" + income_type + "\",\"" + income_period + "\",\"" + rate + "\",\"" + quantity + "\",\"" + amount_received + "\",\"" + percent_payable + "\",\"" + amount_payable + "\"," + "\n")
                nlSource = False
    file.close()


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

