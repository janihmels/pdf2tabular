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


def CheckPdf(text, location,pdfList):
    if ord(text[0]) == 12 or ord(text[0]) == 32:  ############################# Specific cases pay attention!!!!!
        return "Try 3"

    newPdf = None

    textSplited = text.split("\n\n")

    # ------------------- Wixen -------------------
    if textSplited[0].startswith("WIXENMUSICPUBLISHING,INC. SUMMARY Payee") and textSplited[3].startswith(
            "InAccountwith:") and textSplited[4].startswith("ForthePeriod:"):
        newPdf = Pdf("Wixen", location)
        pdfList.append(newPdf)
        details = textSplited[0].split("\n")[1:]
        newPdf.addDict("Payee", textSplited[0][textSplited[0].index("Payee:"):textSplited[0].index("\n")])
        newPdf.addDict("Account", textSplited[3][14:])
        newPdf.addDict("Period", textSplited[4][13:])
        return "Wixen"
    # ------------------- Wixen -------------------

    # ------------------- CMG -------------------
    titles = ["Writer ID", "Account Name", "Statement Date"]
    if all(x in textSplited[0] for x in titles) and textSplited[2].startswith("Account Summary"):
        newPdf = Pdf("CMG", location)
        pdfList.append(newPdf)
        realtitle = textSplited[0].split("\n")
        realValue = textSplited[1].split("\n")
        for t in titles:
            newPdf.addDict(t, realValue[realtitle.index(t) - 1])
        newPdf.addDict("Account Summary", textSplited[3][:textSplited[3].index(" ")])
        return "CMG"
    # ------------------- CMG -------------------

    # ------------------- PRS -------------------
    titles = ["Member Name", "CAE Number"]
    '''
    if textSplited[0] == "Notice of Payment":      ############################# Specific cases pay attention!!!!!
        text = pdf_text[:17] + pdf_text[18:]
        textSplited = text.split("\n\n")
    '''
    currentindex = [text for text, s in enumerate(textSplited) if "Member Name" in s]
    currentindex = currentindex[0]
    if all(x in textSplited[currentindex] for x in titles) and "Distribution Number" in textSplited[currentindex + 1]:
        newPdf = Pdf("PRS", location)
        pdfList.append(newPdf)
        details = textSplited[currentindex].split("\n")
        if len(details) > 3:
            details = details[1:]
        for i in range(len(titles)):
            newPdf.addDict(titles[i], details[i][:details[i].index(titles[i])])
        titleIndex = textSplited[currentindex + 1].index("Distribution Number")
        newPdf.addDict(textSplited[currentindex + 1][titleIndex:],
                       textSplited[currentindex + 1][:titleIndex] + details[-1])
        return "PRS"
    # ------------------- PRS -------------------

    return "None"


class Pdf:
    def __init__(self, name, location):
        self.name = name
        self.location = location
        self.recDict = {}

    def addDict(self, key, value):
        self.recDict[key] = value

    def __str__(self):
        return "\n------" + "\nName: " + self.name + "\nLocation: " + self.location + "\nDict: " + str(
            self.recDict) + "\n------\n"

'''
pdflist = []
path = "exemple"
filelist = []

for root, dirs, files in os.walk(path):
    for file in files:
        pathFile = os.path.join(root, file)
        publisher = pathFile.split("\\")[1]
        if publisher != "NotNow":
            pdf_text = pdf_To_text(pathFile, pages=[0])
            res = CheckPdf(pdf_text, pathFile, pdflist)
            if res == "Try 3":
                pdf_text = pdf_To_text(pathFile, pages=[2])
                res = CheckPdf(pdf_text, pathFile, pdflist)

            if res == "None":
                print(pathFile)
            else:
                print(pdflist[-1])

            # print(res, publisher,pathFile)
'''