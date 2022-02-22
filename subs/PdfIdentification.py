import re
import os
from subs.Pdf_To_Text import pdf_To_text


def CheckPdf(text, location, pdfList):
    try:
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
    except:
        return "None"


def PdfIdentifier(fullfile):
    try:
        pdflist = []
        pdf_text = pdf_To_text(fullfile, pages=[0])
        res = CheckPdf(pdf_text, fullfile, pdflist)
        if res == "Try 3":
            pdf_text = pdf_To_text(fullfile, pages=[2])
            res = CheckPdf(pdf_text, fullfile, pdflist)
        return res
    except FileNotFoundError:
        return "Error File Not Found!"


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
