from subs.Pdf_To_Text import *
import os
import re
import tabula
import math


def pdfAudit(pathFile):
    pdf_text = pdf_To_textPypdf(pathFile, 0)

    if "T C F MUSIC PUBLISHING," in pdf_text:
        format = "FOX"
    else:
        pdf_text = pdf_To_text(pathFile, [1], False)
        if "WarnerChappell.com" in pdf_text or "WB MUSIC CORP" in pdf_text:
            format = "WARNERCHAPPELL"
        elif "Please note that the PRS Account" in pdf_text:
            format = "PRS"
        else:
            pdf_text = pdf_To_text(pathFile, [2], False)
            if "Please note that the PRS Account" in pdf_text:
                format = "PRS"
            else:
                pdf_text = pdf_To_text(pathFile, [0], False)
                format = FindFormat(pdf_text).upper()

    if format == "PEERMUSIC" or format == "GVL":
        pdf_text = pdf_To_text(pathFile, [1], False)
    elif format == "KOBALT":
        pdf_text = pdf_To_text(pathFile, [0], True)

    dicts = Formats(pathFile)

   # try:
    if format == "NONE":
        isSony = dicts.SONY(pdf_text)
        if "result" in isSony.keys():
            return {"result": "Format not supported " + format}
        else:
            return isSony

    return getattr(dicts, format)(pdf_text)
    #except AttributeError as e:
     #   return {"result": "Format not supported " + format}


def FindFormat(pdf_text):
    formatdict = dict()

    formatdict["BMI's Next Distribution Will Occur During: Moving? Visit bmi.com to change your address"] = "BMI"
    formatdict["American Society of Composers, Authors and Publishers"] = "ASCAP"
    formatdict["Please note that the PRS Account"] = "PRS"  # 0
    formatdict["SoundExchange"] = "SOUNDEXCHANGE"
    formatdict["BMG Rights Management"] = "BMG"
    formatdict["www.mybmg.com"] = "BMG"
    formatdict["SUMMARY STATEMENTBMG Rights Management"] = "BMG"
    formatdict["SUMMARY STATEMENTBMG Rights Management"] = "BMG"
    formatdict["Kobalt Music Services America Inc (KMSA)"] = "Kobalt"
    formatdict["Sony/ATV Music Publishing"] = "SONY"  # 8, 2, 10, 11
    formatdict["Sony Music Publishing LLC"] = "SONY"
    formatdict["UNIVERSAL MUSIC PUBL. LTD."] = "UNIVERSAL"
    formatdict["WarnerChappell.com"] = "WARNERCHAPPELL"  # 0,1
    formatdict["WB MUSIC CORP"] = "WARNERCHAPPELL"  # 0,1
    formatdict["Koda"] = "KODA"
    formatdict["AMRA"] = "AMRA"
    formatdict["Mechanical-Copyright Protection Society"] = "MCPS"
    formatdict["Howe Sound Music Publishing, LLC"] = "HOWE"
    formatdict["Beatroot"] = "CURVE"
    formatdict["SACEM"] = "SACEM"
    formatdict["Company:The Administration MP"] = "ADMINMP"
    formatdict["Company:Administration Music Rights"] = "ADMINMP"
    formatdict["UnitedMasters"] = "UnitedMasters"
    formatdict["Earnings Account Summary"] = "ENVATOMARKETPLACE"
    formatdict["peermusic"] = "peermusic"
    formatdict["WALT DISNEY MUSIC COMPANY"] = "Disney"
    formatdict["T C F MUSIC PUBLISHING,"] = "Fox"
    formatdict["BUCKS MUSIC GROUP LTD"] = "BUCKS"
    formatdict["CTM PUBLISHING BV"] = "CTM"
    formatdict["Digital Mechanical Subs"] = "CCMG"
    formatdict["Rondor Music International"] = "Rondor"
    formatdict["Reservoir Media Management"] = "RESERVOIR"
    formatdict["Company:RESERVOIR/REVERB MUSIC LTD"] = "RESERVOIR"
    formatdict["Spirit One"] = "SpiritOne"
    formatdict["Armada Music"] = "ArmadaMusic"
    formatdict["DIM MAK"] = "DIMMAK"
    formatdict["SOCAN"] = "SOCAN"
    formatdict["B-UNIQUE"] = "BUNIQUE"
    formatdict["Essential"] = "Essential"
    formatdict["ole Media Management L.P."] = "REDONE"
    formatdict["Ultra Music Publishing Europe AG"] = "Ultra"
    formatdict["PULSE PUBLISHING ADMINISTRATION, LLC"] = "PULSE"
    #    formatdict["CONCORD MUSIC PUBLISHING"] = "PULSE"
    formatdict["Prior Period Balance Brought Forward"] = "BLUEWATERMUSIC"
    formatdict["carrie@horipro.com"] = "MOJO"
    formatdict["carrie@mojomusicandmedia.com"] = "MOJO"
    formatdict["WIXEN MUSIC PUBLISHING, INC"] = "WIXEN"
    formatdict["WIXENMUSICPUBLISHING"] = "WIXEN"
    formatdict["Red Brick Music Publishing"] = "REDBRICK"
    formatdict["Notting Dale Songs Inc"] = "NOTTINGHILLMUSIC"
    formatdict["Big Machine Music"] = "BigMachine"
    formatdict["STOART"] = "STOART"
    formatdict["GVL-ID"] = "GVL"
    formatdict["www.mushroommusic.com"] = "MUSHROOM_MUSIC"
    formatdict["www.cmrra.ca"] = "CMRRA"
    formatdict["avex music publishing"] = "AVEX"
    formatdict["pubroyalty@concord.com"] = "concord"
    formatdict["Yes Dear Music"] = "HEYDAY_MEDIA"
    formatdict["heydaymediagroup"] = "HEYDAY_MEDIA"
    formatdict["Raleigh Music"] = "RALEIGH"
    formatdict["Statement ContractFrom periodTo periodPayment"] = "WARNER_MUSIC"
    # formatdict[""] = "WELK_MUSIC"  # TODO (ALSO DO PARSING)
    formatdict["Future Classic"] = "FUTURE_CLASSIC"
    formatdict["Future	Classic"] = "FUTURE_CLASSIC"
    # formatdict[""] = "PPCA"  # TODO (ALSO DO PARSING)
    formatdict["PPL costs"] = "PPL"
    formatdict["Royalty Earnings from Hal Leonard Corporation"] = "Hal_Leonard"


    keylst = list(formatdict.keys())

    if pdf_text is None:
        return pdf_text

    for i in range(len(keylst)):
        if keylst[i] in pdf_text:
            return formatdict[keylst[i]]

    return "NONE"


class Formats:

    def __init__(self, pathFile):
        self.alldict = {}
        self.pathFile = pathFile

    def BMI(self, pdf_text):
        try:
            text = pdf_text.split("\n")
            if not (
                    "Questions About Your Statement? Call:\nBMI's Next Distribution Will Occur During: Moving? Visit bmi.com to change your address" in pdf_text and
                    text[7].startswith("Royalty StatementBMI")):
                return {"result": "BMI version not supported"}

            Account_No = text[text.index("Account No:") + 2]
            US_Period = text[text.index("U.S. Performance Period:") + 3]

            details = re.findall("Account\nDateNumberNumber\n\d+-\d\d-\d\d\n.*\n\n", pdf_text)[0].split("\n")
            account_number = details[2][:9]
            distribution_date = details[2][17:]
            total_amount = float(details[3][details[3].rindex("$") + 1:].replace(",", ""))

            self.alldict["Account_No"] = Account_No
            self.alldict["U.S. Performance Period"] = US_Period
            self.alldict["payee_account_number"] = account_number
            self.alldict["distribution_date1"] = distribution_date
            self.alldict["distribution_date2"] = "Year(" + distribution_date[:4] + "), Month(" + distribution_date[
                                                                                                 5:7] + "), Day(" + distribution_date[
                                                                                                                    8:10] + ")"
            self.alldict["total_amount"] = total_amount

            return self.alldict
        except (ValueError, IndexError):
            return {"result": "BMI version are in the current Statements but is changed"}

    def ASCAP(self, pdf_text):
        try:
            text = pdf_text.split("\n")
            if "Writer International Distribution For:\nMember Name:" not in pdf_text:
                if "Writer Domestic Distribution For:\nMember Name:" not in pdf_text:
                    return {"result": "ASCAP version not supported"}
                else:
                    self.alldict["Version"] = "Domestic"
                    index = text.index("Writer Domestic Distribution For:")
            else:
                self.alldict["Version"] = "International"
                index = text.index("Writer International Distribution For:")

            payee_account_number = text[index + 2][11:]
            statement_period = text[index + 4]
            royalty = float(text[index + 6][15:].replace(",", ""))

            self.alldict["payee_account_number"] = payee_account_number
            self.alldict["statement_period"] = statement_period
            self.alldict["royalty"] = royalty

            datecheck = re.findall("Check #: \d+ Date : \d{2}-\d{2}-\d{4}\n", pdf_text)
            if datecheck:
                distribution_date = datecheck[0][datecheck[0].index("Date :") + 7:-1]
                self.alldict["distribution_date1: "] = distribution_date
                self.alldict["distribution_date2: "] = "Year(" + distribution_date[
                                                                 -4:] + "), " + "Month(" + distribution_date[
                                                                                           :2] + "), " + "Day(" + distribution_date[
                                                                                                                  3:5] + ")"

            return self.alldict

        except ValueError:
            return {"result": "ASCAP version are in the current Statements but is changed"}

    def PRS(self, pdf_text):
        try:
            ary = [0, 1, 2]
            for i in ary:
                if not pdf_text.startswith("Notice of Payment"):
                    pdf_text = pdf_To_text(self.pathFile, [i], False)
                else:
                    break
            else:
                return {"result": "PRS version not supported"}

            text = pdf_text.split("\n")
            if text[1] == "":
                pdf_text = pdf_text[0:18] + pdf_text[19:]
                text = pdf_text.split("\n")

            details = re.findall(".+Distribution Number:\n\n", pdf_text)
            details = text.index(details[0].split("\n")[0])
            if text[details + 3] == "Notice of Payment":
                details += 3

            payee_account_number = text[2][:text[2].index("CAE Number:")]
            statement_period = text[3]
            original_currency = text[details + 2][2:]
            royalty = float(text[details + 3].replace(",", ""))

            self.alldict["payee_account_number"] = payee_account_number
            self.alldict["statement_period"] = statement_period
            self.alldict["original_currency"] = original_currency
            self.alldict["royalty"] = royalty

            return self.alldict

        except ValueError:
            return {"result": "PRS version are in the current Statements but is changed"}

    def SOUNDEXCHANGE(self, pdf_text):
        try:
            text = pdf_text.split("\n")

            pdf_text = "\n" + pdf_text
            if "Featured Artist\nDigital Performance Royalty Statement" in pdf_text and "SoundExchange is a non-profit performance rights organization" in pdf_text:
                detailsIndex = pdf_text.index("Featured Artist")
                statement_period = pdf_text[pdf_text[:detailsIndex].rindex("\n"):detailsIndex - 1].split("\n ")[-1]
                idIndex = pdf_text.index("SoundExchange  Payee ID")
                payee_id = pdf_text[idIndex + 23:pdf_text[idIndex:].index("\n") + idIndex]

                royalty = float(text[text.index("Featured Artist Payment") + 2][1:].replace(",", ""))

                self.alldict["statement_period"] = statement_period
                self.alldict["payee_id"] = payee_id
                self.alldict["royalty"] = royalty
                self.alldict["version"] = "SoundExchange I"

                return self.alldict

            pdf_text = pdf_To_text(self.pathFile, [1])
            text = pdf_text.split("\n")

            if pdf_text.startswith("PAYEE:\nSHAWNTAE HARRIS"):

                details = self.findSplitedLine(pdf_text, "Payee ID:")
                Payee_ID = details[10:]
                if text[text.index(details) + 2] == "CURRENT PAYMENT:":
                    details = text[text.index(details) + 4].split(" ")
                    royalty = float(details[0][1:].replace(",", ""))
                    statement_period = details[1] + " " + details[3]
                    distribution_date = " ".join(details[1:])
                else:
                    details = self.findSplitedLine(pdf_text, "CURRENT PAYMENT:")
                    details = text[text[text.index(details) + 2]].split(" ")
                    royalty = float(details[0][1:])
                    statement_period = details[1] + " " + details[3]
                    distribution_date = " ".join(details[1:])

                self.alldict["statement_period"] = statement_period
                self.alldict["distribution_date"] = distribution_date
                self.alldict["royalty"] = royalty
                self.alldict["payee_id"] = Payee_ID
                self.alldict["version"] = "SoundExchange II"

                return self.alldict

            else:
                return {"result": "SoundExchange version not supported"}
        except (ValueError, IndexError):
            return {"result": "SoundExchange version are in the current Statements but is changed"}

    def BMG(self, pdf_text):
        try:
            text = pdf_text.split("\n")
            if text[0].startswith("BMG Rights Management (UK) Ltd."):
                details = re.findall("Date: \d{2}\/\d{2}\/\d{4}\n\nIn Account with : \(\d+\)", pdf_text)[0]
                distribution_date1 = details[6:16]
                distribution_date2 = "Year(" + distribution_date1[-4:] + "), Month(" + distribution_date1[
                                                                                       3:5] + "), " + "Day(" + distribution_date1[
                                                                                                               :2] + ")"
                payee_account_number = details[37:-1]
                statement_period = text[text.index("Date: " + distribution_date1) + 4][17:]
                original_currency = text[text.index(
                    "¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬") + 1][
                                    -3:]
                royalty = float(
                    re.findall("  \d+.\d+ ", re.findall("ROYLTS Royalties for period ending .+ \d+.\d+ ", pdf_text)[0])[
                        0])
                self.alldict["distribution_date1"] = distribution_date1
                self.alldict["distribution_date2"] = distribution_date2
                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["original_currency"] = original_currency
                self.alldict["royalty"] = royalty
                return self.alldict

            elif text[0].startswith("Summary Statement"):
                statement_period = text[0][18:]
                index = pdf_text.index("Royalty Balance")
                if pdf_text[index + 15] == "\n":
                    original_currency = text[text.index("Royalty Balance") + 2][0]
                else:
                    original_currency = pdf_text[index + 15]
                royalty = text[text.index("Royalties ") + 1][:-1].replace(",", "")
                if royalty == "":
                    royalty = text[text.index("Royalties ") + 2][:-1].replace(",", "")
                royalty = float(royalty)
                details = pdf_text.rindex("Payee: ") + 8
                payee_account_number = pdf_text[details:details + pdf_text[details:].index(")")]

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["original_currency"] = original_currency
                self.alldict["royalty"] = royalty
                return self.alldict
            elif "SUMMARY STATEMENTBMG Rights Management (UK) Ltd." in pdf_text and pdf_text.startswith("Page "):

                index = pdf_text.index("Payee   :")
                payee_account_number = pdf_text[index:pdf_text[index:].index(")") + index].replace(" ", "")[7:]

                statement_period = re.findall("\n\w+ \d{4} to \w+ \d{4}\n", pdf_text)[0][1:-1].replace(" ", " ")

                index = pdf_text.index("Amounts in ")
                original_currency = pdf_text[index:index + pdf_text[index:].index("\n")][-3:]
                royalty = float(
                    re.findall("\d+.\d+\.\d+", re.findall("ROYLTS Royalties for period ending .+\n", pdf_text)[0])[
                        0].replace(",", ""))

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["original_currency"] = original_currency
                self.alldict["royalty"] = royalty
                return self.alldict
            else:
                return {"result": "BMG version not supported"}

        except (ValueError, IndexError):
            return {"result": "BMG version are in the current Statements but is changed"}

    def KOBALT(self, pdf_text):
        try:
            text = pdf_text.split("\n")
            if text[0] == "New Royalty List":
                details = self.findSplitedLine(pdf_text, "Grand Total")
                royalty = float(details[11:].replace(",", "").replace(" ", ""))
            else:
                pdf_text = pdf_To_text(self.pathFile, [1], True)
                text = pdf_text.split("\n")
                if pdf_text.startswith("Client Name") and "Commission Summary" in pdf_text:
                    details = self.findSplitedLine(pdf_text, "Totals")
                    royalty = float(text[text.index(details)].split(" ")[-1].replace(",", "").replace(" ", ""))
                else:
                    return {"result": "KOBALT version not supported"}

            details = self.findSplitedLine(pdf_text, "Collection Period:")
            details = text.index(details)
            statement_period = text[details][18:]
            original_currency = text[details + 2][10:]
            payee_account_number = text[details + 4][14:]

            self.alldict["payee_account_number"] = payee_account_number
            self.alldict["statement_period"] = statement_period
            self.alldict["original_currency"] = original_currency
            self.alldict["royalty"] = royalty
            return self.alldict


        except (ValueError, IndexError):
            return {"result": "KOBALT version are in the current Statements but is changed"}

    def SONY(self, pdf_text):
        try:
            ary = [8, 2, 10, 11]
            for i in ary:
                if not (
                        "Financial Summary Statement" in pdf_text and "Questions regarding this statement should be directed to" in pdf_text):
                    pdf_text = pdf_To_text(self.pathFile, [i], False)
                else:
                    break
            else:
                return {"result": "SONY version not supported"}

            text = pdf_text.split("\n")

            details = self.findSplitedLine(pdf_text, "For Period")
            statement_period = details[11:]
            original_currency = re.findall("\n\n[A-Z]{3}\n\n", pdf_text)[0][2:-2]
            accountnumber = re.findall("\d{7} -", pdf_text)
            client_account_number = accountnumber[0][:-2]
            payee_account_number = accountnumber[1][:-2]

            details = text.index(self.findSplitedLine(pdf_text, "Domestic Earnings"))
            detailsIndex = 1
            while text[details - detailsIndex] != "":
                detailsIndex += 1

            detailsIndex *= 2
            royalty1 = float(text[details - (detailsIndex + 1)].replace(",", "").replace(" ", ""))
            royalty2 = float(text[details - detailsIndex].replace(",", "").replace(" ", ""))
            royalty3 = float(text[details - (detailsIndex - 1)].replace(",", "").replace(" ", ""))

            if "Sony/ATV Music Publishing" in pdf_text:
                company = "Sony ATV"
            elif "Sony Music Publishing LLC" in pdf_text:
                company = "Sony Music Publishing"
            else:
                return {"result": "SONY version not supported"}

            self.alldict["payee_account_number"] = payee_account_number
            self.alldict["statement_period"] = statement_period
            self.alldict["original_currency"] = original_currency
            self.alldict["client_account_number"] = client_account_number
            self.alldict["royalty1"] = royalty1
            self.alldict["royalty2"] = royalty2
            self.alldict["royalty3"] = royalty3
            self.alldict["company"] = company

            return self.alldict
        except (ValueError, IndexError):
            return {"result": "SONY version are in the current Statements but is changed"}

    def UNIVERSAL(self, pdf_text):
        try:
            if not (pdf_text.startswith("Payee: ") and "UNIVERSAL MUSIC PUBL. LTD." in pdf_text):
                return {"result": "UNIVERSAL version not supported"}

            text = pdf_text.split("\n")

            payee_account_number = pdf_text.split(" ")[1]
            if text[2] == "Financial Summary" and text[3].startswith("Royalty Period: "):
                statement_period = text[3][16:]
            else:
                statement_period = text[text.index(self.findSplitedLine(pdf_text, "Royalty Period"))][16:]

            royalty = float(
                re.findall("Balance last period.*\n*.* \d+.\d+.\d+\n", pdf_text)[0].split(" ")[-1].replace(",", ""))
            original_currency = self.rfindSplitedLine(pdf_text, "Balance this period ")[20:21]

            self.alldict["payee_account_number"] = payee_account_number
            self.alldict["statement_period"] = statement_period
            self.alldict["original_currency"] = original_currency
            self.alldict["royalty"] = royalty

            return self.alldict

        except (ValueError, IndexError):
            return {"result": "UNIVERSAL version are in the current Statements but is changed"}

    def WARNERCHAPPELL(self, pdf_text):
        # try:
        ary = [1, 0]
        text = pdf_text.split("\n")
        for i in ary:
            pdf_text = pdf_To_text(self.pathFile, [i])
            if "WB MUSIC CORP" in pdf_text and "S U M M A R Y   S T A T E M E N T" in pdf_text:
                details = re.findall(" PAYEE : *\(\d+\) ", pdf_text)[0]
                payee_account_number = details[details.index("(") + 1:details.index(")")]
                details = re.findall("ROYALTIES FOR PERIOD TO(.*) +(\d+(|\,\d+)\.\d+).*\n", pdf_text)
                statement_period = details[0][0].replace(" ", "")
                royalty = float(details[0][1].replace(" ", "").replace(",", ""))

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royalty"] = royalty
                return self.alldict
            elif "WarnerChappell.com" in text:
                if "WB Music Corp." in pdf_text or "WB MUSIC CORP." in pdf_text:
                    self.alldict["company"] = "WB Music Corp."
                elif "WC Music Corp." in pdf_text or "WC MUSIC CORP." in pdf_text:
                    self.alldict["company"] = "WC Music Corp."
                elif "MUSICALLSTARS PUBLISHING" in pdf_text:
                    self.alldict["company"] = "MUSICALLSTARS PUBLISHING"
                elif "MUSICALLSTARS B.V." in pdf_text:
                    self.alldict["company"] = "MUSICALLSTARS B.V."
                else:
                    return {"result": "WarnerChappell version are in the current Statements but is changed"}

                details = text.index(self.findSplitedLine(pdf_text, "Period: "))
                statement_period = text[details][8:]
                original_currency = text[details + 1][15:]
                payee_account_number = self.findSplitedLine(pdf_text, "Payee Account Code")[18:]
                royalty = float(
                    self.findSplitedLine(pdf_text, "Gross Payable Royalties ").replace(",", "").split(" ")[-1])

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["original_currency"] = original_currency
                self.alldict["royalty"] = royalty
                return self.alldict
        else:
            return {"result": "WarnerChappell version not supported"}

    # except (ValueError, IndexError):
    #    return {"result": "WarnerChappell version are in the current Statements but is changed"}

    def KODA(self, pdf_text):
        try:
            text = pdf_text.split("\n")
            isV1 = "Earnings Koda "
            if pdf_text.startswith("Mit Koda \n\nÅrsoverblik, Morten Ristorp Jensen "):
                isV1 = "År Indtjening Koda"
            if pdf_text.startswith("My Koda \n\nYear overview, Morten Ristorp Jensen ") or isV1 == "År Indtjening Koda":
                details = text.index(self.findSplitedLine(pdf_text, isV1))
                original_currency = text[details][text[details].index("(") + 1:text[details].index(")")]
                royalty = []
                statement_period = []
                for i in range(details + 2, len(text) - 2):
                    details = text[i].split(" ")[0]
                    statement_period.append(details[:4])
                    royal = details[4:]
                    if isV1 == "År Indtjening Koda":
                        royal = royal.replace(".", "").replace(",", ".")
                    royalty.append(float(royal.replace(",", "")))

                self.alldict["statement_period"] = statement_period
                self.alldict["original_currency"] = original_currency
                self.alldict["royalty"] = royalty  # problem with royalty
                return self.alldict
            elif pdf_text.startswith("My Koda \nMy Koda Account Export,"):

                statement_period = text[1].split(",")[-1][1:-1].replace(" ", " ")

                payee_account_number = re.findall("\n\d{7} ", pdf_text)[0][1:-1]

                details = text.index(self.findSplitedLine(pdf_text, "Date Description Payments"))
                original_currency = text[details][text[details].index("(") + 1:text[details].index(")")]
                royalty = []
                distribution_date = re.findall("\d+\/\d+\/\d{4}", pdf_text)

                df = tabula.read_pdf(self.pathFile, pages=1, area=(0, 410, 800, 539))
                for royal in df[0]["Payments (DKK)"]:
                    royal = float(str(royal).replace("\xad", ""))
                    if not math.isnan(royal):
                        royalty.append(royal)

                self.alldict["statement_period"] = statement_period
                self.alldict["original_currency"] = original_currency
                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["distribution_date"] = distribution_date
                self.alldict["royalty"] = royalty  # problem with royalty

                return self.alldict
            else:
                return {"result": "Koda version not supported"}

        except (ValueError, IndexError):
            return {"result": "Koda version are in the current Statements but is changed"}

    def AMRA(self, pdf_text):
        try:
            text = pdf_text.split("\n")

            if pdf_text.startswith("AMRA") and "Collection Period:" in pdf_text:
                details = text.index(self.findSplitedLine(pdf_text, "Collection Period:"))
                statement_period = text[details][18:]
                original_currency = text[details + 2][10:]
                payee_contract_id = text[details + 4][15:]

                pdf_text = pdf_To_text(self.pathFile, [0], True)
                royalty = float(self.rfindSplitedLine(pdf_text, "Grand Total").split(" ")[-1].replace(",", ""))

                self.alldict["statement_period"] = statement_period
                self.alldict["original_currency"] = original_currency
                self.alldict["payee_contract_id"] = payee_contract_id
                self.alldict["royalty"] = royalty  # problem with royalty

                return self.alldict
            elif "AMRA" in text[0] and re.search("Email: \w+@amra.com", pdf_text):
                distribution_date = re.findall("(\d+(|[a-z]{2}) [a-zA-z]+ \d{4})", pdf_text)[0][0]
                pdf_text = pdf_To_text(self.pathFile, [2])
                text = pdf_text.split("\n")
                details = text[text.index(
                    "Agreement Id Counterparty Agreement Description Currency Closing BalanceBalance Action") + 2].split(
                    " ")
                payee_account_number = details[0]
                original_currency = details[3]
                try:
                    pdf_text = pdf_To_text(self.pathFile, [3])
                    statement_period = self.findSplitedLine(pdf_text, "Collection From:")[
                                       16:] + " To " + self.findSplitedLine(pdf_text, "Collection To:")[15:]
                except ValueError:
                    pdf_text = pdf_To_text(self.pathFile, [4])
                    statement_period = self.findSplitedLine(pdf_text, "Statement From:")[
                                       15:] + " To " + self.findSplitedLine(pdf_text, "Statement To:")[14:]
                    royalty = float(self.rfindSplitedLine(pdf_text, "Grand Total").split(" ")[-1])
                else:
                    pdf_text = pdf_To_text(self.pathFile, [4])
                    royalty = float(self.rfindSplitedLine(pdf_text, "Totals").split(" ")[-1])

                self.alldict["statement_period"] = statement_period
                self.alldict["original_currency"] = original_currency
                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["royalty"] = royalty
                self.alldict["distribution_date"] = distribution_date

                return self.alldict
            elif "Total" in pdf_text and "AMRA" in pdf_text and "Collection Period" in pdf_text:
                details = text.index(self.findSplitedLine(pdf_text, "Collection Period:"))
                statement_period = text[details][18:]
                original_currency = text[details + 2][10:]
                payee_contract_id = text[details + 4][15:]

                pdf_text = pdf_To_text(self.pathFile, [0], True)
                royalty = float(self.rfindSplitedLine(pdf_text, "Total").split(" ")[-1].replace(",", ""))

                self.alldict["statement_period"] = statement_period
                self.alldict["original_currency"] = original_currency
                self.alldict["payee_contract_id"] = payee_contract_id
                self.alldict["royalty"] = royalty

                return self.alldict
            else:
                return {"result": "ARMA version not supported"}

        except (ValueError, IndexError):
            return {"result": "ARMA version are in the current Statements but is changed"}

    def MCPS(self, pdf_text):
        try:
            if pdf_text.startswith("Notice of Payment") and "MCPS" in pdf_text:
                text = pdf_text.split("\n")
                payee_account_number = text[2][:text[2].index("CAE Number:")]
                statement_period = text[3][2:-1]

                original_currency = text[10][2:]
                royalty = float(text[11].replace(",", ""))

                self.alldict["statement_period"] = statement_period
                self.alldict["original_currency"] = original_currency
                self.alldict["payee_contract_id"] = payee_account_number
                self.alldict["royalty"] = royalty
                return self.alldict
            else:
                return {"result": "MCPS version not supported"}
        except (ValueError, IndexError):
            return {"result": "MCPS version are in the current Statements but is changed"}

    def HOWE(self, pdf_text):
        try:
            if pdf_text.startswith("New Royalty List\n\nHowe Sound Music Publishing, LLC"):
                statement_period = self.findSplitedLine(pdf_text, "Collection Period:")[18:]
                original_currency = self.findSplitedLine(pdf_text, "Currency: ")[10:]
                payee_account_number = self.findSplitedLine(pdf_text, "Agreement Id: ")[14:]
                pdf_text = pdf_To_text(self.pathFile, [0], True)
                royalty = float(self.rfindSplitedLine(pdf_text, "Grand Total ")[11:].replace(",", ""))

                self.alldict["statement_period"] = statement_period
                self.alldict["original_currency"] = original_currency
                self.alldict["payee_contract_id"] = payee_account_number
                self.alldict["royalty"] = royalty
                return self.alldict

            else:
                return {"result": "HOWE version not supported"}
        except (ValueError, IndexError):
            return {"result": "HOWE version are in the current Statements but is changed"}

    def CURVE(self, pdf_text):
        try:
            text = pdf_text.split("\n")
            if "DISTRIBUTION STATEMENT" in pdf_text and pdf_text.startswith("Beatroot"):
                statement_period = " ".join(
                    text[text.index(self.findSplitedLine(pdf_text, "DISTRIBUTION STATEMENT")) + 2].split(" ")[:-2])
                royalties = float(self.findSplitedLine(pdf_text, "TOTAL INCOME ")[14:].replace(",", ""))
                original_currency = self.rfindSplitedLine(pdf_text, "All amounts are in ")[19:]

                self.alldict["original_currency"] = original_currency
                self.alldict["royalties"] = royalties
                self.alldict["statement_period"] = statement_period
                return self.alldict
            else:
                return {"result": "CURVE version not supported"}
        except (ValueError, IndexError):
            return {"result": "CURVE version are in the current Statements but is changed"}

    def SACEM(self, pdf_text):
        try:

            if "SACEM - RELEVÉ DE VOS DROITS D'AUTEUR" in pdf_text:
                text = pdf_text.split("\n")

                details = text.index(self.findSplitedLine(pdf_text, "RÉPARTITION"))
                statement_period = " ".join(text[details].split(" ")[-3:])
                rights_type = text[details + 2][:text[details + 2].index("AVANT") - 1]
                royalties = float(text[details + 2][len(rights_type) + 28:-1].replace(",", "").replace("\xa0", ""))
                original_currency = text[details + 2][-1]
                try:
                    payee_contract_id = self.findSplitedLine(pdf_text, "N° compte : ")[12:]
                except ValueError:
                    payee_contract_id = self.findSplitedLine(pdf_text, "N° de compte : ")[14:]

                self.alldict["statement_period"] = statement_period
                self.alldict["rights_type"] = rights_type
                self.alldict["royalties"] = royalties
                self.alldict["payee_contract_id"] = payee_contract_id
                self.alldict["original_currency"] = original_currency
                return self.alldict
            else:
                return {"result": "SACEM version not supported"}

        except (ValueError, IndexError):
            return {"result": "SACEM version are in the current Statements but is changed"}

    def ADMINMP(self, pdf_text):
        try:
            text = pdf_text.split("\n")

            if "Client Royalty Summary" in pdf_text and (
                    "Company:Administration Music Rights" in pdf_text or "Company:The Administration MP" in pdf_text):
                payee_account_number = self.findSplitedLine(pdf_text, "Payee: ").split(" ")[-1].replace(".", "")[1:-1]
                statement_period = self.findSplitedLine(pdf_text, "Quarterly for period ")[21:]
                royalties = float(text[text.index("TOTAL ROYALTIES") + 7].replace(",", ""))

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royalties"] = royalties
                return self.alldict
            else:
                return {"result": "ADMINMP version not supported"}

        except (ValueError, IndexError):
            return {"result": "ADMINMP version are in the current Statements but is changed"}

    def UNITEDMASTERS(self, pdf_text):
        try:
            text = pdf_text.split("\n")

            if "UnitedMasters" in pdf_text:
                statement_period = self.findSplitedLine(pdf_text, "Reference")[10:text[2].index("_")]
                distribution_date = self.findSplitedLine(pdf_text, "Created on ")[11:]

                try:
                    page5 = pdf_To_text(self.pathFile, [4])
                    details = self.findSplitedLine(page5, "Total Balance").split(" ")[-1]
                except ValueError:
                    page5 = pdf_To_text(self.pathFile, [3])
                    details = self.findSplitedLine(page5, "Total Balance").split(" ")[-1]

                royalties = float(details[1:details[1:].index(details[0]) + 1])

                self.alldict["distribution_date"] = distribution_date
                self.alldict["statement_period"] = statement_period
                self.alldict["royalties"] = royalties
                return self.alldict
            else:
                return {"result": "UnitedMasters version not supported"}
        except (ValueError, IndexError):
            return {"result": "UnitedMasters version are in the current Statements but is changed"}

    def ENVATOMARKETPLACE(self, pdf_text):
        try:
            text = pdf_text.split("\n")

            if "Earnings Account Summary" in text[0]:
                statement_period = self.findSplitedLine(pdf_text, "Period:")[7:]
                royalties = float(text[text.index(
                    self.findSplitedLine(pdf_text, "Income Summary to Earnings Account Amount")) + 1].split(" ")[-1][
                                  1:].replace(",", ""))
                original_currency = self.findSplitedLine(pdf_text, "Total: ")[7:10]

                self.alldict["statement_period"] = statement_period
                self.alldict["royalties"] = royalties
                self.alldict["original_currency"] = original_currency

                return self.alldict
            else:
                return {"result": "ENVATOMARKETPLACE version not supported"}
        except (ValueError, IndexError):
            return {"result": "ENVATOMARKETPLACE version are in the current Statements but is changed"}

    def PEERMUSIC(self, pdf_text):
        try:
            pdf_text = pdf_text.replace("\xa0", " ")
            text = pdf_text.split("\n")
            if "SUMMARY STATEMENT" == text[0]:
                payee_account_number = self.findSplitedLine(pdf_text, "Payee: ")[8:-1]
                statement_period = text[text.index("For the Period:") + 2]
                royaliy = float(
                    re.findall("([A-Za-z]\d+(,|)\d+\.\d{2}Balance( | )this( | )Period)", pdf_text)[0][0][1:-20].replace(
                        ",", ""))

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royaliy"] = royaliy
                return self.alldict
            elif text[1].startswith("Earnings") and "Acct" in text[0]:
                payee_account_number = text[0].split(" ")[-1][:-1]
                statement_period = text[1][16:]
                details = text[text.index("BALANCE") + 1].split(" ")
                royalties = float(details[-1].replace(",", ""))
                original_currency = details[0]

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royaliy"] = royalties
                self.alldict["original_currency"] = original_currency
                return self.alldict

            else:
                return {"result": "PEERMUSIC version not supported"}
        except (ValueError, IndexError):
            return {"result": "PEERMUSIC version are in the current Statements but is changed"}

    def DISNEY(self, pdf_text):
        return self.BasicStatement(pdf_text, "WALT DISNEY MUSIC COMPANY", "Fox")

    def FOX(self, pdf_text):
        return self.BasicStatement(pdf_text, "T C F MUSIC PUBLISHING,", "Fox")

    def BUCKS(self, pdf_text):
        return self.BasicStatement(pdf_text, "BUCKS MUSIC GROUP LTD", "BUCKS")

    def CTM(self, pdf_text):
        self.alldict = self.BasicStatement(pdf_text, "CTM PUBLISHING BV", "CTM")
        pdf_text = pdf_To_text(self.pathFile, [0], True)
        royalties = float(self.findSplitedLine(pdf_text, "Statement Total").split(" ")[-1].replace(",", ""))
        self.alldict["royaliy"] = royalties
        return self.alldict

    def BasicStatement(self, pdf_text, startwith, company):
        try:
            pdf_text = pdf_To_textPypdf(self.pathFile, 0)
            if pdf_text.startswith(startwith):
                statement_period = re.findall("For the Period +: +(\w+ \d+ \w+ \w+ \d+)\w", pdf_text)[0]
                payee_account_number = re.findall("In Account with +: +(\(\d+\))", pdf_text)[0][1:-1]
                try:
                    details = re.findall("Balance this period +: +(\w*.) +(\d+(,|)\d+\.\d{2}) ", pdf_text)[0]
                    royalties = float(details[1].replace(",", ""))
                    original_currency = details[0]
                except:
                    royalties = None
                    original_currency = None  # check later

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royaliy"] = royalties
                self.alldict["original_currency"] = original_currency
                return self.alldict
            else:
                return {"result": company + " version not supported"}
        except (ValueError, IndexError):
            return {"result": company + " version are in the current Statements but is changed"}

    def CCMG(self, pdf_text):
        try:
            text = pdf_text.split("\n")
            if pdf_text.startswith("Publishing Detail Statement"):
                statement_period = text[1][15:]
                pdf_text = pdf_To_text(self.pathFile, [0], True)
                details = re.findall("Gross Royalties Earned this Statement (.+) \n", pdf_text)[0].replace(",",
                                                                                                           "").replace(
                    " ", "")
                royaliy = float(details[1:])
                original_currency = details[0]
                payee_account_number = re.findall("Payee:.+(\(\d+\))", pdf_text)[0][1:-1]

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royaliy"] = royaliy
                self.alldict["original_currency"] = original_currency
                return self.alldict
            else:
                return {"result": "CCMG version not supported"}
        except (ValueError, IndexError):
            return {"result": "CCMG version are in the current Statements but is changed"}

    def RONDOR(self, pdf_text):
        try:
            text = pdf_text.split("\n")
            if text[0] == "Rondor Music International":

                text = pdf_text.split("\n")
                payee_account_number = text[2].split(" ")[-1][1:-1]
                if "Client: " in pdf_text:
                    pdf_text = pdf_To_text(self.pathFile, [14])
                    client_account_number = re.findall("Client: (\w+) - ", pdf_text)[0]
                    royalties = float(
                        text[text.index(self.findSplitedLine(pdf_text, "Balance last period")) + 1].split(" ")[
                            -1].replace(",", ""))
                    self.alldict["payee_account_number"] = client_account_number
                else:
                    pdf_text = pdf_To_text(self.pathFile, [0], True)
                    royalties = float(
                        self.findSplitedLine(pdf_text, "Final Totals").split(" ")[-3][1:].replace(",", ""))

                statement_period = self.findSplitedLine(pdf_text, "Royalty Period: ")[17:]

                self.alldict["statement_period"] = statement_period
                self.alldict["royaliy"] = royalties

                return self.alldict
            else:
                return {"result": "Rondor version not supported"}
        except (ValueError, IndexError):
            return {"result": "Rondor version are in the current Statements but is changed"}

    def RESERVOIR(self, pdf_text):
        try:
            text = pdf_text.split("\n")
            if "Payee: Company:" in pdf_text:
                payee_account_number = text[text.index(self.findSplitedLine(pdf_text, "Payee:")) + 2].split(" ")[-1][
                                       1:-2]
                statement_period = text[text.index(self.findSplitedLine(pdf_text, "In Account with:")) + 2]
                royaliy = float(text[text.index("TOTAL TRANSACTIONS") + 8].replace(",", ""))

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royaliy"] = royaliy
                return self.alldict

            elif text[0] == "Client Royalty Summary":
                payee_account_number = re.findall("Payee:.+\((\d+)\)", pdf_text)[0]
                statement_period = \
                    re.findall("\n.(.+\d+(\.|\/)\d+(\.|\/)\d+ to \d+(\.|\/)\d+(\.|\/)\d+)\n", pdf_text)[0][0]
                details = self.findSplitedLine(pdf_text, "TOTAL ROYALTIES")
                original_currency = details[16]
                if original_currency.isnumeric():
                    royaliy = float(details[16:].replace(",", ""))
                    original_currency = self.findSplitedLine(pdf_text, "BALANCE CARRIED FORWARD")[-3:]
                else:
                    royaliy = float(details[17:].replace(",", ""))

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royaliy"] = royaliy
                self.alldict["original_currency"] = original_currency

                return self.alldict
            else:
                return {"result": "Reservoir version not supported"}
        except (ValueError, IndexError):
            return {"result": "Reservoir version are in the current Statements but is changed"}

    def SPIRITONE(self, pdf_text):
        try:
            if "Client Royalty Summary" in pdf_text:
                payee_account_number = re.findall("Payee:.+\((\d+)\)", pdf_text)[0]
                statement_period = \
                    re.findall("\n(.+\d+(\.|\/)\d+(\.|\/)\d+ to \d+(\.|\/)\d+(\.|\/)\d+)\n", pdf_text)[0][0]

                pdf_text = pdf_To_text(self.pathFile, [1])
                text = pdf_text.split("\n")

                for i in range(1, 4):
                    if "BALANCE CARRIED FORWARD" in text[-3]:
                        details = text.index("TOTAL ROYALTIES")
                        if text[details + 8] == "Royalty Transfers":
                            details += 2
                        royalty = float(text[details + 8].replace(",", ""))
                        break
                    elif "BALANCE CARRIED FORWARD" in text[-7]:
                        royalty = float(text[0].replace(",", ""))
                        break
                    else:
                        pdf_text = pdf_To_text(self.pathFile, [i])
                        text = pdf_text.split("\n")
                else:
                    return {"result": "SPIRITONE version not supported"}

                original_currency = text[-3][:3]
                if not original_currency.isalpha():
                    pdf_text = pdf_To_text(self.pathFile, [3])
                    text = pdf_text.split("\n")
                    original_currency = text[-3][:3]

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royaliy"] = royalty
                self.alldict["original_currency"] = original_currency

                return self.alldict
            else:
                return {"result": "SPIRITONE version not supported"}
        except (ValueError, IndexError):
            return {"result": "SPIRITONE version are in the current Statements but is changed"}

    def ARMADAMUSIC(self, pdf_text):
        try:
            if pdf_text.startswith("Royalty Summary Page"):
                original_currency = self.rfindSplitedLine(pdf_text, "All amounts printed in ")[23:]
                pdf_text = pdf_To_text(self.pathFile, [1])
                payee_account_number = self.findSplitedLine(pdf_text, "Account:").split(" ")[-1][1:-1]
                statement_period = self.findSplitedLine(pdf_text, "FOR PERIOD")[11:]
                royaliy = float(self.findSplitedLine(pdf_text, "TOTAL ROYALTIES").split(" ")[-1].replace(",", ""))

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royaliy"] = royaliy
                self.alldict["original_currency"] = original_currency
                return self.alldict
            else:
                return {"result": "ARMADAMUSIC version not supported"}
        except (ValueError, IndexError):
            return {"result": "ARMADAMUSIC version are in the current Statements but is changed"}

    def DIMMAK(self, pdf_text):
        try:
            if pdf_text.startswith("Royalty Summary Page"):
                original_currency = self.rfindSplitedLine(pdf_text, "All amounts printed in ")[23:]
                payee_account_number = self.findSplitedLine(pdf_text, "Account:").split(" ")[-1][1:-1]
                statement_period = self.findSplitedLine(pdf_text, "FOR PERIOD")[11:]
                royaliy = float(self.findSplitedLine(pdf_text, "TOTAL ROYALTIES").split(" ")[-1].replace(",", ""))

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royaliy"] = royaliy
                self.alldict["original_currency"] = original_currency
                return self.alldict
            else:
                return {"result": "DIMMAK version not supported"}
        except (ValueError, IndexError):
            return {"result": "DIMMAK version are in the current Statements but is changed"}

    def ESSENTIAL(self, pdf_text):
        text = pdf_text.split("\n")
        try:
            if pdf_text.startswith("Client Royalty Summary"):
                payee_account_number = re.findall("Payee:.+\((\d+ \/ \d+)\)", pdf_text)[0]
                detalils = text.index(self.findSplitedLine(pdf_text, "In Account with:"))
                client_account_number = text[detalils].split(" ")[-1][1:-1]
                statement_period = text[detalils + 2]
                pdf_text = pdf_To_text(self.pathFile, [1])

                royaliy = float(self.findSplitedLine(pdf_text, "TOTAL TRANSACTIONS").split(" ")[-1].replace(",", ""))

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royaliy"] = royaliy
                self.alldict["client_account_number"] = client_account_number
                return self.alldict
            else:
                return {"result": "ESSENTIAL version not supported"}
        except (ValueError, IndexError):
            return {"result": "ESSENTIAL version are in the current Statements but is changed"}

    def SOCAN(self, pdf_text):
        text = pdf_text.split("\n")
        try:
            if "Member Statement" in text:
                details = text.index(self.findSplitedLine(pdf_text, "SOCAN NO"))
                payee_account_number = text[details][10:]
                statement_period = text[details + 2][13:]
                distribution_date = text[details + 3][18:]
                royalties = self.findSplitedLine(pdf_text, "Earnings").split(" ")[-1]
                original_currency = "CAD"

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royaliy"] = royalties
                self.alldict["distribution_date"] = distribution_date
                self.alldict["original_currency"] = original_currency

                return self.alldict
            else:
                return {"result": "SOCAN version not supported"}
        except (ValueError, IndexError):
            return {"result": "SOCAN version are in the current Statements but is changed"}

    def BUNIQUE(self, pdf_text):
        text = pdf_text.split("\n")
        try:
            if "Client Royalty Summary" in text:
                details = re.findall("(In Account with: .+\((\d+)\).*)\n", pdf_text)[0]
                payee_account_number = details[1]
                statement_period = text[text.index(details[0]) + 2]
                details = text.index(self.findSplitedLine(pdf_text, "TOTAL ROYALTIES")) + 10
                if "BALANCE CARRIED FORWARD" in text[details]:
                    details -= 12

                royalties = float(text[details][1:].replace(",", ""))
                original_currency = text[details][0]

                self.alldict["payee_account_number"] = payee_account_number
                self.alldict["statement_period"] = statement_period
                self.alldict["royalty"] = royalties
                self.alldict["original_currency"] = original_currency

                return self.alldict
            else:
                return {"result": "SOCAN version not supported"}

        except (ValueError, IndexError):
            return {"result": "SOCAN version are in the current Statements but is changed"}

    def REDONE(self, pdf_text):

        pdf_text = pdf_To_text(path=self.pathFile,
                               pages=[0, 1])

        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        payee_account_number_row = rows[6]

        try:
            payee_account_number_start_idx = \
            [i for i in range(len(payee_account_number_row)) if payee_account_number_row[i] == '('][0]
            payee_account_number_end_idx = \
            [i for i in range(len(payee_account_number_row)) if payee_account_number_row[i] == ')'][0]
            payee_account_number = payee_account_number_row[
                                   payee_account_number_start_idx + 1: payee_account_number_end_idx]
        except IndexError:
            payee_account_number_row = rows[7]
            payee_account_number_start_idx = \
            [i for i in range(len(payee_account_number_row)) if payee_account_number_row[i] == '('][0]
            payee_account_number_end_idx = \
            [i for i in range(len(payee_account_number_row)) if payee_account_number_row[i] == ')'][0]
            payee_account_number = payee_account_number_row[
                                   payee_account_number_start_idx + 1: payee_account_number_end_idx]

        client_account_number_row = rows[13]
        client_account_number_start_idx = max(
            [i for i in range(len(client_account_number_row)) if client_account_number_row[i] == '('])
        client_account_number_end_idx = max(
            [i for i in range(len(client_account_number_row)) if client_account_number_row[i] == ')'])
        client_account_number = client_account_number_row[
                                client_account_number_start_idx + 1: client_account_number_end_idx]

        period_row = rows[14]
        period_row_splitted = period_row.split()
        to_index = period_row_splitted.index('to')
        period_start = period_row_splitted[to_index - 1]
        period_end = period_row_splitted[to_index + 1]

        royalties = rows[rows.index('TOTAL ROYALTIES') - 1]

        if not (re.match('[0-9]*[.][0-9]*', royalties) is None):
            span = re.match('[0-9]*[.][0-9]*', royalties).span()

            if not (span[0] == 0 and span[1] == len(royalties)):
                # royalties is not a decimal number
                royalties = rows[rows.index('Royalty Transfers') + 1]

        original_currency = re.search('[A-Z]*',
                                      rows[max([i for i in range(len(rows)) if
                                                'BALANCE CARRIED FORWARD' in rows[i]])]).group(0)

        self.alldict['payee_account_number'] = payee_account_number
        self.alldict['client_account_number'] = client_account_number
        self.alldict['statement_period'] = period_start + ' - ' + period_end
        self.alldict['royalty'] = royalties
        self.alldict['original_currency'] = original_currency

        return self.alldict

    def ULTRA(self, pdf_text):
        try:
            rows = pdf_text.split('\n')

            rows = [item.strip() for item in rows]
            rows = [item for item in rows if item != '']

            payee_account_number_row = rows[3]
            payee_account_number_start_idx = \
                [i for i in range(len(payee_account_number_row)) if payee_account_number_row[i] == '('][0]
            payee_account_number_end_idx = \
                [i for i in range(len(payee_account_number_row)) if payee_account_number_row[i] == ')'][0]
            payee_account_number = payee_account_number_row[
                                   payee_account_number_start_idx + 1: payee_account_number_end_idx]

            period_row = rows[[i for i in range(len(rows)) if 'Half-Yearly for period' in rows[i]][0]]
            period_row_splitted = period_row.split()
            to_index = period_row_splitted.index('to')
            period_start = period_row_splitted[to_index - 1]
            period_end = period_row_splitted[to_index + 1]

            royalties = rows[rows.index('TOTAL ROYALTIES') - 1]
            original_currency_row = rows[[i for i in range(len(rows)) if 'BALANCE CARRIED FORWARD' in rows[i]][0]]
            original_currency = re.search('[A-Z]*', original_currency_row).group(0)

            self.alldict['payee_account_number'] = payee_account_number
            self.alldict['statement_period'] = period_start + ' - ' + period_end
            self.alldict['royalty'] = royalties
            self.alldict['original_currency'] = original_currency
            return self.alldict
        except (ValueError, IndexError, AttributeError):
            return {"result": "ULTRA version are in the current Statements but is changed"}

    '''
    def CURVE(self, pdf_text):
        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        period = rows[rows.index('DISTRIBUTION STATEMENT') + 1].split()[0]

        royalties = rows[[i for i in range(len(rows)) if 'TOTAL INCOME' in rows[i]][0]].split()[-1][1:]
        original_currency_row = rows[[i for i in range(len(rows)) if 'All amounts are in' in rows[i]][0]]
        original_currency = original_currency_row.split()[-1]

        self.alldict['period'] = period
        self.alldict['royalty'] = royalties
        self.alldict['original_currency'] = original_currency

    def HOWE(self, pdf_text):

        pdf_text_last_page = pdf_To_text(path='../exempleAudit/AGR664240_31_December_2020_clients_client_royalty_list.pdf',
                                         pages=[0],
                                         isLastpage=True)

        rows_last_page = pdf_text_last_page.split('\n')

        rows_last_page = [item.strip() for item in rows_last_page]
        rows_last_page = [item for item in rows_last_page if item != '']

        period = rows_last_page[min([i for i in range(len(rows_last_page)) if 'Collection Period:' in rows_last_page[i]])].split(':')[1]
        currency = rows_last_page[min([i for i in range(len(rows_last_page)) if 'Currency:' in rows_last_page[i]])].split(':')[1][1:]
        payee_account_number = rows_last_page[min([i for i in range(len(rows_last_page)) if 'Agreement Id:' in rows_last_page[i]])].split(':')[1][1:]
        royalties = rows_last_page[max([i for i in range(len(rows_last_page)) if 'Grand Total' in rows_last_page[i]])].split()[2]

        self.alldict['period'] = period
        self.alldict['royalty'] = royalties
        self.alldict['original_currency'] = currency
        self.alldict['payee_account_number'] = payee_account_number
    '''

    def PULSE(self, pdf_text):

        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']


        payee_account_number_row = rows[2]
        payee_account_number = re.findall("\(([0-9]+)\)", payee_account_number_row)[0]

        period_row = rows[[i for i in range(len(rows)) if 'for period' in rows[i]][0]].split()
        period_start = period_row[-3]
        period_end = period_row[-1]

        royalties = rows[rows.index('Royalty Transfers') + 1]

        currency_row = rows[max([i for i in range(len(rows)) if 'BALANCE CARRIED FORWARD' in rows[i]])]
        currency = re.search("[A-Z]*", currency_row).group(0)

        self.alldict['payee_account_number'] = payee_account_number
        self.alldict['statement_period'] = period_start + ' - ' + period_end
        self.alldict['royalty'] = royalties
        self.alldict['original_currency'] = currency

        return self.alldict

    def BLUEWATERMUSIC(self, pdf_text):

        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        period_idx = [i for i in range(len(rows)) if 'Period:' in rows[i]][0]
        period_row = rows[period_idx].split()
        period_start = period_row[-3]
        period_end = period_row[-1]

        royalties_row = rows[[i for i in range(len(rows)) if 'Balance Due:' in rows[i]][0]].split()
        royalties = royalties_row[1]

        original_currency = royalties_row[0]

        payee_account_number_row = rows[period_idx + 1]
        payee_account_number = payee_account_number_row.split('-')[1].strip()[1:-1]

        self.alldict['payee_account_number'] = payee_account_number
        self.alldict['statement_period'] = period_start + ' - ' + period_end
        self.alldict['royalty'] = royalties
        self.alldict['original_currency'] = original_currency

        return self.alldict

    def MOJO(self, pdf_text):

        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        period_idx = [i for i in range(len(rows)) if 'for period' in rows[i]][0]
        period_row = rows[period_idx].split()
        period_start = period_row[-3]
        period_end = period_row[-1]

        royalties_row = rows[[i for i in range(len(rows)) if 'AMOUNT OWED:' in rows[i]][0]]
        royalties = royalties_row[:-len('AMOUNT OWED:')]

        original_currency, royalties = royalties.split()

        self.alldict['statement_period'] = period_start + ' - ' + period_end
        self.alldict['royalty'] = royalties
        self.alldict['original_currency'] = original_currency

        return self.alldict

    def WIXEN(self, pdf_text):
        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        payee_account_number = rows[0].split(':')[1].split()[0][1:-1]

        period_idx = [i for i in range(len(rows)) if 'ForthePeriod:' in rows[i]][0]
        period = rows[period_idx].split(':')[1]
        to_index = period.index('to')
        from_year = period[to_index - 4: to_index]
        from_month = period[0: to_index - 4]
        to_year = period[-4:]
        to_month = period[to_index + 2: -4]

        period = from_month + ' ' + from_year + ' - ' + to_month + ' ' + to_year

        royalties_row = rows[[i for i in range(len(rows)) if 'ROYLTSRoyaltiesforperiodending' in rows[i]][0]]
        royalties = re.search("([0-9]*)['.']([0-9]*)", royalties_row.split()[-1]).group(0)

        currency_row = rows[[i for i in range(len(rows)) if 'Balancethisperiod' in rows[i]][0]]
        original_currency = currency_row.split(':')[1].strip()[0]

        self.alldict['statement_period'] = period
        self.alldict['royalty'] = royalties
        self.alldict['original_currency'] = original_currency
        return self.alldict

    def REDBRICK(self, pdf_text):
        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        payee_account_number_idx = [i for i in range(len(rows)) if 'Payee:' in rows[i]][0] + 1
        payee_account_number = rows[payee_account_number_idx].split()[-1].strip()[1:-2]

        period_row = rows[[i for i in range(len(rows)) if 'for period' in rows[i]][0]].split()
        period_start = period_row[-3]
        period_end = period_row[-1]

        try:
            royalties_idx = rows.index('TOTAL TRANSACTIONS')
            royalties_idx = royalties_idx + 3 if 'DB' in rows else royalties_idx + 2
            royalties = rows[royalties_idx]
        except ValueError:
            # 'TOTAL TRANSACTIONS' not found in the 'rows' list
            royalties_idx = min([i for i in range(len(rows)) if 'TOTAL ROYALTIES' in rows[i]])
            royalties = re.search('([0-9]*[.,][0-9]*)*', rows[royalties_idx]).group(0)

        self.alldict['payee_account_number'] = payee_account_number
        self.alldict['statement_period'] = period_start + ' - ' + period_end
        self.alldict['royalty'] = royalties

        return self.alldict

    def NOTTINGHILLMUSIC(self, pdf_text):

        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        payee_account_number_idx = [i for i in range(len(rows)) if 'In Account with:' in rows[i]][0]
        payee_account_number = rows[payee_account_number_idx].split(':')[-1].split()[-1][1:-1]

        period_row = rows[[i for i in range(len(rows)) if 'for period' in rows[i]][0]].split()
        period_start = period_row[-3]
        period_end = period_row[-1]

        royalties_idx = rows.index('TOTAL ROYALTIES') - 1
        royalties = rows[royalties_idx]

        original_currency_row = rows[[i for i in range(len(rows)) if 'BALANCE CARRIED FORWARD' in rows[i]][0]]
        original_currency = re.search('[A-Z]*', original_currency_row).group(0)

        self.alldict['payee_account_number'] = payee_account_number
        self.alldict['statement_period'] = period_start + ' - ' + period_end
        self.alldict['royalty'] = royalties
        self.alldict['original_currency'] = original_currency

        return self.alldict

    def BIGMACHINE(self, pdf_text):
        try:
            text = pdf_text.split("\n")
            if "Client Royalty Summary" in pdf_text:
                payee_account_number = re.findall("In Account with:.+\((\d+)\)", pdf_text)[0]
                statement_period = \
                    re.findall("\n.(.+\d+(\.|\/)\d+(\.|\/)\d+ to \d+(\.|\/)\d+(\.|\/)\d+)\n", pdf_text)[0][0]
                royalties = float(text[text.index("TOTAL ROYALTIES") + 7].replace(",", ""))

                self.alldict['payee_account_number'] = payee_account_number
                self.alldict['statement_period'] = statement_period
                self.alldict['royalty'] = royalties
                return self.alldict
            elif "Golden Vault Royalty Summary" in pdf_text:
                statement_period = \
                    re.findall("\n(.+\d+(\.|\/)\d+(\.|\/)\d+ to \d+(\.|\/)\d+(\.|\/)\d+)\n", pdf_text)[0][0]
                royalties = float(
                    re.findall("AMOUNT DUE FOR PERIOD +(\d+(,|)\d+.\d{2})", pdf_text)[0][0].replace(",", ""))

                self.alldict['statement_period'] = statement_period
                self.alldict['royalty'] = royalties
                return self.alldict
            else:
                return {"result": "BIGMACHINE version not supported"}

        except (ValueError, IndexError):
            return {"result": "BIGMACHINE version are in the current Statements but is changed"}

    def STOART(self, pdf_text):
        try:
            text = pdf_text.split("\n")
            if text[0] == "STOART":
                distribution_date = self.findSplitedLine(pdf_text, "Payment from:").split(" ")[2]
                statement_period = text[text.index("Settlement of the payment repertoire:") + 2]
                original_currency = text[text.index("hment") + 2].split(" ")[-1]
                royalties = float(
                    text[text.index("hment", text.index("hment") + 1) + 2].split(" ")[-2].replace(",", ""))
                self.alldict['distribution_date'] = distribution_date
                self.alldict['statement_period'] = statement_period
                self.alldict['original_currency'] = original_currency
                self.alldict['royalties'] = royalties
                return self.alldict
            else:
                return {"result": "STOART version not supported"}

        except (ValueError, IndexError):
            return {"result": "STOART version are in the current Statements but is changed"}

    def GVL(self, pdf_text):
      #  try:
            text = pdf_text.split("\n")
            if text[0] == "Howard Simon Bernstein":
                details = re.findall("GVL-ID: (\d+) / Contract number: (\d+)",pdf_text)[0]
                client_account_number = details[0]
                payee_account_number = details[1]
                statement_period = self.findSplitedLine(pdf_text, "Distribution").split(" ")[-1]
                details = re.findall("Total amount \(rounded, please see note in glossary\) (\d+(\.|)\d+\,\d+) (.)",pdf_text)
                if details == []:
                    details = re.findall("Subtotals and total amount (\d+(,|)\d{2}) (.)",pdf_text)
                royalties = float(details[0][0].replace(",",""))
                original_currency = details[0][2]

                self.alldict['client_account_number'] = client_account_number
                self.alldict['payee_account_number'] = payee_account_number
                self.alldict['statement_period'] = statement_period
                self.alldict['royalties'] = royalties
                self.alldict['original_currency'] = original_currency
                return self.alldict
            else:
                return {"result": "GVL version not supported"}

        #except (ValueError, IndexError):
        #    return {"result": "GVL version are in the current Statements but is changed"}

    def MUSHROOM_MUSIC(self, pdf_text):

        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        payee_account_number = rows[min([i for i in range(len(rows)) if 'Payee' in rows[i]]) + 1].split()[-1][1:-2]
        statement_period = rows[min([i for i in range(len(rows)) if 'for period' in rows[i]])].split()[-3:]
        period_start = statement_period[0]
        period_end = statement_period[2]

        royalties_idx = max([i for i in range(len(rows)) if 'W/H TAX' in rows[i]]) + 1
        royalties = rows[royalties_idx]

        currency_idx = max([i for i in range(len(rows)) if 'BALANCE CARRIED FORWARD' in rows[i]])
        currency = re.search('[A-Z]+', rows[currency_idx]).group(0)

        self.alldict['payee_account_number'] = payee_account_number
        self.alldict['statement_period'] = period_start + ' - ' + period_end
        self.alldict['royalty'] = royalties
        self.alldict['original_currency'] = currency

        return self.alldict

    def CMRRA(self, pdf_text):

        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        payee_account_number = rows[min([i for i in range(len(rows)) if 'Payee Account Number' in rows[i]])].split(':')[
            1].strip()
        distribution_date = rows[min([i for i in range(len(rows)) if 'Distribution Date' in rows[i]])].split(':')[
            1].strip()
        statement_period = \
        rows[min([i for i in range(len(rows)) if 'Distribution Quarter' in rows[i]])].split(':')[1].split()[0]
        royalties_n_currency = rows[max([i for i in range(len(rows)) if 'Total Net Payable' in rows[i]])].split()[-1]

        sep_idx = re.search('[A-Z]+', royalties_n_currency).span()[0]
        royalties, currency = royalties_n_currency[:sep_idx], royalties_n_currency[sep_idx:]

        self.alldict['payee_account_number'] = payee_account_number
        self.alldict['distribution date'] = distribution_date
        self.alldict['statement_period'] = statement_period
        self.alldict['royalty'] = royalties
        self.alldict['original_currency'] = currency

        return self.alldict

    def AVEX(self, pdf_text):

        pdf_text = pdf_To_text(path=self.pathFile, pages=[0, 1])
        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        distribution_date = ' '.join(rows[min([i for i in range(len(rows)) if 'Output Date' in rows[i]])].split()[2:])
        royalties = rows[rows.index('Royalty for This Term') + 1]
        currency = rows[min([i for i in range(len(rows)) if '*If Total equals less than' in rows[i]])].split()[5]
        statement_period = rows[max([i for i in range(len(rows)) if 'period ending' in rows[i]])].split(':')[1][
                           :-1].strip()

        self.alldict['distribution date'] = distribution_date
        self.alldict['statement_period'] = statement_period
        self.alldict['royalty'] = royalties
        self.alldict['original_currency'] = currency

        return self.alldict

    def CONCORD(self, pdf_text):

        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        statement_period = rows[rows.index('Period:') + 2]
        payee_account_number = \
        rows[min([i for i in range(len(rows)) if 'In Account with' in rows[i]])].split(':')[1].split()[0].strip()[1:-1]
        currency_n_royalty = rows[max([i for i in range(len(rows)) if 'Balance this Period' in rows[i]])].split(':')[
                                 1].strip()[:-2]
        currency, royalty = currency_n_royalty[0], currency_n_royalty[1:]

        self.alldict['statement_period'] = statement_period
        self.alldict['royalty'] = royalty
        self.alldict['original_currency'] = currency
        self.alldict['payee_account_number'] = payee_account_number

        return self.alldict

    def HEYDAY_MEDIA(self, pdf_text):

        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        statement_period = rows[rows.index('CLIENT ROYALTY SUMMARY') + 1]
        currency_n_royalty = rows[min([i for i in range(len(rows)) if 'Royalty Income' in rows[i]])].split()[2]

        currency, royalty = currency_n_royalty[0], currency_n_royalty[1:]

        self.alldict['statement_period'] = statement_period
        self.alldict['royalty'] = royalty
        self.alldict['original_currency'] = currency

        return self.alldict

    def HAL_LEONARD(self, pdf_text):

        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        statement_period = rows[min([i for i in range(len(rows)) if 'For the period' in rows[i]])].split()
        period_start = statement_period[3]
        period_end = statement_period[5]

        payee_contract_id = rows[min([i for i in range(len(rows)) if 'ID' in rows[i]])].split()[2]

        last_page_pdf_text = pdf_To_text(path=self.pathFile, pages=[0], isLastpage=True)
        rows = last_page_pdf_text.split('\n')
        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        royalties = rows[max([i for i in range(len(rows)) if 'EARNINGS THIS PERIOD' in rows[i]])].split()[-1]

        self.alldict['statement_period'] = period_start + ' - ' + period_end
        self.alldict['royalty'] = royalties
        self.alldict['payee_contract_id'] = payee_contract_id

        return self.alldict

    def RALEIGH(self, pdf_text):

        rows = pdf_text.split('\n')
        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        payee_account_number = rows[min([i for i in range(len(rows)) if 'In Account with' in rows[i]])].split()[
                                   -1].strip()[1:-1]
        period = rows[min([i for i in range(len(rows)) if 'for period' in rows[i]])].split()
        period_start = period[-3]
        period_end = period[-1]
        royalties = rows[min([i for i in range(len(rows)) if 'TOTAL ROYALTIES' in rows[i]]) - 1]

        self.alldict['statement_period'] = period_start + ' - ' + period_end
        self.alldict['royalty'] = royalties
        self.alldict['payee_account_number'] = payee_account_number

        return self.alldict

    def WARNER_MUSIC(self, pdf_text):

        rows = pdf_text.split('\n')
        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        distribution_date = rows[min([i for i in range(len(rows)) if 'Print run ID' in rows[i]]) - 1]
        number_n_period = rows[rows.index("method") + 1].split()
        payee_account_number = number_n_period[1]
        period = number_n_period[2].strip()
        period_start = period[:10]
        period_end = period[10:]

        # -- second page --

        pdf_text = pdf_To_text(path=self.pathFile, pages=[1])

        rows = pdf_text.split('\n')
        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        royalties = rows[max([i for i in range(len(rows)) if 'Total Earnings' in rows[i]])].split()[2]
        currency = rows[max([i for i in range(len(rows)) if 'All amounts expressed in' in rows[i]])].split()[-1]

        self.alldict['distribution_date'] = distribution_date
        self.alldict['payee_account_number'] = payee_account_number
        self.alldict['statement_period'] = period_start + ' - ' + period_end
        self.alldict['royalty'] = royalties
        self.alldict['original_currency'] = currency

        return self.alldict

    def FUTURE_CLASSIC(self, pdf_text):

        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        rows = [item.replace('\t', ' ') for item in rows]

        try:
            rights_type = ' '.join(rows[min([i for i in range(len(rows)) if 'Statement' in rows[i]])].split()[:-1])

            try:
                client_account_number = rows[min([i for i in range(len(rows)) if 'Client #:' in rows[i]])].split(':')[
                    1].strip()
            except ValueError:
                client_account_number = None

            payee_account_number = rows[min([i for i in range(len(rows)) if 'ATTN' in rows[i]])].split(':')[1].strip()
            statement_period = rows[min([i for i in range(len(rows)) if 'Amount Check # ' in rows[i]]) - 1]

            royalty_n_currency = rows[min([i for i in range(len(rows)) if 'Amount Payable:' in rows[i]])].split(':')[
                1].strip()

            royalty = royalty_n_currency.split()[0][1:]
            currency = royalty_n_currency.split()[1].strip()[1:-1]

        except ValueError:
            # the file is from the secondary format:

            period_idx = min([i for i in range(len(rows)) if 'Period Activity Total' in rows[i]]) - 1

            rights_type = rows[period_idx - 1][:rows[period_idx - 1].find("Royalties")].strip() + ' Royalties'
            statement_period = rows[period_idx]
            client_account_number = None
            payee_account_number = None

            currency_idx = min([i for i in range(len(rows)) if 'Total Payable To Writer' in rows[i]])
            currency = rows[currency_idx].split()[-1][1:-1]
            royalty = rows[currency_idx - 1][1:]

        self.alldict['rights_type'] = rights_type
        self.alldict['client_account_number'] = client_account_number
        self.alldict['payee_account_number'] = payee_account_number
        self.alldict['statement_period'] = statement_period
        self.alldict['royalty'] = royalty
        self.alldict['original_currency'] = currency

        return self.alldict

    def PPL(self, pdf_text):

        rows = pdf_text.split('\n')

        rows = [item.strip() for item in rows]
        rows = [item for item in rows if item != '']

        try:
            payee_account_number = rows[rows.index('Member ID:') + 1]
        except ValueError:
            payee_account_number = rows[min([i for i in range(len(rows)) if 'Member ID:' in rows[i]])].split(':')[1].strip()

        statement_period = ' '.join(rows[min([i for i in range(len(rows)) if 'Payment' in rows[i]])].split()[1:3])

        # -- second page parsing

        royalty_extracted = False

        try:
            royalties = rows[max([i for i in range(len(rows)) if 'Amount:' in rows[i]])].split(':')[1].strip()[1:]
            royalty_extracted = True
            currency = rows[max([i for i in range(len(rows)) if 'Currency:' in rows[i]])].split(':')[1].strip()
        except ValueError:
            pdf_text = pdf_To_text(path=self.pathFile, pages=[1])
            rows = pdf_text.split('\n')

            rows = [item.strip() for item in rows]
            rows = [item for item in rows if item != '']

            if not royalty_extracted:
                royalties = rows[max([i for i in range(len(rows)) if 'Amount:' in rows[i]])].split(':')[1].strip()[1:]

            currency = rows[max([i for i in range(len(rows)) if 'Currency:' in rows[i]])].split(':')[1].strip()

        self.alldict['payee_account_number'] = payee_account_number
        self.alldict['statement_period'] = statement_period
        self.alldict['royalty'] = royalties
        self.alldict['original_currency'] = currency

        return self.alldict

    def findSplitedLine(self, source, text):
        detailsIndex = source.index(text)
        return source[detailsIndex:detailsIndex + source[detailsIndex:].index("\n")]

    def rfindSplitedLine(self, source, text):
        detailsIndex = source.rindex(text)
        return source[detailsIndex:detailsIndex + source[detailsIndex:].index("\n")]
