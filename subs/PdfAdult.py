from subs.Pdf_To_Text import pdf_To_text
import os
import re


def pdfAudit(pathFile, format, page):
    pdf_text = pdf_To_text(pathFile, [page], format == "KOBALT")
    dicts = Formats(pathFile)
    return getattr(dicts, format.upper())(pdf_text)


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
            ary = [0, 2]
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
                details = findSplitedLine(pdf_text, "Grand Total")
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
                return "Sony"

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
        try:
            ary = [1, 0]
            text = pdf_text.split("\n")
            for i in ary:
                if "WB MUSIC CORP" in pdf_text and "S U M M A R Y   S T A T E M E N T " in pdf_text:
                    details = re.findall(" PAYEE : *\(\d+\) ", pdf_text)[0]
                    payee_account_number = details[details.index("(") + 1:details.index(")")]
                    details = re.findall("ROYALTIES FOR PERIOD TO.*\n", pdf_text)[0][24:]
                    endDate = details.index(" ")
                    statement_period = details[:endDate]
                    royalty = float(details[endDate:].replace(" ", "").replace(",", ""))

                    self.alldict["payee_account_number"] = payee_account_number
                    self.alldict["statement_period"] = statement_period
                    self.alldict["royalty"] = royalty
                    return self.alldict

                elif "WarnerChappell.com" in text:
                    if "WB Music Corp." in pdf_text or "WB MUSIC CORP." in pdf_text:
                        self.alldict["company"] = "WB Music Corp."
                    elif "WC Music Corp." in pdf_text or "WC MUSIC CORP." in pdf_text:
                        self.alldict["company"] = "WC Music Corp."
                    else:
                        return {"result": "WarnerChappell version are in the current Statements but is changed"}

                    details = text.index(self.findSplitedLine(pdf_text, "Period: "))
                    statement_period = text[details][8:]
                    original_currency = text[details + 1][15:]
                    payee_account_number = self.findSplitedLine(pdf_text, "Payee Account Code").split(" ")[-2]
                    royalty = float(
                        self.findSplitedLine(pdf_text, "Gross Payable Royalties ").replace(",", "").split(" ")[-1])

                    self.alldict["payee_account_number"] = payee_account_number
                    self.alldict["statement_period"] = statement_period
                    self.alldict["original_currency"] = original_currency
                    self.alldict["royalty"] = royalty
                    return self.alldict
                else:
                    pdf_text = pdf_To_text(self.pathFile, [i])
                    text = pdf_text.split("\n")
            else:
                return {"result": "WarnerChappell version not supported"}

        except (ValueError, IndexError):
            return {"result": "WarnerChappell version are in the current Statements but is changed"}

    def KODA(self, pdf_text):
        try:
            if not pdf_text.startswith("My Koda \nMy Koda Account Export,"):
                return {"result": "KODA version not supported"}

            text = pdf_text.split("\n")

            statement_period = text[1].split(",")[-1][1:-1].replace(" ", " ")

            details = text[3]
            payee_account_number = details.split(" ")[0]

            details = text.index(self.findSplitedLine(pdf_text, "Date Description Payments"))
            original_currency = text[details][text[details].index("(") + 1:text[details].index(")")]
            royalty = []
            distribution_date = re.findall("\d+\/\d+\/\d{4}", pdf_text)
            royalty = re.findall(" \d+\.\d+\n", pdf_text)
            for i in range(len(royalty)):
                royalty[i] = royalty[i][1:-1]

            self.alldict["statement_period"] = statement_period
            self.alldict["original_currency"] = original_currency
            self.alldict["payee_account_number"] = payee_account_number
            self.alldict["distribution_date"] = distribution_date
            self.alldict["royalty"] = royalty  # problem with royalty

            return self.alldict
        except (ValueError, IndexError):
            return {"result": "Koda version are in the current Statements but is changed"}

    def findSplitedLine(self, source, text):
        detailsIndex = source.index(text)
        return source[detailsIndex:detailsIndex + source[detailsIndex:].index("\n")]

    def rfindSplitedLine(self, source, text):
        detailsIndex = source.rindex(text)
        return source[detailsIndex:detailsIndex + source[detailsIndex:].index("\n")]
