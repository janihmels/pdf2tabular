from subs.Pdf_To_Text import pdf_To_text
import os
import re


def pdfAudit(pathFile, format, page):
    pdf_text = pdf_To_text(pathFile, pages=[page])
    dicts = Formats()
    return getattr(dicts,format.upper())(pdf_text)


class Formats:
    def __init__(self):
        self.alldict = {}

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
        except ValueError:
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
            if not pdf_text.startswith("Notice of Payment"):
                return {"result": "PRS version not supported try page = 1 or 2"}

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
            if text[0] != re.findall(".+Featured Artist", pdf_text)[0]:
                return{"result": "SoundExchange version not supported "}

            statement_period = text[0][1:text[0].index("Featured Artist") - 1]
            idIndex = pdf_text.index("SoundExchange  Payee ID")
            payee_id = pdf_text[idIndex + 23:pdf_text[idIndex:].index("\n")+idIndex]

            royalty = float(text[text.index("Featured Artist Payment") + 2][1:].replace(",", ""))

            self.alldict["statement_period"] = statement_period
            self.alldict["payee_id"] = payee_id
            self.alldict["royalty"] = royalty

            return self.alldict
        except (ValueError, IndexError):
            return {"result": "SoundExchange version not supported"}

    def BMG(self,pdf_text):
        try:
            text = pdf_text.split("\n")
            if text[0].startswith("BMG Rights Management (UK) Ltd."):
                details = re.findall("Date: \d{2}\/\d{2}\/\d{4}\n\nIn Account with : \(\d+\)",pdf_text)[0]
                distribution_date1 = details[6:]
                distribution_date2 = "Year("+distribution_date1[-4:]+"), Month("+distribution_date1[3:5]+"), "+"Day("+distribution_date1[:2]+")"
                payee_account_number = details[36:]
                statement_period = text[text.index("Date: "+distribution_date1)+4][17:]
                original_currency = text[text.index("¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬")+1][-3]
                royalty = text[pdf_text.index("¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦")-2]
            elif text[0].startswith("Summary Statement"):
                pass
            else:
                return{"result": "BMG version not supported "}

        except (ValueError, IndexError):
            return {"result": "SoundExchange version not supported"}
