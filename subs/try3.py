from subs.Pdf_To_Text import pdf_To_text
import os
import re


def pdfAudit(pathFile, format, page):
    alldict = {"result": "Format Doesn't exist"}
    pdf_text = pdf_To_text(pathFile, pages=[page])

    format = format.upper()

    if format == "BMI":
        alldict = BMI(pdf_text)
    elif format == "ASCAP":
        alldict = ASCAP(pdf_text)
    elif format == "PRS":
        alldict = PRS(pdf_text)
    elif format == "SoundExchange":
        alldict = SoundExchange(pdf_text)

    return alldict


def BMI(pdf_text):
    try:
        alldict = {}
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

        alldict["Account_No"] = Account_No
        alldict["U.S. Performance Period"] = US_Period
        alldict["payee_account_number"] = account_number
        alldict["distribution_date1"] = distribution_date
        alldict["distribution_date2"] = "Year(" + distribution_date[:4] + "), Month(" + distribution_date[
                                                                                        5:7] + "), Day(" + distribution_date[
                                                                                                           8:10] + ")"
        alldict["total_amount"] = total_amount
        return alldict

    except ValueError:
        return {"result": "PRS version are in the current Statements but is changed"}


def ASCAP(pdf_text):
    try:
        alldict = {}
        text = pdf_text.split("\n")
        if "Writer International Distribution For:\nMember Name:" not in pdf_text:
            if "Writer Domestic Distribution For:\nMember Name:" not in pdf_text:
                return {"result": "ASCAP version not supported"}
            else:
                alldict["Version"] = "Domestic"
                index = text.index("Writer Domestic Distribution For:")
        else:
            alldict["Version"] = "International"
            index = text.index("Writer International Distribution For:")

        payee_account_number = text[index + 2][11:]
        statement_period = text[index + 4]
        royalty = float(text[index + 6][15:].replace(",", ""))

        alldict["payee_account_number"] = payee_account_number
        alldict["statement_period"] = statement_period
        alldict["royalty"] = royalty

        datecheck = re.findall("Check #: \d+ Date : \d{2}-\d{2}-\d{4}\n", pdf_text)
        if datecheck:
            distribution_date = datecheck[0][datecheck[0].index("Date :") + 7:-1]
            alldict["distribution_date1: "] = distribution_date
            alldict["distribution_date2: "] = "Year(" + distribution_date[-4:] + "), " + "Month(" + distribution_date[
                                                                                                    :2] + "), " + "Day(" + distribution_date[
                                                                                                                           3:5] + ")"

        return alldict

    except ValueError:
        return {"result": "PRS version are in the current Statements but is changed"}


def PRS(pdf_text):
    alldict = {}
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

        alldict["payee_account_number"] = payee_account_number
        alldict["statement_period"] = statement_period
        alldict["original_currency"] = original_currency
        alldict["royalty"] = royalty

        return alldict

    except ValueError:
        return {"result": "PRS version are in the current Statements but is changed"}


def SoundExchange(pdf_text):
    alldict = {}
    try:
        text = pdf_text.split("\n")
        if text[0] != re.findall(".+Featured Artist", pdf_text)[0]:
            return {"result": "SoundExchange version not supported "}

        statement_period = text[0][1:text[0].index("Featured Artist") - 1]
        payee_id = pdf_text[pdf_text.index("SoundExchange  Payee ID") + 23:]
        royalty = float(text[text.index("Featured Artist Payment") + 2][1:].replace(",", ""))

        alldict["statement_period"] = statement_period
        alldict["payee_id"] = payee_id
        alldict["royalty"] = royalty

        return alldict

    except ValueError:
        return {"result": "PRS version not supported"}
