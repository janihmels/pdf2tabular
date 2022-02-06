from subs.Pdf_To_Text import pdf_To_text
import os
import re

def pdfAdultBMI(pdf_text):
    alldict = {}
    text = pdf_text.split("\n")
    if not ("Questions About Your Statement? Call:\nBMI's Next Distribution Will Occur During: Moving? Visit bmi.com to change your address" in pdf_text and text[7].startswith("Royalty StatementBMI")):
        return {"result" :"BMI version not supported"}
    royalty = re.findall("Description U.S. Admin Services Total\n\nInternational\n\nCurrent Earnings [$].*[$]", pdf_text)[0]
    endTotal = royalty.rindex("$")
    startTotal = royalty[:endTotal].rindex("$")
    royalty = float(royalty[startTotal+1:endTotal].replace(",",""))
    alldict["royalty"] = royalty

    details = re.findall("Account\nDateNumberNumber\n\d+-\d\d-\d\d\n.*\n\n", pdf_text)[0].split("\n")
    account_number = details[2][:9]
    distribution_date = details[2][17:]
    total_amount = float(details[3][details[3].rindex("$")+1:].replace(",",""))

    alldict["account_number"] = account_number
    alldict["distribution_date1"] = distribution_date
    alldict["distribution_date2"] = "Year("+distribution_date[:4]+"), Month("+distribution_date[5:7]+"), Day("+distribution_date[8:10]+")"
    alldict["total_amount"] = total_amount
    return alldict
