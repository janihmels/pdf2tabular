from Pdf_To_Text import pdf_To_text
import re

string = '--------------'
print(re.match('[-]+', string).span() == (0, len(string)))

