
import re
import os

rxcountpages = re.compile(r"/Type\s*/Page([^s]|$)", re.MULTILINE|re.DOTALL)

def count_pages(filename):
    data = open(filename, "rb").read()
    return len(rxcountpages.findall(data))

if __name__=="__main__":
    filename = '../exempleParsing/PRS/2016/Essex David_070172307_2016041_052_089780_26722.PDF'
    print(count_pages(filename))