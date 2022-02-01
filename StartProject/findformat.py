
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


pdf_text = pdf_To_text("data.pdf", pages=[1])

print(pdf_text)