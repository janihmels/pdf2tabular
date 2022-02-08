from subs.PdfIdentification import CheckPdf, PdfIdentifier
from subs.Pdf_To_Text import pdf_To_text
from subs.PdfAdult import *
from subs.WixenParser import WixenParser
from subs.PRSParser import PRSParser
from subs.CMGParser import CMGParser

import flask
from flask import request, jsonify
from flask_cors import cross_origin

app = flask.Flask(__name__)
app.config["DEBUG"] = True


@app.route('/pdfIdentification', methods=['POST'])
@cross_origin()
def PdfPublisher():
    try:
        pdflist = []
        filename = request.form.get('filename')
        filepath = request.form.get('path')
        fullfile = str(filepath) + "/" + str(filename)
        pdf_text = pdf_To_text(fullfile, pages=[0])
        res = CheckPdf(pdf_text, fullfile, pdflist)
        if res == "Try 3":
            pdf_text = pdf_To_text(fullfile, pages=[2])
            res = CheckPdf(pdf_text, fullfile, pdflist)
        return jsonify({"result: ": res})
    except FileNotFoundError:
        return jsonify({"result: ": "Error File Not Found!"})


@app.route('/pdfAudit', methods=['POST'])
@cross_origin()
def PdfAudit():
    try:
        filename = request.form.get('filename')
        filepath = request.form.get('path')
        format = request.form.get('format')
        page = request.form.get('page')
        fullfile = str(filepath) + "/" + str(filename)

        if page is None:
            page = 0
        else:
            page = int(page)

        return jsonify(pdfAudit(fullfile,format,page))
    except FileNotFoundError:
        return jsonify({"result: ": "Error File Not Found"})


@app.route('/pdfParse', methods=['POST'])
@cross_origin()
def PdfParse():
    src_filename = request.form.get('src_filename')
    src_filepath = request.form.get('src_filepath')

    dst_filename = request.form.get('dst_filename')
    dst_filepath = request.form.get('dst_filepath')

    src_fullfile = str(src_filepath) + '/' + str(src_filename)
    dst_fullfile = str(dst_filepath) + '/' + str(dst_filename)

    pdf_type = PdfIdentifier(src_fullfile)

    if pdf_type == "PRS":
        parser = PRSParser(pdf_filepath=src_fullfile)
    elif pdf_type == "Wixen":
        parser = WixenParser(pdf_filepath=src_fullfile)
    elif pdf_type == "CMG":
        parser = CMGParser(pdf_filepath=src_fullfile)
    else:
        # format wasn't found:
        return jsonify({"result: ": "Error File Not Found"})

    parser.parse()
    parser.save_result(dst_fullfile)

    return jsonify({"result": "file successfully extracted to {0}".format(dst_fullfile)})


if __name__ == "__main__":
    app.run(port=5100)
