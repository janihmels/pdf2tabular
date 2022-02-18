from subs.PdfIdentification import CheckPdf, PdfIdentifier
from subs.Pdf_To_Text import pdf_To_text
from subs.PdfAudit import *
from subs.WixenParser import WixenParser
from subs.PRSParser import PRSParser
from subs.CMGParser import CMGParser
from subs.sql2xlsx import sql2xlsx

import flask
from flask import request, jsonify
from flask_cors import cross_origin

app = flask.Flask(__name__)
app.config["DEBUG"] = True


@app.route('/pdfAudit', methods=['POST'])
@cross_origin()
def PdfAudit():
    try:
        filename = request.form.get('filename')
        filepath = request.form.get('path')
        page = request.form.get('page')
        fullfile = str(filepath) + "/" + str(filename)

        if page is None:
            page = 0
        else:
            page = int(page)

        return jsonify(pdfAudit(fullfile, page))
    except FileNotFoundError:
        return jsonify({"result": "Error File Not Found"})


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
        return jsonify({"result": "Error File Not Found"})

    parser.parse()
    parser.save_result(dst_fullfile)

    return jsonify({"result": "file successfully extracted to {0}".format(dst_fullfile)})


@app.route('/sql2xlsx', methods=['POST'])
@cross_origin()
def SQL2XLSX():

    dbname = request.form.get('dbname')
    queries = request.form.get('queries')
    queries = eval(queries)

    dst_filename = request.form.get('dst_filename')
    dst_filepath = request.form.get('dst_filepath')
    dst_fullfile = str(dst_filepath) + '/' + str(dst_filename)

    sql2xlsx(dbname=dbname, queries=queries, output_filename=dst_fullfile)


if __name__ == "__main__":
    app.run(port=5100)
