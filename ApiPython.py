

from subs.PdfIdentification import CheckPdf, PdfIdentifier
from subs.Pdf_To_Text import pdf_To_text
from subs.PdfAudit import *
from subs.WixenParser import WixenParser
from subs.PRSParser import PRSParser
from subs.CMGParser import CMGParser
from subs.sql2xlsx import sql2xlsx
from werkzeug.utils import secure_filename
from subs.names_classifier.model import NamesClassifier
import torch
from transformers import logging

import flask
from flask import request, jsonify
from flask_cors import cross_origin, CORS

app = flask.Flask(__name__)
cors = CORS(app)
app.config["DEBUG"] = True
app.config['UPLOAD_FOLDER'] = "Files"
app.config['MAX_CONTENT_PATH'] = 4194304


class Filesfunc:
    def PdfAudit(self, filepath):
        try:
            return pdfAudit(filepath)
        except FileNotFoundError:
            return {"result": "Error File Not Found"}

    def PdfParse(self, filepath):
        pdf_type = PdfIdentifier(filepath)

        if pdf_type == "PRS":
            parser = PRSParser(pdf_filepath=filepath)
        elif pdf_type == "WIXEN":
            parser = WixenParser(pdf_filepath=filepath)
        elif pdf_type == "CMG":
            parser = CMGParser(pdf_filepath=filepath)
        else:
            print(pdf_type)
            # format wasn't found:
            return {"result": "format wasn't found"}

        parser.parse()
        parser.save_result(os.getcwd() + "\\Files\\xlss\\1.csv")

        return {"Success": flask.url_for("static",filename=os.getcwd()+"\\Files\\xlss\\Thefile.csv",_external=True) }


    def Sql2Xlsx(self, filepath):

        # dbname = request.form.get('dbname')
        # queries = request.form.get('queries')                     =============================here=============================
        # queries = eval(queries)

        return sql2xlsx(dbname=dbname, queries=queries, output_filename=filepath)


@app.route('/upload')
def upload_file():
    return render_template('index.html')


@app.route('/uploader', methods=['GET', 'POST'])
def upload_file1():
    if request.method == 'POST':
        function = request.form.get("thefunction")
        f = request.files["thefile"]
        f.save("Files/" + secure_filename(f.filename))
        f.close()
        funcs = Filesfunc()
        value = getattr(funcs, function)(os.getcwd() + "\\Files\\" + f.filename)
        return jsonify(value)


@app.route('/parse', methods=['POST'])
def parse():
    path_pdf = request.form.get('path_pdf')
    path_csv= request.form.get('path_csv')

    pdf_type = PdfIdentifier(path_pdf)

    print("-------")
    print(path_pdf)
    print(pdf_type)

    if pdf_type == "PRS":
        parser = PRSParser(pdf_filepath=path_pdf)
    elif pdf_type == "WIXEN":
        parser = WixenParser(pdf_filepath=path_pdf)
    elif pdf_type == "CMG":
        parser = CMGParser(pdf_filepath=path_pdf)
    else:
        return jsonify({"Type": "not found"})

    parser.parse()
    parser.save_result(path_csv)
    return jsonify({"Type": pdf_type})


@app.route('/classify_name', methods=['POST'])
def classify():
    name = request.form.get('name')

    return jsonify({"is title": model.classify(name)})


if __name__ == "__main__":

    logging.set_verbosity_error()
    model = NamesClassifier()
    model.load_state_dict(torch.load('./subs/names_classifier/best_model.pth'))

    app.run(port=5100)

