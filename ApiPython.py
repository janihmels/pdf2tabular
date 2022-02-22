from subs.PdfIdentification import CheckPdf, PdfIdentifier
from subs.Pdf_To_Text import pdf_To_text
from subs.PdfAudit import *
from subs.WixenParser import WixenParser
from subs.PRSParser import PRSParser
from subs.CMGParser import CMGParser
from subs.sql2xlsx import sql2xlsx
from werkzeug.utils import secure_filename

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
            parser = PRSParser(pdf_filepath=src_fullfile)
        elif pdf_type == "Wixen":
            parser = WixenParser(pdf_filepath=src_fullfile)
        elif pdf_type == "CMG":
            parser = CMGParser(pdf_filepath=src_fullfile)
        else:
            # format wasn't found:
            return {"result": "format wasn't found"}

        parser.parse()
        parser.save_result(dst_fullfile)

        return {"result": "file successfully extracted to {0}".format(dst_fullfile)}

    def Sql2Xlsx(self, filepath):

        # dbname = request.form.get('dbname')
        # queries = request.form.get('queries')                     =============================here=============================
        # queries = eval(queries)

        return sql2xlsx(dbname=dbname, queries=queries, output_filename=filepath)


@app.route('/upload')
def upload_file():
    return render_template('main.html')


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


if __name__ == "__main__":
    app.run(port=5100)
