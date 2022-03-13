from subs.PdfIdentification import CheckPdf, PdfIdentifier
from subs.Pdf_To_Text import pdf_To_text
from subs.PdfAudit import *
from subs.WixenParser import WixenParser
from subs.PRSParser import PRSParser
from subs.CMGParser import CMGParser
from subs.sql2xlsx import sql2xlsx
from werkzeug.utils import secure_filename
from zipfile import *
import io
import flask
import  time
from flask import request, jsonify
from flask_cors import cross_origin, CORS

app = flask.Flask(__name__, static_url_path='')
cors = CORS(app)
app.config["DEBUG"] = True
app.config['UPLOAD_FOLDER'] = "Files"
app.config['MAX_CONTENT_PATH'] = 4194304


class Filesfunc:
    def PdfAudit(self, filepath):
        try:
            return pdfAudit(filepath)
        except FileNotFoundError:
            return "None"

    def PdfParse(self, filepath):
        pdf_type = PdfIdentifier(filepath)

        if pdf_type == "PRS":
            parser = PRSParser(pdf_filepath=filepath)
        elif pdf_type == "WIXEN":
            parser = WixenParser(pdf_filepath=filepath)
        elif pdf_type == "CMG":
            parser = CMGParser(pdf_filepath=filepath)
        else:
            return "None"

        parser.parse()

        return parser

    def Sql2Xlsx(self, filepath):

        # dbname = request.form.get('dbname')
        # queries = request.form.get('queries')                     =============================here=============================
        # queries = eval(queries)

        return jsonify(sql2xlsx(dbname=dbname, queries=queries, output_filename=filepath))


@app.route('/uploader', methods=['GET', 'POST'])
def upload_file1():
    if request.method == 'POST':
        function = request.form.get("thefunction")
        fileslist = request.files.getlist("thefiles")
        funcs = Filesfunc()
        value = []
        with ZipFile("Files\\zip\\myFile.zip", 'w', ZIP_DEFLATED) as zip:
            for fileNum in range(len(fileslist)):
                fileslist[fileNum].save("Files/pdf/Thefile.pdf")
                fileslist[fileNum].close()
                value.append(getattr(funcs, function)(os.getcwd() + "\\Files\\pdf\\Thefile.pdf"))
                if function == "PdfParse":
                    value[-1].save_result(os.getcwd() + "\\Files\\csv\\Thefile"+str(fileNum)+".csv")
                    zip.write(os.getcwd() + "\\Files\\csv\\Thefile"+str(fileNum)+".csv","Thefile"+str(fileNum)+".csv")
            zip.close()

    if function == "PdfParse":
        path = os.getcwd() + "\\Files\\zip"
        zipname = "myFile.zip"
        return flask.send_from_directory(path, zipname)

    '''
        print("here")
        if function == "PdfParse":
            path = os.getcwd() + "\\Files\\zip"
            zipname="myFile.zip"
            FILEPATH = path+"\\"+zipname
            fileobj = io.BytesIO()
            with ZipFile(fileobj, 'w') as zip_file:
                zip_info = ZipInfo(FILEPATH)
                zip_info.date_time = time.localtime(time.time())[:6]
                zip_info.compress_type = ZIP_DEFLATED
                with open(FILEPATH, 'rb') as fd:
                    zip_file.writestr(zip_info, fd.read())
            fileobj.seek(0)
            response = flask.make_response(fileobj.read())
            response.headers.set('Content-Type', 'zip')
            response.headers.set('Content-Disposition', 'attachment', filename='%s.zip' % os.path.basename(FILEPATH))
            return response
        else:
            return jsonify(value)
            '''
    '''
    if function == "PdfParse":
        import base64
        path = os.getcwd() + "\\Files\\zip"
        zipname = "myFile.zip"
        FILEPATH = path + "\\" + zipname
        file = open(FILEPATH,"rb")
        message = file.read()
        base64_bytes = base64.b64encode(message)
        print(base64_bytes)
        return base64_bytes
    '''


@app.route('/parse', methods=['POST'])
def parse():
    path_pdf = request.form.get('path_pdf')
    path_csv = request.form.get('path_csv')

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


if __name__ == "__main__":
    app.run(port=5100)

