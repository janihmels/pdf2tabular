@ -1,3 +1,5 @@
import os

from subs.PdfIdentification import CheckPdf, PdfIdentifier
from subs.Pdf_To_Text import pdf_To_text
from subs.PdfAudit import *
@ -21,9 +23,9 @@ app.config['MAX_CONTENT_PATH'] = 4194304
class Filesfunc:
    def PdfAudit(self, filepath):
        try:
            return pdfAudit(filepath)
            return jsonify(pdfAudit(filepath))
        except FileNotFoundError:
            return {"result": "Error File Not Found"}
            return jsonify({"result": "Error File Not Found"})

    def PdfParse(self, filepath):
        pdf_type = PdfIdentifier(filepath)
@ -37,12 +39,12 @@ class Filesfunc:
        else:
            print(pdf_type)
            # format wasn't found:
            return {"result": "format wasn't found"}
            return jsonify({"result": "format wasn't found"})

        parser.parse()
        parser.save_result(os.getcwd() + "\\Files\\xlss\\1.csv")

        return {"Success": "file successfully extracted to {0}".format(filepath)}
        parser.save_result(os.getcwd() + "\\Files\\xlss\\Thefile.csv")
        #flask.send_file(,as_attachment=False)
        return flask.url_for("static",filename=os.getcwd()+"\\Files\\xlss\\Thefile.csv",_external=True)

    def Sql2Xlsx(self, filepath):

@ -63,12 +65,12 @@ def upload_file1():
    if request.method == 'POST':
        function = request.form.get("thefunction")
        f = request.files["thefile"]
        f.save("Files/" + secure_filename(f.filename))
        f.save("Files/" + secure_filename("Thefile.pdf"))
        f.close()
        funcs = Filesfunc()
        value = getattr(funcs, function)(os.getcwd() + "\\Files\\" + f.filename)

        return jsonify(value)
        value = getattr(funcs, function)(os.getcwd() + "\\Files\\Thefile.pdf")
        print(value)
        return value

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

if __name__ == "__main__":
