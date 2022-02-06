import subs.PdfIdentification
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
        pdf_text = PdfIdentification.pdf_To_text(fullfile, pages=[0])
        res = firstformat.CheckPdf(pdf_text, fullfile, pdflist)
        if res == "Try 3":
            pdf_text = PdfIdentification.pdf_To_text(fullfile, pages=[2])
            res = PdfIdentification.CheckPdf(pdf_text, fullfile, pdflist)
        return jsonify({"result: ": res})
    except FileNotFoundError:
        return jsonify({"result: ": "Error File Not Found"})


if __name__ == "__main__":
    app.run(port=5100)
