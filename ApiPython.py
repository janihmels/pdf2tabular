import firstformat
import flask
from flask import request, jsonify
from flask_cors import cross_origin

app = flask.Flask(__name__)
app.config["DEBUG"] = True


@app.route('/getpdf', methods=['POST'])
@cross_origin()
def test__():
    pdflist = []
    filepath = request.args.get('url')
    pdf_text = firstformat.pdf_To_text(filepath, pages=[0])
    res = firstformat.CheckPdf(pdf_text, filepath, pdflist)
    if res == "Try 3":
        pdf_text = firstformat.pdf_To_text(filepath, pages=[2])
        res = firstformat.CheckPdf(pdf_text, filepath, pdflist)
    return jsonify({"result: ":res})


if __name__ == "__main__":
    app.run(port=5100)

