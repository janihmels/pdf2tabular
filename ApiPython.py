

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
import json
from subs.source_classifier.model import SourceClassifier
import pickle
import timeit
import flask
from flask import request, jsonify
from flask_cors import cross_origin, CORS
import pandas as pd
import numpy as np

from ouputsApi.main import *

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
    path_csv = '.'.join(path_pdf.split('.')[:-1]) + '.csv'

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



@app.route('/classify_names', methods=['POST'])
def classify_name():
    names = request.form.get('names')
    names = json.loads(s=names)

    return jsonify({name: title_classifier.classify(name) for name in names})


@app.route('/classify_sources', methods=['POST'])
def classify_source():
    names = request.form.get('sources')
    names = json.loads(s=names)

    return jsonify({name: source_classifier.classify(name) for name in names})


def toDict(df):
    return df.replace(np.nan ,None).to_dict('list')

@app.route('/PublishCatalog', methods=['POST'])
def home_PublishCatalog():
    start = timeit.default_timer()
    projectid = request.form.get('projectid')
    pathParquet = request.form.get('path_to_parquet')
    pathResult = request.form.get('path_to_result')

    parquet_file = pd.read_parquet(pathParquet+projectid + ".gzip", engine='pyarrow')
    print("PublishCatalog start")

    catalogDict = {}
    catalogDict["Catalog_Details"] = [defualtDetails(parquet_file).replace(np.nan, None).to_dict('list')]
    catalogDict["songXrevXhalf"] = [SimpleExtract("Song_Name_9LC" ,parquet_file).replace(np.nan, None).to_dict('list')]  # 1 sec
    catalogDict["incomeXrevXhalf"] = [SimpleExtract("Normalized_Income_Type_9LC", parquet_file).replace(np.nan, None).to_dict('list')]# 1 sec
    catalogDict["sourceXrevXhalf"] = [SimpleExtract("Normalized_Source_9LC", parquet_file).replace(np.nan, None).to_dict('list')]
    catalogDict["artistXrevXhalf"] = [artistxrevxhalf(parquet_file).replace(np.nan, None).to_dict('list')]
    catalogDict["payorXincomeXtypeXrevXhalf"] = list(map(toDict, payorXincomeXtypeXrevXhalf(parquet_file)))
    catalogDict["payorXsongXrevXhalf"] = list(map(toDict, payorXsongXrevXhalf(parquet_file)))
    catalogDict["payorXsourceXrevXhalf"] = list(map(toDict, payorXsourceXrevXhalf(parquet_file)))

    filePath = pathResult+projectid
    if os.path.exists(filePath):
        os.remove(filePath)

    file = open(filePath, "w",encoding="utf-8")
    file.write(str(catalogDict))

    stop = timeit.default_timer()
    print(stop - start, "seconds")
    print("PublishCatalog end")

    return jsonify()

@app.route('/PullTable', methods=['POST'])
def PullTable():
    projectid = request.form.get('projectid')
    pathResult = request.form.get('path_to_result')
    filePath = pathResult+projectid
    with open(filePath) as json_file:
        data = json.load(json_file)
    return jsonify(data)


if __name__ == "__main__":
    logging.set_verbosity_error()
    title_classifier = NamesClassifier()
    title_classifier.load_state_dict(torch.load('./subs/names_classifier/best_model.pth', map_location=torch.device('cpu')))

    label_to_source_map = pickle.load(open('subs/source_classifier/label_to_name_mapper.pkl', 'rb'))
    source_classifier = SourceClassifier(num_cls=len(label_to_source_map.values()), label_to_name=label_to_source_map)
    source_classifier.load_state_dict(torch.load('./subs/source_classifier/best_model.pth', map_location=torch.device('cpu')))

    app.run(port=5100)


