from flask import Flask, render_template, request
from werkzeug.utils import secure_filename

app = Flask(__name__)


if __name__ == '__main__':
    app.run(debug=True)