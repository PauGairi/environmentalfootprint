from flask import Flask

app = Flask(__name__)


@app.route("/")
def home():
    return "<h2>This is Python working Online!! We get it Gerard!</h2>"
