
from flask import Flask, render_template

app = Flask(__name__)

@app.route('/')
def home():
    return "Redutron ERP - Render OK"

if __name__ == '__main__':
    app.run()
