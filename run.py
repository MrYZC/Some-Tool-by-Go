from flask import Flask, render_template, jsonify
from random import * 
from flask_cors import CORS
import requests
'''
    flask web框架
    random函数
    flask_cors 跨域
    requests HTTP请求包
    app是Flask实例
    __name__意思是传入python的__name__变量是程序的名字
'''
app = Flask(__name__,
            static_folder = "./dist/static",
            template_folder = "./dist")
cors = CORS(app, resources={r"/api/*": {"origins": "*"}})

@app.route('/api/random')
def random_number():
    response = {
        'randomNumber': randint(1, 5)
    }
    return jsonify(response)

@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def catch_all(path):
    if app.debug:
        return requests.get('http://localhost:8080/{}'.format(path)).text
    return render_template("index.html")

# from flask import Flask
# form requests import *
# app = Flask(__name__)

# @app.route('/', defaults={'path': ''})
# @app.route('/<path:path>')
# def catch_all(path):
#     if app.debug:
#         return requests.get('http://localhost:8080/{}'.format(path)).text
#     return render_template("index.html")

if __name__ == '__main__':
    app.debug = False
    app.run(host='localhost', port=5000)