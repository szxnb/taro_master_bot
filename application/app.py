from flask import Flask
from taro import TaroThread
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # 允许所有域进行跨域请求


@app.route("/")
def hello_world():
    return "<p>Hello, World!</p>"


@app.get('/make_report/<question>/<card1>/<card2>/<card3>/<email>')
def make_report(question, card1, card2, card3, email):
    print(f"收到参数:{question}/{card1}/{card2}/{card3}/{email}")
    card = f"{card1},{card2},{card3}"
    taro_thread = TaroThread(question, card, email)
    taro_thread.start()
    return f"<p>生成完成，报告已发送至{email}邮箱，请查收</p>"


# /make_report/我什么时候能结婚/Sun/Moon/Star/3500466989@qq.com
if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
