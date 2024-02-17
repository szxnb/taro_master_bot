from flask import Flask
from main import MainProcess

app = Flask(__name__)

@app.route("/")
def hello_world():
    return "<p>Hello, World!</p>"

@app.get('/make_report/<user_question>/<card1>/<card2>/<card3>/<user_email>')
def make_report(user_question, card1, card2, card3, user_email):
    print("收到参数:" + user_question + "/" + card1 + "/" + card2 + "/" + card3 + "/" + user_email)
    question = user_question
    card = f"{card1},{card2},{card3}"
    email = user_email

    main_process_instance = MainProcess()
    main_process_instance.start_process(question, card, email)
    return f"<p>生成完成，报告已发送至{email}邮箱，请查收</p>"

# /make_report/我什么时候能结婚/战车/力量/月亮/3500466989@qq.com