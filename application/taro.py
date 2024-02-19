import threading
import time

import requests


import report
import json


class TaroThread(threading.Thread):
    def __init__(self, question, card, email):
        super(TaroThread, self).__init__()
        self.question = question
        self.card = card
        self.email = email
        self.url = "https://oa.api2d.net/v1/chat/completions"
        self.taro_system_prompt = "你是塔罗牌占卜大师，你会根据我所抽取的三张卡牌来对我的问题进行占卜。在回答问题的过程中还会带一些幽默。"

    def make_payload(self, content):
        return json.dumps({
            "model": "gpt-3.5-turbo",
            "messages": [
                {"role": "system", "content": self.taro_system_prompt},
                {"role": "user", "content": content},
            ],
            "safe_mode": False,
            "max_tokens": 2000
        })

    def make_request(self, payload):
        headers = {
            'Authorization': 'Bearer fk200296-DZuQfeFnQLL40v09VWbT8VF58i6T1E4P',
            'User-Agent': 'Apifox/1.0.0 (https://apifox.com)',
            'Content-Type': 'application/json'
        }
        retries = 0
        max_retries = 20
        while retries < max_retries:
            try:
                # 发送请求并获取响应
                response = requests.request("POST", self.url, headers=headers, data=payload)
                # 检查响应状态码是否为200
                if response.status_code == 200:
                    # 解析JSON响应内容并返回
                    data = response.json()
                    content = data['choices'][0]['message']['content']
                    return content
                else:
                    # 输出错误信息
                    print("Error:", response.status_code)
            except requests.exceptions.RequestException as e:
                # 输出请求异常信息
                print("Request Exception:", e)
            # 如果请求失败，则等待一段时间后重试
            retries += 1
            print(f"Retrying... {retries}/{max_retries}")
            time.sleep(1)

        # 如果达到最大重试次数仍然失败，则输出错误信息并返回空值
        print("Max retries reached. Request failed.")
        return None

    def run(self):
        content_overview = f"我的问题是: {self.question} 我所抽到的三张卡牌是: {self.card} 请用300字的英文来回答我的问题"
        content_love = f"请你帮我预测一下我未来一个月到三个月的爱情状况，我抽到的三张卡牌是: {self.card} 请用300字的英文来回答我的问题"
        content_career = f"请你帮我预测一下我未来一个月到三个月的职业状况，我抽到的三张卡牌是: {self.card} 请用300字的英文来回答我的问题"
        content_finances = f"请你帮我预测一下我未来一个月到三个月的财务状况，我抽到的三张卡牌是: {self.card} 请用300字的英文来回答我的问题"

        result_overview = self.make_request(self.make_payload(content_overview))
        result_love = self.make_request(self.make_payload(content_love))
        result_career = self.make_request(self.make_payload(content_career))
        result_finances = self.make_request(self.make_payload(content_finances))

        print("开始生成报告")
        print(f"1/4 总述:\n{result_overview}\n----------------------------------------------")
        print(f"2/4 爱情方面:\n{result_love}\n----------------------------------------------")
        print(f"3/4 职业方面:\n{result_career}\n----------------------------------------------")
        print(f"4/4 财务方面:\n{result_finances}")

        # 生成报告
        report.generate_ppt(result_overview, result_love, result_career, result_finances)
