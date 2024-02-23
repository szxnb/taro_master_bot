import threading
import time
from queue import Queue
import requests
import report
import json
import report_pdf


class TaroThread(threading.Thread):
    def __init__(self, question, card, email):
        super(TaroThread, self).__init__()
        self.question = question
        self.card = card
        self.email = email
        self.url = "https://oa.api2d.net/v1/chat/completions"
        self.taro_system_prompt = "你是塔罗牌占卜大师，你会根据我所抽取的三张卡牌来对我的问题进行占卜。你说起话来很神秘但很有亲和力。"

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

    def make_request_overview(self, result_queue):
        content_overview = f"我的问题是: {self.question} 我所抽到的三张卡牌是: {self.card} 请用500字的英文来回答我的问题，每段回答缩进为2个空格"
        result_overview = self.make_request(self.make_payload(content_overview))
        result_queue.put(("overview", result_overview))

    def make_request_love(self, result_queue):
        content_love = f"请结合我抽到的卡牌来给我提一些爱情方面的建议，我抽到的三张卡牌是: {self.card} 请用500字的英文来回答我的问题，每段回答缩进为2个空格"
        result_love = self.make_request(self.make_payload(content_love))
        result_queue.put(("love", result_love))

    def make_request_career(self, result_queue):
        content_career = f"请结合我抽到的卡牌来给我提一些职业方面的建议，我抽到的三张卡牌是: {self.card} 请用500字的英文来回答我的问题，每段回答缩进为2个空格"
        result_career = self.make_request(self.make_payload(content_career))
        result_queue.put(("career", result_career))

    def make_request_finances(self, result_queue):
        content_finances = f"请结合我抽到的卡牌来给我提一些财务方面的建议，我抽到的三张卡牌是: {self.card} 请用500字的英文来回答我的问题，每段回答缩进为2个空格"
        result_finances = self.make_request(self.make_payload(content_finances))
        result_queue.put(("finances", result_finances))

    def run(self):
        # 创建一个队列用于存储每个线程的结果，线程安全
        results_queue = Queue()
        # 创建并启动线程
        thread_overview = threading.Thread(target=self.make_request_overview, args=(results_queue,))
        thread_love = threading.Thread(target=self.make_request_love, args=(results_queue,))
        thread_career = threading.Thread(target=self.make_request_career, args=(results_queue,))
        thread_finances = threading.Thread(target=self.make_request_finances, args=(results_queue,))
        thread_overview.start()
        thread_love.start()
        thread_career.start()
        thread_finances.start()
        # 线程阻塞，等待所有线程完成后继续
        thread_overview.join()
        thread_love.join()
        thread_career.join()
        thread_finances.join()
        # 在这里汇总结果
        results = {}
        while not results_queue.empty():
            result = results_queue.get()
            results[result[0]] = result[1]
        result_overview = results['overview']
        result_love = results['love']
        result_career = results['career']
        result_finances = results['finances']
        print("开始生成报告")
        print(f"1/4 总述:\n{result_overview}\n----------------------------------------------")
        print(f"2/4 爱情方面:\n{result_love}\n----------------------------------------------")
        print(f"3/4 职业方面:\n{result_career}\n----------------------------------------------")
        print(f"4/4 财务方面:\n{result_finances}")
        # 生成报告
        # report.generate_ppt(result_overview, result_love, result_career, result_finances)
        cards = self.card.split(",")
        answers = {
            'result_overview' : result_overview,
            'result_love' : result_love,
            'result_career' : result_career,
            'result_finances' : result_finances
        }
        report_pdf.generate(self.question, cards[0], cards[1], cards[2], answers, self.email)