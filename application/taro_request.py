import requests
import json
import time

url = "https://oa.api2d.net/v1/chat/completions"
taro_system_prompt = "你是塔罗牌占卜大师，你会根据我所抽取的三张卡牌来对我的问题进行占卜。在回答问题的过程中还会带一些幽默。"
taro_user_question = "我和我男朋友还有可能吗？"
taro_user_card = "女教主,高塔,战车"
taro_user_content_overview = "我的问题是:" + taro_user_question +"我所抽到的三张卡牌是:" + taro_user_card + "请用1000字来回答我的问题"
taro_user_content_love = "请你帮我预测一下我未来一个月到三个月的爱情状况，我抽到的三张卡牌是:" + taro_user_card + "请用1000字来回答我的问题"
taro_user_content_career = "请你帮我预测一下我未来一个月到三个月的职业状况，我抽到的三张卡牌是:" + taro_user_card + "请用1000字来回答我的问题"
taro_user_content_finances  = "请你帮我预测一下我未来一个月到三个月的财务状况，我抽到的三张卡牌是:" + taro_user_card + "请用1000字来回答我的问题"
headers = {
   'Authorization': 'Bearer fk200296-DZuQfeFnQLL40v09VWbT8VF58i6T1E4P',
   'User-Agent': 'Apifox/1.0.0 (https://apifox.com)',
   'Content-Type': 'application/json'
}

# 总述
payload_overview = json.dumps({
   "model": "gpt-3.5-turbo",
   "messages": [
      {"role": "system", "content": taro_system_prompt},
      {"role": "user", "content": taro_user_content_overview},
   ],
   "safe_mode": False,
   "max_tokens": 2000
})

# 爱情方面
payload_love = json.dumps({
   "model": "gpt-3.5-turbo",
   "messages": [
      {"role": "system", "content": taro_system_prompt},
      {"role": "user", "content": taro_user_content_love},
   ],
   "safe_mode": False,
   "max_tokens": 2000
})

# 职业方面
payload_career = json.dumps({
   "model": "gpt-3.5-turbo",
   "messages": [
      {"role": "system", "content": taro_system_prompt},
      {"role": "user", "content": taro_user_content_career},
   ],
   "safe_mode": False,
   "max_tokens": 2000
})

# 财务方面
payload_finances = json.dumps({
   "model": "gpt-3.5-turbo",
   "messages": [
      {"role": "system", "content": taro_system_prompt},
      {"role": "user", "content": taro_user_content_finances},
   ],
   "safe_mode": False,
   "max_tokens": 2000
})



def make_request(url, headers, payload, max_retries=20):
    retries = 0
    while retries < max_retries:
        try:
            # 发送请求并获取响应
            response = requests.request("POST", url, headers = headers, data=payload)
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

# 调用方法
print("总述:\n"  + make_request(url, headers, payload))
print("----------------------------------------------")
print("爱情方面:\n" + make_request(url, headers, payload_love))
print("----------------------------------------------")
print("职业方面:\n" + make_request(url, headers, payload_career))
print("----------------------------------------------")
print("财务方面:\n" + make_request(url, headers, payload_finances))
