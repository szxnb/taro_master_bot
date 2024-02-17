import taro_request
import generate_report
import json

url = "https://oa.api2d.net/v1/chat/completions"
taro_system_prompt = "你是塔罗牌占卜大师，你会根据我所抽取的三张卡牌来对我的问题进行占卜。在回答问题的过程中还会带一些幽默。"
taro_user_question = "我和我男朋友还有可能吗？"
taro_user_card = "女教主,高塔,战车"
taro_user_content_overview = "我的问题是:" + taro_user_question +"我所抽到的三张卡牌是:" + taro_user_card + "请用300字的英文来回答我的问题"
taro_user_content_love = "请你帮我预测一下我未来一个月到三个月的爱情状况，我抽到的三张卡牌是:" + taro_user_card + "请用300字的英文来回答我的问题"
taro_user_content_career = "请你帮我预测一下我未来一个月到三个月的职业状况，我抽到的三张卡牌是:" + taro_user_card + "请用300字的英文来回答我的问题"
taro_user_content_finances  = "请你帮我预测一下我未来一个月到三个月的财务状况，我抽到的三张卡牌是:" + taro_user_card + "请用300字的英文来回答我的问题"
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

# 测算结果
result_overview = taro_request.make_request(url, headers, payload_overview)
result_love = taro_request.make_request(url, headers, payload_love)
result_career = taro_request.make_request(url, headers, payload_career)
result__finances = taro_request.make_request(url, headers, payload_finances)

print("开始生成报告")
print("1/4 总述:\n"  + result_overview)
print("----------------------------------------------")
print("2/4 爱情方面:\n" + result_love)
print("----------------------------------------------")
print("3/4 职业方面:\n" + result_career)
print("----------------------------------------------")
print("4/4 财务方面:\n" + result__finances)

# 生成报告
generate_report.generate_report_ppt(result_overview, result_love, result_career, result__finances)

# 发送邮件



