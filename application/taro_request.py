import requests
import json
import time


def make_request(url, headers, payload, max_retries=20):
    retries = 0
    while retries < max_retries:
        try:
            # 发送请求并获取响应
            response = requests.request("POST", url, headers=headers, data=payload)
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
