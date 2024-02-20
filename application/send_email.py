import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header

# 创建带有样式的 HTML 内容
html_content = """
<!DOCTYPE html>
<html>
<head>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f9f9f9;
            margin: 0;
            padding: 0;
        }
        .container {
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
        }
        h1 {
            color: #333333;
        }
        p {
            color: #666666;
        }
        a {
            color: #007bff;
            text-decoration: none;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>请查收您的专属塔罗牌测算报告</h1>
        <p>尊敬的用户，</p>
        
        <p>感谢您购买我们的塔罗牌测算报告。为了您的方便，我们已经在此邮件中附上了您购买的塔罗牌测算报告，请查收附件。</p>
        
        <p>您也可以<a href="https://taromaster.streamlit.app/" target="_blank">点击这里</a>访问我们的网站了解更多有关塔罗牌的信息。</p>
        
        <p>如果您有任何关于报告内容的问题或需要进一步的解释，请随时与我们联系。我们将竭诚为您提供帮助。</p>
        
        <p>再次感谢您选择我们的服务。</p>
        
        <p>祝您一切顺利！</p>
        
        <p>诚挚地，<br>
        Pakqoo Studio<br>
    </div>
</body>
</html>
"""

# 你的邮箱账号和密码
email = 'pakqoostudio1@hotmail.com'
password = 'suegaapafsiugteu'

# 收件人邮箱
recipient = '3500466989@qq.com'
# 1191976408@qq.com

# 创建包含邮件内容的消息对象
msg = MIMEMultipart()
# 将 HTML 内容附加到邮件中
msg.attach(MIMEText(html_content, 'html', 'utf-8'))
msg['From'] = email
msg['To'] = recipient
msg['Subject'] = Header('塔罗牌测算报告', 'utf-8')

# 添加PDF附件
with open('../reports/example.pdf', 'rb') as file:
    attach_pdf = MIMEApplication(file.read(), _subtype="pdf")

attach_pdf.add_header('Content-Disposition', 'attachment', filename='example.pdf')
msg.attach(attach_pdf)

# 连接到SMTP服务器并发送邮件
server = smtplib.SMTP('smtp-mail.outlook.com', 587)  # 你的SMTP服务器地址和端口号
server.starttls()  # 开启安全传输模式
server.login(email, password)  # 登录邮箱
server.sendmail(email, recipient, msg.as_string())  # 发送邮件
server.quit()  # 退出SMTP会话