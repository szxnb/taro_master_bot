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
        <h1>Please check your exclusive tarot card calculation report</h1>
        <p>respected user,</p>
        
        <p>Thank you for purchasing our tarot card reading report. For your convenience, we have attached the tarot card calculation report you purchased to this email. Please check the attachment.</p>
        
        <p>You can also visit our website <a href="http://tarot.pakqoostudio.com/" target="_blank">click here</a> to learn more about tarot.</p>
        
        <p>If you have any questions about the report content or require further clarification, please feel free to contact us. We will be happy to assist you.</p>
        
        <p>Thank you again for choosing our service.</p>
        
        <p>I wish you all the best!</p>
        
        <p>Sincerely,<br>
        Pakqoo Studio<br>
    </div>
</body>
</html>
"""
def send_email(file_locate, recipient_email):
    # 你的邮箱账号和密码
    email = 'pakqoostudio1@hotmail.com'
    password = 'suegaapafsiugteu'

    # 收件人邮箱
    recipient = recipient_email

    # 创建包含邮件内容的消息对象
    msg = MIMEMultipart()
    # 将 HTML 内容附加到邮件中
    msg.attach(MIMEText(html_content, 'html', 'utf-8'))
    msg['From'] = email
    msg['To'] = recipient
    msg['Subject'] = Header('Tarot Card Calculation Report', 'utf-8')

    # 添加pdf附件
    with open(file_locate, 'rb') as file:
        attach_pdf = MIMEApplication(file.read(), _subtype="pdf")

    attach_pdf.add_header('Content-Disposition', 'attachment', filename= file_locate)
    msg.attach(attach_pdf)

    # 连接到SMTP服务器并发送邮件
    server = smtplib.SMTP('smtp-mail.outlook.com', 587)  # 你的SMTP服务器地址和端口号
    server.starttls()  # 开启安全传输模式
    server.login(email, password)  # 登录邮箱
    server.sendmail(email, recipient, msg.as_string())  # 发送邮件
    server.quit()  # 退出SMTP会话
    print(file_locate + "已发送")