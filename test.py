import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

# 创建示例 DataFrame
data = {
    '1WEEK': [0.02, 0.03, 0.01],
    '1MONTH': [0.05, 0.06, 0.07],
    '2MONTH': [None, None, None],
    '3MONTH': [0.08, 0.09, 0.1],
    '4MONTH': [None, None, None],
    '5MONTH': [None, None, None],
    '6MONTH': [0.15, 0.16, 0.14],
    '7MONTH': [None, None, None],
    '8MONTH': [None, None, None],
    '9MONTH': [None, None, None],
    '10MONTH': [None, None, None],
    '11MONTH': [None, None, None],
    '12MONTH': [0.2, 0.21, 0.22]
}
dates = pd.date_range(start='2024-04-01', periods=3, freq='D')
df = pd.DataFrame(data, index=dates)

# 将缺失值替换为 0
df.fillna(0, inplace=True)

first_row = df.iloc[0]
second_row = df.iloc[1]

# 计算变化超过 0.1% 的列
change_columns = []
for column in df.columns:
    if abs(first_row[column] - second_row[column]) > 0.001 * abs(first_row[column]):
        change_columns.append(column)

# 构造邮件内容
mail_content = ""
if change_columns:
    mail_content = " ".join(change_columns) + " change more than 0.1%"

# 发送邮件
if mail_content:
    sender_email = "chengguoyu_82@163.com"
    receiver_email = "wo_oplove@163.com"
    password = "SSJTQGALEZMNHNGE"  # 你的邮箱密码

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = "Alert: Rate Changes"

    message.attach(MIMEText(mail_content, "plain"))

    # 建立连接并登录到你的邮箱
    with smtplib.SMTP("smtp.163.com", 25) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())

        print("Email sent successfully!")
else:
    print("No significant changes to report.")
