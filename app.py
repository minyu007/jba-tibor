import requests
import pandas as pd
from datetime import datetime
import tabula
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


current_date = datetime.now().strftime("%y%m%d")

base_url = "https://www.jbatibor.or.jp/rate/pdf/JAPANESEYENTIBOR{}.pdf"

pdf_url = base_url.format(current_date)

tables = tabula.read_pdf(pdf_url, pages="all")

dfs = [pd.DataFrame(table) for table in tables]

df = pd.concat(dfs, ignore_index=True)

df.set_index(df.columns[0], inplace=True)
df.index.rename('date', inplace=True)


threshold = 0.001  # 0.1%
columns_to_check = ['1WEEK', '1MONTH', '3MONTH', '6MONTH', '12MONTH']
for column in columns_to_check:
    if (df[column].pct_change().iloc[0] > threshold).any():
        smtp_server = "smtp.163.com"
        smtp_port = 465
        sender_email = "chengguoyu_82@163.com"  # 你的邮箱地址
        sender_password = "xiaoyu8740"  # 你的邮箱密码

        receiver_email = "chengguoyu_82@163.com"  # 收件人邮箱地址

        # 创建邮件
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = f"{column} change more than 0.1%"

        body = f"The {column} column change more than 0.1%"
        msg.attach(MIMEText(body, 'plain'))

        # 发送邮件
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())

print(df)
