import requests
import pandas as pd
from datetime import datetime
import tabula
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging

# 忽略 PDFBox 的字体警告
logging.getLogger("org.apache.fontbox").setLevel(logging.ERROR)

import os

def check_file_exists():
    current_date = datetime.now().strftime("%y%m%d")
    filename = f"{current_date}.pdf"
    return os.path.exists(filename)

def save_file(url, filename):
    response = requests.get(url)
    if response.status_code == 200:
        with open(filename, 'wb') as f:
            f.write(response.content)
        print(f"文件 '{filename}' 下载成功！")
    else:
        print(f"下载失败，状态码：{response.status_code}")

def send_email(sender_email, sender_password, recipient_emails, subject, body, attachments=None):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipient_emails)  # 邮件头部用逗号连接
    msg['Subject'] = subject
    
    css = '''
        <style>
        table{
            border-collapse: collapse;
            width:100%;
            border:1px solid #c6c6c6 !important;
            margin-bottom:20px;
        }
        table th{
            border-collapse: collapse;
            border-right:1px solid #c6c6c6 !important;
            border-bottom:1px solid #c6c6c6 !important;
            background-color:#ddeeff !important; 
            padding:5px 9px;
            font-size:14px;
            font-weight:normal;
            text-align:center;
        }
        table td{
            border-collapse: collapse;
            border-right:1px solid #c6c6c6 !important;
            border-bottom:1px solid #c6c6c6 !important; 
            padding:5px 9px;
            font-size:12px;
            font-weight:normal;
            text-align:center;
            word-break: break-all;
        }
        table tr:nth-child(odd){
            background-color:#fff !important; 
        }
        table tr:nth-child(even){
            background-color: #f8f8f8 !important;
        }
        </style>
    '''
    msg.attach(MIMEText(css + body, 'html'))

    if attachments:
        for attachment in attachments:
            with open(attachment, 'rb') as file:
                msg.attach(MIMEText(file.read(), 'plain', _charset='utf-8'))

    with smtplib.SMTP_SSL('smtp.163.com', 465) as server:
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_emails, msg.as_string())

def calculate_change(df):
    change_list = []
    for column in df.columns:
        try:
            change = df[column].iloc[0] - df[column].iloc[1]
            if abs(change) > 0.001:
                change_list.append(column)
        except (IndexError, TypeError):
            continue
    return change_list

if __name__ == "__main__":
    if not check_file_exists():
        try:
            current_date = datetime.now().strftime("%y%m%d")
            pdf_url = f"https://www.jbatibor.or.jp/rate/pdf/JAPANESEYENTIBOR{current_date}.pdf"
            filename = f"{current_date}.pdf"
            
            # 先下载文件
            save_file(pdf_url, filename)
            
            # 解析本地 PDF
            tables = tabula.read_pdf(filename, pages="all")
            
            dfs = [pd.DataFrame(table) for table in tables]
            df = pd.concat(dfs, ignore_index=True)
            
            df.set_index(df.columns[0], inplace=True)
            df.index.rename('date', inplace=True)
            html_table = df.fillna('').to_html(border=1)
            df.fillna(0, inplace=True)
            
            sender_email = "chengguoyu_82@163.com"
            sender_password = "DUigKtCtMXw34MnB"
            recipient_emails = ["chengguoyu_82@163.com"]
            subject = "Japanese Yen TIBOR"
            
            # 修复链接
            body = f"<p>Download PDF <a href='{pdf_url}' target='_blank'>click me!</a></p><br/><div>{html_table}</div><br/>"
            
            change_list = calculate_change(df)
            if change_list:
                change_message = ", ".join(change_list) + " changed by more than 0.1%"
                body = f"**<h3><font color='red'><b>Please note that {change_message}</b></font></h3>**<br/>" + body
            
            # 发送邮件（传递收件人列表）
            send_email(sender_email, sender_password, recipient_emails, subject, body)
            
            print(df)
        except Exception as e:
            print("运行时错误:", e)
    else:
        print("文件已存在，跳过程序执行。")
