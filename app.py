import requests
import pandas as pd
from datetime import datetime
import tabula
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import matplotlib.pyplot as plt


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

def send_email(sender_email, sender_password, recipient_email, subject, body, attachments=None):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    css='''
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
    msg.attach(MIMEText(css+body, 'html'))

    if attachments:
        for attachment in attachments:
            with open(attachment, 'rb') as file:
                img_data = MIMEImage(file.read(), name=attachment)
                msg.attach(img_data)

    with smtplib.SMTP_SSL('smtp.163.com', 465) as server:
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())


def calculate_change(df):
    change_list = []
    for column in df.columns:
        change = df[column].iloc[0] - df[column].iloc[1]
        if abs(change) > 0.001:
            change_list.append(column)
    
    return change_list



if not check_file_exists():
    try:
        current_date = datetime.now().strftime("%y%m%d")

        base_url = "https://www.jbatibor.or.jp/rate/pdf/JAPANESEYENTIBOR{}.pdf"

        pdf_url = base_url.format(current_date)

        tables = tabula.read_pdf(pdf_url, pages="all")

        dfs = [pd.DataFrame(table) for table in tables]

        df = pd.concat(dfs, ignore_index=True)

        df.set_index(df.columns[0], inplace=True)
        df.index.rename('date', inplace=True)
        html_table = df.fillna('').to_html(border=1)
        df.fillna(0, inplace=True)

        
        sender_email = "chengguoyu_82@163.com"
        sender_password = "DUigKtCtMXw34MnB"
        recipient_emails = ["zling@jenseninvest.com", "13889632722@163.com"]
        # recipient_emails = ["chengguoyu_82@163.com", "wo_oplove@163.com"]
        subject = "Japanese Yen TIBOR"
        body = "<p>Download PDF <a href='" + \
            "dd"+"' target='_blank'>click me!</a></p><br/><div>"+html_table+"</div><br/>"

        # attachments = ['table_image.png', 'line_chart_image.png']
        attachments = []
        
        change_list = calculate_change(df)
        # change_message = ""
        if len(change_list) > 0:
            change_message = f", ".join(change_list) + " change by more than 0.1%"
            body = f"**<h3><font color='red'><b>Please note that {change_message}</b></font></h3>**<br/>" + body

        send_email(sender_email, sender_password, ','.join(
            recipient_emails), subject, body, attachments)
        print(df)
        save_file(pdf_url, f"{current_date}.pdf")
    except Exception as e:
        print("running into an error:", e)
else:
    print("File already exists. Skipping program execution.")

