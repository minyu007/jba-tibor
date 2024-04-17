import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from email.mime.image import MIMEImage
import matplotlib.pyplot as plt

# 创建示例 DataFrame
data = {
    '1WEEK': [0.02, 0.02, 0.01],
    '1MONTH': [0.06, 0.06, 0.07],
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
df.fillna(0, inplace=True)


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


html_table = df.fillna('').to_html(border=1)
df.fillna(0, inplace=True)


sender_email = "chengguoyu_82@163.com"
sender_password = "SSJTQGALEZMNHNGE"
recipient_emails = ["wo_oplove@163.com", "chengguoyu_82@163.com"]
subject = "Japanese Yen TIBOR"
body = "<p>Download PDF <a href='" + \
    "dd"+"' target='_blank'>click me!</a></p><br/><div>"+html_table+"</div><br/>"

# attachments = ['table_image.png', 'line_chart_image.png']
attachments = []


change_list = calculate_change(df)
# change_message = ""
if len(change_list) > 0:
    change_message = f",".join(change_list) + " change by more than 0.1%"
    body = f"**<h3><font color='red'><b>Please note that {change_message}</b></font></h3>**<br/>" + body

send_email(sender_email, sender_password, ','.join(
    recipient_emails), subject, body, attachments)

print(df)