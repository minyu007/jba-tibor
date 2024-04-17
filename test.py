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

    msg.attach(MIMEText(body, 'html'))

    if attachments:
        for attachment in attachments:
            with open(attachment, 'rb') as file:
                img_data = MIMEImage(file.read(), name=attachment)
                msg.attach(img_data)

    with smtplib.SMTP_SSL('smtp.163.com', 465) as server:
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())


def calculate_change(df):
    change_message = ""
    for column in df.columns:
        change = df[column].iloc[0] - df[column].iloc[1]
        if abs(change) > 0.001:
            change_message += f"{column},"
    change_message += f" change more than 0.1%\n"
    return change_message


html_table = df.fillna('').to_html(border=1)
df.fillna(0, inplace=True)


sender_email = "chengguoyu_82@163.com"
sender_password = "SSJTQGALEZMNHNGE"
recipient_emails = ["wo_oplove@163.com", "chengguoyu_82@163.com"]
subject = "Japanese Yen TIBOR"
body = "<div>"+html_table+"</div><br/><p>Download PDF <a href='" + \
    "dd"+"' target='_blank'>click me!</a></p><br/>"

# attachments = ['table_image.png', 'line_chart_image.png']
attachments = []


change_list = calculate_change(df)
# change_message = ""
if len(change_list) > 0:
    change_message = f",".join(change_list) + " change more than 0.1%"
    body += f"**<font color='red'><b>{change_message}</b></font>**"

send_email(sender_email, sender_password, ','.join(
    recipient_emails), subject, body, attachments)

print(df)