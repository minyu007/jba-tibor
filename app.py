import requests
import pandas as pd
from datetime import datetime
import tabula
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


current_date = datetime.now().strftime("%y%m%d")

base_url = "https://www.jbatibor.or.jp/rate/pdf/JAPANESEYENTIBOR{}.pdf"

pdf_url = base_url.format(current_date)

tables = tabula.read_pdf(pdf_url, pages="all")

dfs = [pd.DataFrame(table) for table in tables]

df = pd.concat(dfs, ignore_index=True)

df.set_index(df.columns[0], inplace=True)
df.index.rename('date', inplace=True)


df.fillna(0, inplace=True)

first_row = df.iloc[0]
second_row = df.iloc[1]

change_columns = []
for column in df.columns:
    if abs(first_row[column] - second_row[column]) > 0.001 * abs(first_row[column]):
        change_columns.append(column)

mail_content = ""
if change_columns:
    mail_content = " ".join(change_columns) + " change more than 0.1%"

if mail_content:
    sender_email = "chengguoyu_82@163.com"
    receiver_email = "wo_oplove@163.com"
    password = "SSJTQGALEZMNHNGE"
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = "Alert: Rate Changes"
    message.attach(MIMEText(mail_content, "plain"))
    with smtplib.SMTP("smtp.163.com", 25) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        print("Email sent successfully!")
else:
    print("No significant changes to report.")

print(df)
