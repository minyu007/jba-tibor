import requests
import pandas as pd
from datetime import datetime
import tabula
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

try:
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
        mail_content = " ".join(change_columns) + " changes more than 0.1%"
        mail_content = "Hi, !\n\n" + mail_content + "\n\n Regards,\nRobot"
        mail_content += "\n\nRefer to the link for more information:" + base_url

    if mail_content:
        sender_email = "chengguoyu_82@163.com"
        receiver_emails = ["zling@jenseninvest.com", "13889632722@163.com"]
        password = "SSJTQGALEZMNHNGE"
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = ", ".join(receiver_emails)
        message["Subject"] = "Alert: Rate Changes"
        message.attach(MIMEText(mail_content, "plain"))
        with smtplib.SMTP("smtp.163.com", 25) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_emails, message.as_string())
            print("Email sent successfully!")
    else:
        print("No significant changes to report.")

    print(df)

except Exception as e:
    print("running into an error:", e)
