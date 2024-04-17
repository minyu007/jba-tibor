import requests
import pandas as pd
from datetime import datetime
import tabula
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import matplotlib.pyplot as plt


def send_email(sender_email, sender_password, recipient_email, subject, body, attachments=None):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

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


try:
    current_date = datetime.now().strftime("%y%m%d")

    base_url = "https://www.jbatibor.or.jp/rate/pdf/JAPANESEYENTIBOR{}.pdf"

    pdf_url = base_url.format(current_date)

    tables = tabula.read_pdf(pdf_url, pages="all")

    dfs = [pd.DataFrame(table) for table in tables]

    df = pd.concat(dfs, ignore_index=True)

    df.set_index(df.columns[0], inplace=True)
    df.index.rename('date', inplace=True)
    html_table = df.to_html(border=1)
    df.fillna(0, inplace=True)

    # plt.figure(figsize=(10, 6))
    # plt.table(cellText=df.values, colLabels=df.columns, rowLabels=df.index.strftime(
    #     '%Y-%m-%d'), loc='center', cellLoc='center')
    # plt.axis('off')
    # plt.title('Table Image 1', fontsize=16)
    # plt.savefig('table_image.png')

    # plt.figure(figsize=(10, 6))
    # for column in df.columns:
    #     plt.plot(df.index, df[column], label=column)
    # plt.xlabel('Date')
    # plt.ylabel('Value')
    # plt.title('Line Chart Image 2')
    # plt.legend()
    # plt.savefig('line_chart_image.png')
    # html_table = df.to_html()
    sender_email = "chengguoyu_82@163.com"
    sender_password = "SSJTQGALEZMNHNGE"
    recipient_emails = ["wo_oplove@163.com", "13889632722@163.com"]
    subject = "Japanese Yen TIBOR"
    body = "<p>Hi All, </p><br/><div>"+html_table+"</div><br/><p>refer to the link for more information <a href='" + \
        base_url+"' target='_blank'>click me!</a></p><br/>"

    # attachments = ['table_image.png', 'line_chart_image.png']
    attachments = []
    change_message = calculate_change(df)
    if change_message:
        body += f"**<font color='red'><b>{change_message}</b></font>**"

    send_email(sender_email, sender_password, ','.join(
        recipient_emails), subject, body, attachments)

    print(df)

except Exception as e:
    print("running into an error:", e)
