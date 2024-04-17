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


plt.figure(figsize=(10, 6))
plt.table(cellText=df.values, colLabels=df.columns, rowLabels=df.index.strftime(
    '%Y-%m-%d'), loc='center', cellLoc='center')
plt.axis('off')
plt.title('Table Image 1', fontsize=16)
plt.savefig('table_image.png')

plt.figure(figsize=(10, 6))
for column in df.columns:
    plt.plot(df.index, df[column], label=column)
plt.xlabel('Date')
plt.ylabel('Value')
plt.title('Line Chart Image 2')
plt.legend()
plt.savefig('line_chart_image.png')

sender_email = "chengguoyu_82@163.com"
sender_password = "SSJTQGALEZMNHNGE"  # 请替换为你的密码
recipient_emails = ["wo_oplove@163.com", "chengguoyu_82@163.com"]
subject = "Data and Charts"
body = "**Please see the attached images.**\n\n" \
    "Below are the changes:\n\n"
attachments = ['table_image.png', 'line_chart_image.png']

change_message = calculate_change(df)
if change_message:
    body += f"**<font color='red'><b>{change_message}</b></font>**"

send_email(sender_email, sender_password,
           ','.join(recipient_emails), subject, body, attachments)

# first_row = df.iloc[0]
# second_row = df.iloc[1]

# change_columns = []
# for column in df.columns:
#     if abs(first_row[column] - second_row[column]) > 0.001 * abs(first_row[column]):
#         change_columns.append(column)

# mail_content = ""
# if change_columns:
#     mail_content = " ".join(change_columns) + " change more than 0.1%"
#     mail_content = "Hi, recipient!\n\n" + \
#         mail_content + "\n\nBest regards,\nYour Sender"
#     mail_content += "\n\nClick the link for more information: http://baid.com"

# # mail_content = ""
# # if change_columns:
# #     mail_content = " ".join(change_columns) + " change more than 0.1%"

# if mail_content:
#     sender_email = "chengguoyu_82@163.com"
#     receiver_email = "wo_oplove@163.com"
#     password = "SSJTQGALEZMNHNGE"  # 你的邮箱密码
#     message = MIMEMultipart()
#     message["From"] = sender_email
#     message["To"] = receiver_email
#     message["Subject"] = "Alert: Rate Changes"
#     message.attach(MIMEText(mail_content, "plain"))
#     with smtplib.SMTP("smtp.163.com", 25) as server:
#         server.login(sender_email, password)
#         server.sendmail(sender_email, receiver_email, message.as_string())
#         print("Email sent successfully!")
# else:
#     print("No significant changes to report.")
