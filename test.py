import requests
import pandas as pd
from datetime import datetime
import tabula
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging

# 更全面的日志配置来消除所有 PDFBox 警告
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfminer.psparser").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdfdocument").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdfinterp").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdfpage").setLevel(logging.ERROR)
logging.getLogger("pdfminer.converter").setLevel(logging.ERROR)
logging.getLogger("pdfminer.cmapdb").setLevel(logging.ERROR)
logging.getLogger("org.apache.fontbox").setLevel(logging.ERROR)
logging.getLogger("org.apache.pdfbox").setLevel(logging.ERROR)

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
        return True
    else:
        print(f"下载失败，状态码：{response.status_code}")
        return False

def send_email(sender_email, sender_password, recipient_emails, subject, body, attachments=None):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipient_emails)
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
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(attachment)}"')
                msg.attach(part)

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
            current_date='250603'
            pdf_url = f"https://www.jbatibor.or.jp/rate/pdf/JAPANESEYENTIBOR{current_date}.pdf"
            filename = f"{current_date}.pdf"
            
            # 下载文件
            if not save_file(pdf_url, filename):
                raise Exception("Failed to download PDF file")
            
            # 尝试不同的提取方法
            try:
                # 方法1: 使用lattice模式
                tables = tabula.read_pdf(filename, pages="all", lattice=True)
            except:
                try:
                    # 方法2: 使用stream模式
                    tables = tabula.read_pdf(filename, pages="all", stream=True)
                except Exception as e:
                    raise Exception(f"Failed to extract tables from PDF: {str(e)}")
            
            if not tables:
                raise Exception("No tables found in PDF")
            
            # 合并表格前检查
            valid_tables = [table for table in tables if isinstance(table, pd.DataFrame) and not table.empty]
            if not valid_tables:
                raise Exception("No valid tables to concatenate")
            
            df = pd.concat(valid_tables, ignore_index=True)
            
            if df.empty:
                raise Exception("Empty DataFrame after concatenation")
            
            df.set_index(df.columns[0], inplace=True)
            df.index.rename('date', inplace=True)
            html_table = df.fillna('').to_html(border=1)
            df.fillna(0, inplace=True)
            
            sender_email = "chengguoyu_82@163.com"
            sender_password = "DUigKtCtMXw34MnB"
            recipient_emails = ["wo_oplove@163.com"]
            subject = "Japanese Yen TIBOR"
            
            body = f"<p>Download PDF <a href='{pdf_url}' target='_blank'>click me!</a></p><br/><div>{html_table}</div><br/>"
            
            change_list = calculate_change(df)
            if change_list:
                change_message = ", ".join(change_list) + " changed by more than 0.1%"
                body = f"**<h3><font color='red'><b>Please note that {change_message}</b></font></h3>**<br/>" + body
            
            send_email(sender_email, sender_password, recipient_emails, subject, body)
            
            print("Successfully processed and sent email.")
            print(df)
            
        except Exception as e:
            print("Error:", e)
            # 可以在这里添加发送错误通知邮件的代码
    else:
        print("文件已存在，跳过程序执行。")
