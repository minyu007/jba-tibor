import requests
import pandas as pd
from datetime import datetime
import tabula
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging
import os

# 忽略 PDFBox 警告
logging.getLogger("org.apache.fontbox").setLevel(logging.ERROR)

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
    msg['To'] = ", ".join(recipient_emails)  
    msg['Subject'] = subject
    
    # 保留原始 CSS 确保表格样式
    css = '''
        <style>
        table{border-collapse: collapse;width:100%;border:1px solid #c6c6c6 !important;margin-bottom:20px;}
        table th{border-collapse: collapse;border-right:1px solid #c6c6c6 !important;border-bottom:1px solid #c6c6c6 !important;background-color:#ddeeff !important;padding:5px 9px;font-size:14px;font-weight:normal;text-align:center;}
        table td{border-collapse: collapse;border-right:1px solid #c6c6c6 !important;border-bottom:1px solid #c6c6c6 !important;padding:5px 9px;font-size:12px;font-weight:normal;text-align:center;word-break: break-all;}
        table tr:nth-child(odd){background-color:#fff !important;}
        table tr:nth-child(even){background-color: #f8f8f8 !important;}
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
            
            # 下载 PDF
            save_file(pdf_url, filename)
            
            # 关键修改：强制指定表头 + 修复解析区域
            tables = tabula.read_pdf(
                filename,
                pages="all",
                lattice=True,  
                pandas_options={
                    "header": 0,  
                    "names": ["date", "1WEEK", "1MONTH", "2MONTH", "3MONTH", "4MONTH", 
                              "5MONTH", "6MONTH", "7MONTH", "8MONTH", "9MONTH", 
                              "10MONTH", "11MONTH", "12MONTH"]  
                },
                area=(50, 20, 700, 550),
                # 根据 PDF 实际列间隔，手动设置列分割坐标，比如 [x1, x2, x3,...] ，x 是水平方向分割点
                columns=[50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700]  
            )
            
            # 处理空表格
            if not tables or all(table.empty for table in tables):
                raise ValueError("未检测到PDF表格数据")
            
            # 合并 + 清洗：仅保留日期行
            df = pd.concat(tables, ignore_index=True)
            df = df[df['date'].str.match(r'^\d{4}/\d{2}/\d{2}$', na=False)]  # 过滤日期
            
            # 设置索引（保持原逻辑）
            df.set_index('date', inplace=True)
            
            # 生成 HTML（确保表头正确）
            html_table = df.fillna('').to_html(border=1)
            df.fillna(0, inplace=True)
            
            # 邮件参数
            sender_email = "chengguoyu_82@163.com"
            sender_password = "DUigKtCtMXw34MnB"
            # recipient_emails = ["zling@jenseninvest.com","hwang@jenseninvest.com", "yqguo@jenseninvest.com", "13889632722@163.com"]
            recipient_emails = ["wo_oplove@163.com"]
            subject = "Japanese Yen TIBOR"
            
            # 拼接邮件内容
            body = f"<p>Download PDF <a href='{pdf_url}' target='_blank'>click me!</a></p><br/><div>{html_table}</div><br/>"
            
            # 计算变化列
            change_list = calculate_change(df)
            if change_list:
                change_message = ", ".join(change_list) + " changed by more than 0.1%"
                body = f"**<h3><font color='red'><b>Please note that {change_message}</b></font></h3>**<br/>" + body
            
            # 发送邮件
            send_email(sender_email, sender_password, recipient_emails, subject, body)
            
            # 打印完整 DataFrame（含正确表头）
            print(df)
            
        except ValueError as ve:
            print("解析错误:", ve)
        except Exception as e:
            print("运行时错误:", e)
    else:
        print("文件已存在，跳过程序执行。")
