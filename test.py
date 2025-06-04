import requests
import pandas as pd
from datetime import datetime
import pdfplumber  # 替换 tabula 的关键库
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging
import os

# 忽略 PDFBox 的字体警告（pdfplumber 无此警告，可保留或删除）
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
            if abs(change) > 0.001:  # 0.001 对应 0.1% 的变化（假设数值为百分比格式）
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
            
            # 下载文件
            save_file(pdf_url, filename)
            
            # ====================== 使用 pdfplumber 解析 PDF ======================
            dfs = []
            with pdfplumber.open(filename) as pdf:
                for page in pdf.pages:
                    # 提取表格（自动检测，可根据需要调整参数）
                    tables = page.extract_tables()
                    for table in tables:
                        if table:  # 过滤空表格
                            # 处理首行作为表头（若表格包含表头）
                            if len(table) >= 1:
                                header = table[0]
                                data = table[1:]
                                df_table = pd.DataFrame(data, columns=header)
                            else:
                                # 若没有表头，使用默认列名
                                df_table = pd.DataFrame(table)
                            dfs.append(df_table)
            
            if not dfs:
                raise ValueError("未在 PDF 中检测到表格数据")
            
            # 合并所有表格数据
            df = pd.concat(dfs, ignore_index=True)
            
            # ====================== 原代码的数据处理逻辑 ======================
            if not df.empty:
                # 设置索引（假设第一列为日期）
                df.set_index(df.columns[0], inplace=True)
                df.index.rename('date', inplace=True)
                
                # 处理空值并生成 HTML 表格
                html_table = df.fillna('').to_html(border=1)
                df.fillna(0, inplace=True)
            else:
                raise ValueError("解析后的 DataFrame 为空")
            
            # 邮件发送配置
            sender_email = "chengguoyu_82@163.com"
            sender_password = "DUigKtCtMXw34MnB"  # 注意：需替换为实际邮箱授权码
            # recipient_emails = ["zling@jenseninvest.com", "hwang@jenseninvest.com", "yqguo@jenseninvest.com", "13889632722@163.com"]
            recipient_emails = ["wo_oplove@163.com"]
            subject = "Japanese Yen TIBOR"
            
            # 构建邮件正文
            body = f"<p>Download PDF <a href='{pdf_url}' target='_blank'>click me!</a></p><br/><div>{html_table}</div><br/>"
            
            # 检测数据变化
            change_list = calculate_change(df)
            if change_list:
                change_message = ", ".join(change_list) + " changed by more than 0.1%"
                body = f"**<h3><font color='red'><b>Please note that {change_message}</b></font></h3>**<br/>" + body
            
            # 发送邮件
            send_email(sender_email, sender_password, recipient_emails, subject, body)
            
            print("解析结果：")
            print(df)
            
        except Exception as e:
            print("运行时错误:", e)
    else:
        print("文件已存在，跳过程序执行。")
