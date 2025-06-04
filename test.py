import requests
import pandas as pd
from datetime import datetime
import pdfplumber  # 替换 tabula
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging
import os

# 忽略警告
logging.getLogger("pdfminer").setLevel(logging.WARNING)
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
    # 保持邮件发送逻辑不变
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipient_emails)
    msg['Subject'] = subject
    
    css = '''
        <style>
        table{border-collapse:collapse;width:100%;border:1px solid #c6c6c6;margin-bottom:20px;}
        table th{border-right:1px solid #c6c6c6;border-bottom:1px solid #c6c6c6;background-color:#ddeeff;padding:5px 9px;text-align:center;}
        table td{border-right:1px solid #c6c6c6;border-bottom:1px solid #c6c6c6;padding:5px 9px;text-align:center;word-break:break-all;}
        table tr:nth-child(odd){background-color:#fff;}
        table tr:nth-child(even){background-color:#f8f8f8;}
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

def parse_multi_line_table(tables):
    """拆分单元格内的多行文本为独立行"""
    parsed_data = []
    for table in tables:
        if not table:
            continue
        
        # 提取表头（假设第一行为表头）
        if len(table) >= 1:
            header = table[0]
            data_rows = table[1:]
        else:
            header = None
            data_rows = table
        
        # 处理每行数据
        for row in data_rows:
            new_cells = []
            for cell in row:
                # 按换行符拆分并过滤空值
                values = [v.strip() for v in cell.split('\n') if v.strip()]
                new_cells.extend(values)
            
            # 按表头列数分组
            if header:
                num_cols = len(header)
                if len(new_cells) % num_cols != 0:
                    continue  # 跳过格式异常的行
                split_rows = [new_cells[i:i+num_cols] for i in range(0, len(new_cells), num_cols)]
                parsed_data.extend(split_rows)
            else:
                parsed_data.append(new_cells)
    
    # 转换为 DataFrame
    return pd.DataFrame(parsed_data, columns=header) if header else pd.DataFrame(parsed_data)

def calculate_change(df):
    change_list = []
    for column in df.columns:
        try:
            change = float(df[column].iloc[0]) - float(df[column].iloc[1])
            if abs(change) > 0.001:  # 0.1% 变化阈值
                change_list.append(column)
        except (IndexError, TypeError, ValueError):
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
            
            # ================== 使用 pdfplumber 解析 ==================
            dfs = []
            with pdfplumber.open(filename) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()  # 提取原始表格
                    if tables:
                        # 拆分多行文本并解析
                        df_table = parse_multi_line_table(tables)
                        dfs.append(df_table)
            
            if not dfs:
                raise ValueError("未检测到表格数据")
            
            # 合并数据
            df = pd.concat(dfs, ignore_index=True).dropna(how='all')
            
            # 设置索引（假设第一列为日期）
            if not df.empty and len(df.columns) > 0:
                df.set_index(df.columns[0], inplace=True)
                df.index.rename('date', inplace=True)
            
            # 处理空值
            html_table = df.fillna('').to_html(border=1, index=True)
            df.fillna(0, inplace=True)
            
            # 邮件配置（需替换为实际邮箱信息）
            sender_email = "chengguoyu_82@163.com"
            sender_password = "DUigKtCtMXw34MnB"  # 授权码
            # recipient_emails = ["zling@jenseninvest.com", "hwang@jenseninvest.com", "yqguo@jenseninvest.com", "13889632722@163.com"]
            recipient_emails = ["wo_oplove@163.com"]
            subject = "Japanese Yen TIBOR"
            
            # 构建正文
            body = f"<p>Download PDF <a href='{pdf_url}' target='_blank'>click me!</a></p><br/>{html_table}<br/>"
            
            # 检测数据变化
            if len(df) >= 2:
                change_list = calculate_change(df)
                if change_list:
                    change_msg = ", ".join(change_list) + " changed by more than 0.1%"
                    body = f"<h3><font color='red'><b>⚠️ Notice: {change_msg}</b></font></h3><br/>" + body
            
            # 发送邮件
            send_email(sender_email, sender_password, recipient_emails, subject, body)
            
            print("解析结果：")
            print(df)
            
        except Exception as e:
            print("运行时错误:", e)
    else:
        print("文件已存在，跳过程序执行。")
