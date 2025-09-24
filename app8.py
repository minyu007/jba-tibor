import requests
import pandas as pd
from datetime import datetime, timedelta
import tabula
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import logging
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import io
import os

# 屏蔽冗余日志
logging.getLogger("org.apache.fontbox").setLevel(logging.ERROR)
logging.getLogger("tabula").setLevel(logging.WARNING)

def check_file_exists():
    current_date = datetime.now().strftime("%y%m%d")
    filename = f"{current_date}.pdf"
    return os.path.exists(filename)

def save_file(url, filename):
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        with open(filename, 'wb') as f:
            f.write(response.content)
        print(f"文件 '{filename}' 下载成功！")
        return True
    except Exception as e:
        print(f"下载失败：{str(e)}")
        return False

def create_line_chart(df):
    plt.figure(figsize=(12, 6))
    
    # 筛选数值列（排除非数值列）
    plot_columns = [
        col for col in df.columns 
        if pd.api.types.is_numeric_dtype(df[col]) and not df[col].fillna(0).eq(0).all()
    ]
    
    if not plot_columns:
        print("无有效数值列用于绘图")
        return None
    
    # 确保日期索引有效
    if not isinstance(df.index, pd.DatetimeIndex):
        try:
            df.index = pd.to_datetime(df.index, errors='coerce')
            df = df.dropna(subset=[df.index.name])
        except Exception as e:
            print(f"日期转换错误: {e}")
            return None
    
    df = df.sort_index()  # 按日期升序
    
    # 绘制曲线并标注变化
    for column in plot_columns:
        plt.plot(df.index, df[column], marker='o', label=column, linewidth=1.5)
        
        if len(df) >= 2:
            changes = df[column].diff()
            for i in range(1, len(df)):
                change = changes.iloc[i]
                if abs(change) > 0.001:  # 变化超0.1%标注
                    date = df.index[i]
                    y_val = df[column].iloc[i]
                    prev_val = df[column].iloc[i-1]
                    change_pct = (change / prev_val) * 100
                    
                    arrow_dir = '↑' if change > 0 else '↓'
                    color = 'red' if change > 0 else 'blue'
                    bg_color = 'lightcoral' if change > 0 else 'lightblue'
                    
                    plt.annotate(
                        f'{arrow_dir}{abs(change_pct):.2f}%',
                        xy=(date, y_val),
                        xytext=(0, 15 if change > 0 else -15),
                        textcoords='offset points',
                        ha='center', va='center',
                        bbox=dict(boxstyle='round,pad=0.3', fc=bg_color, alpha=0.7),
                        arrowprops=dict(arrowstyle='->', color=color, linewidth=1)
                    )
    
    # 图表格式化
    plt.title('Japanese Yen TIBOR Rates (Daily Changes)', fontsize=14)
    plt.ylabel('Rate (%)', fontsize=12)
    plt.xlabel('Date', fontsize=12)
    
    ax = plt.gca()
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
    plt.xticks(rotation=45, ha='right')
    
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=10)
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    
    # 保存为字节流
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
    buf.seek(0)
    plt.close()
    return buf

def send_email(sender_email, sender_password, recipient_emails, subject, body, chart_data=None):
    msg = MIMEMultipart('related')
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipient_emails)
    msg['Subject'] = subject
    
    # CSS样式
    css = '''
        <style>
        table {border-collapse: collapse; width:100%; margin-bottom:20px;}
        th {border:1px solid #c6c6c6; background:#ddeeff; padding:6px; font-size:13px; text-align:center;}
        td {border:1px solid #c6c6c6; padding:6px; font-size:12px; text-align:center;}
        tr:nth-child(even) {background:#f8f8f8;}
        </style>
    '''
    
    # 邮件正文
    msg_alternative = MIMEMultipart('alternative')
    msg.attach(msg_alternative)
    
    if chart_data:
        body = f'<h3>TIBOR Rates Trend</h3><img src="cid:chart" width="800"><br/>{body}'
    msg_alternative.attach(MIMEText(css + body, 'html', 'utf-8'))
    
    # 内嵌图表
    if chart_data:
        image = MIMEImage(chart_data.read())
        image.add_header('Content-ID', '<chart>')
        msg.attach(image)
    
    # 发送邮件
    try:
        with smtplib.SMTP_SSL('smtp.163.com', 465, timeout=10) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_emails, msg.as_string())
        print("邮件发送成功！")
    except Exception as e:
        print(f"邮件发送失败：{str(e)}")

def calculate_change(df):
    """检测变化超0.1%的列"""
    change_list = []
    if len(df) < 2:
        return change_list
    
    for column in df.columns:
        if not pd.api.types.is_numeric_dtype(df[column]):
            continue
        latest_val = df[column].iloc[-1]
        prev_val = df[column].iloc[-2]
        if prev_val == 0:
            continue
        change_pct = abs((latest_val - prev_val) / prev_val * 100)
        if change_pct > 0.1:
            change_list.append(column)
    return change_list

def split_data_to_dates(df, pdf_date):
    """
    拆分\r分隔的利率数据，补充Date列
    df: 读取的PDF表格（2行13列：第0行列名，第1行数据）
    pdf_date: PDF对应的日期（从文件名提取，如250924→2025-09-24）
    """
    # 1. 提取列名（第0行：1WEEK~12MONTH）
    columns = df.iloc[0].tolist()
    # 2. 提取利率数据（第1行：每个单元格是\r分隔的多个数值）
    rate_data = df.iloc[1].tolist()
    
    # 3. 拆分每个期限的利率，确保所有期限的数值个数一致
    split_rates = []
    for rates in rate_data:
        # 拆分\r，转换为浮点数（跳过空值）
        split = [float(r.strip()) for r in str(rates).split('\r') if r.strip() and r.strip().replace('.','').isdigit()]
        split_rates.append(split)
    
    # 4. 检查所有期限的数值个数是否一致（确保每行对应一个日期）
    num_dates = len(split_rates[0]) if split_rates else 0
    for rates in split_rates:
        if len(rates) != num_dates:
            raise Exception(f"各期限数据个数不一致：预期{num_dates}个，实际{len(rates)}个")
    
    # 5. 生成Date列（从PDF日期倒推，如PDF是2025-09-24，数据有5个则生成24/23/22/21/20）
    dates = [pdf_date - timedelta(days=i) for i in range(num_dates)][::-1]  # 倒序→正序
    
    # 6. 构建最终DataFrame（Date + 各期限利率）
    final_data = {
        'Date': dates,
        **{columns[i]: split_rates[i] for i in range(len(columns))}
    }
    final_df = pd.DataFrame(final_data)
    
    return final_df

if __name__ == "__main__":
    # 1. 基础配置（日期、文件名、URL）
    current_date_str = datetime.now().strftime("%y%m%d")  # 如250924
    pdf_date = datetime.strptime(current_date_str, "%y%m%d")  # 转换为日期对象：2025-09-24
    filename = f"{current_date_str}.pdf"
    pdf_url = f"https://www.jbatibor.or.jp/rate/pdf/JAPANESEYENTIBOR{current_date_str}.pdf"
    
    # 2. 检查文件是否已存在
    if check_file_exists():
        print(f"文件 '{filename}' 已存在，跳过程序执行。")
        exit()
    
    try:
        # 3. 下载PDF
        if not save_file(pdf_url, filename):
            raise Exception("PDF下载失败，终止程序")
        
        # 4. 读取PDF表格（关键：只读取1个表格，保留原始结构）
        tables = tabula.read_pdf(
            filename,
            pages="all",
            multiple_tables=True,
            lattice=True,  # 适配网格线表格
            guess=True,
            pandas_options={'header': None},  # 不自动设表头（手动处理第0行）
            encoding='utf-8'
        )
        
        # 检查表格数量（预期1个）
        if len(tables) != 1:
            raise Exception(f"预期读取1个表格，实际读取{len(tables)}个")
        
        df_raw = tables[0].copy()
        print(f"原始表格形状：行数={df_raw.shape[0]}, 列数={df_raw.shape[1]}")
        print("原始表格前2行（列名+数据）：")
        print(df_raw.head(2))
        
        # 5. 拆分数据并补充Date列（核心修正！）
        df = split_data_to_dates(df_raw, pdf_date)
        print(f"\n处理后表格形状：行数={df.shape[0]}, 列数={df.shape[1]}")
        print("处理后表格预览：")
        print(df.head())
        
        # 6. 日期列处理（设为索引）
        df['Date'] = pd.to_datetime(df['Date'])
        df = df.dropna(subset=['Date']).reset_index(drop=True)
        df.set_index('Date', inplace=True)
        df.index.rename('date', inplace=True)
        
        # 7. 数据清洗（NaN填充为0）
        df_numeric = df.fillna(0)
        
        # 8. 构建邮件内容
        html_table = df.fillna('').to_html(index=True, border=0)
        body = f"<p>PDF下载链接：<a href='{pdf_url}' target='_blank'>点击下载</a></p><br/><div>{html_table}</div><br/>"
        
        # 9. 检测变化超0.1%的列，添加警告
        change_list = calculate_change(df_numeric)
        if change_list:
            change_msg = ", ".join(change_list) + " 变动超过0.1%"
            body = f"<h3 style='color:red;'>注意：{change_msg}</h3><br/>" + body
        
        # 10. 生成图表（至少2行数据就绘图，之前5行太严格）
        chart_data = None
        if len(df_numeric) >= 2:
            chart_data = create_line_chart(df_numeric)
            print("图表生成成功")
        else:
            print("数据行数不足2行，不生成图表")
        
        # 11. 发送邮件（替换为你的邮箱配置）
        sender_email = "chengguoyu_82@163.com"
        sender_password = "DUigKtCtMXw34MnB"  # 重要：163邮箱需用【授权码】，不是登录密码！
        recipient_emails = ["chengguoyu_82@163.com"]
        subject = f"Japanese Yen TIBOR - {pdf_date.strftime('%Y-%m-%d')}"
        
        send_email(sender_email, sender_password, recipient_emails, subject, body, chart_data)
        
    except Exception as e:
        print(f"\n运行时错误：{str(e)}")
        # 可选：错误时发送告警邮件
        # send_email(sender_email, sender_password, ["你的邮箱"], "TIBOR程序报错", f"错误信息：{str(e)}")
