import requests
import pandas as pd
from datetime import datetime
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

# 屏蔽tabula的冗余日志
logging.getLogger("org.apache.fontbox").setLevel(logging.ERROR)
logging.getLogger("tabula").setLevel(logging.WARNING)

def check_file_exists():
    current_date = datetime.now().strftime("%y%m%d")
    filename = f"{current_date}.pdf"
    return os.path.exists(filename)

def save_file(url, filename):
    try:
        response = requests.get(url, timeout=10)  # 加超时防止卡住
        response.raise_for_status()  # 主动抛出HTTP错误（如404、500）
        with open(filename, 'wb') as f:
            f.write(response.content)
        print(f"文件 '{filename}' 下载成功！")
        return True
    except Exception as e:
        print(f"下载失败：{str(e)}")
        return False

def create_line_chart(df):
    plt.figure(figsize=(12, 6))
    
    # 筛选数值列（排除非数值列，避免绘图错误）
    plot_columns = [
        col for col in df.columns 
        if pd.api.types.is_numeric_dtype(df[col]) and not df[col].fillna(0).eq(0).all()
    ]
    
    if not plot_columns:
        print("无有效数值列用于绘图")
        return None
    
    # 日期索引处理
    if not isinstance(df.index, pd.DatetimeIndex):
        try:
            df.index = pd.to_datetime(df.index, errors='coerce')  # 错误日期转为NaT
            df = df.dropna(subset=[df.index.name])  # 删除无效日期行
        except Exception as e:
            print(f"日期转换错误: {e}")
            return None
    
    df = df.sort_index()  # 按日期升序
    
    # 绘制每条曲线并标注变化
    for column in plot_columns:
        plt.plot(df.index, df[column], marker='o', label=column, linewidth=1.5)
        
        if len(df) >= 2:
            changes = df[column].diff()  # 与前一天的差值
            for i in range(1, len(df)):
                change = changes.iloc[i]
                if abs(change) > 0.001:  # 变化超0.1%才标注
                    date = df.index[i]
                    y_val = df[column].iloc[i]
                    prev_val = df[column].iloc[i-1]
                    change_pct = (change / prev_val) * 100
                    
                    # 标注样式
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
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))  # 每天显示一个刻度
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
    msg = MIMEMultipart('related')  # 支持内嵌图片
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipient_emails)
    msg['Subject'] = subject
    
    # CSS样式（优化表格显示）
    css = '''
        <style>
        table {border-collapse: collapse; width:100%; margin-bottom:20px;}
        th {border:1px solid #c6c6c6; background:#ddeeff; padding:6px; font-size:13px; text-align:center;}
        td {border:1px solid #c6c6c6; padding:6px; font-size:12px; text-align:center;}
        tr:nth-child(even) {background:#f8f8f8;}
        </style>
    '''
    
    # 邮件正文（含图表内嵌）
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
    
    # 发送邮件（163邮箱SMTP配置）
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
        return change_list  # 至少2行才计算变化
    
    for column in df.columns:
        if not pd.api.types.is_numeric_dtype(df[column]):
            continue
        # 取最新两行数据（按日期升序，最后一行是最新）
        latest_val = df[column].iloc[-1]
        prev_val = df[column].iloc[-2]
        if prev_val == 0:
            continue  # 避免除以0
        change_pct = abs((latest_val - prev_val) / prev_val * 100)
        if change_pct > 0.1:
            change_list.append(column)
    return change_list

def split_row_to_rows(df):
    """处理PDF中换行的单元格（如日期/利率分行显示）"""
    if df.empty:
        return pd.DataFrame()
    
    # 先处理第一行（通常是数据行，可能有\r换行）
    first_row = df.iloc[0].copy().fillna('')
    split_data = {}
    
    # 按\r分割每个单元格的值
    for col in first_row.index:
        values = [v.strip() for v in first_row[col].split('\r') if v.strip()]
        split_data[col] = values if values else ['']
    
    # 统一所有列的长度（用空字符串填充）
    max_len = max(len(v) for v in split_data.values()) if split_data else 0
    for col in split_data:
        if len(split_data[col]) < max_len:
            split_data[col].extend([''] * (max_len - len(split_data[col])))
    
    # 转换为DataFrame并处理数值类型
    new_df = pd.DataFrame(split_data)
    for col in new_df.columns:
        if col != 'Date':  # 日期列不转数值
            new_df[col] = pd.to_numeric(new_df[col], errors='coerce')  # 非数值转为NaN
    
    # 合并剩余行（除了第一行）
    remaining_df = df.iloc[1:].reset_index(drop=True)
    final_df = pd.concat([new_df, remaining_df], ignore_index=True)
    return final_df

if __name__ == "__main__":
    current_date = datetime.now().strftime("%y%m%d")
    filename = f"{current_date}.pdf"
    pdf_url = f"https://www.jbatibor.or.jp/rate/pdf/JAPANESEYENTIBOR{current_date}.pdf"
    
    if check_file_exists():
        print(f"文件 '{filename}' 已存在，跳过程序执行。")
        exit()
    
    try:
        # 1. 下载PDF
        if not save_file(pdf_url, filename):
            raise Exception("PDF下载失败，终止程序")
        
        # 2. 读取PDF表格（关键：优化tabula参数，确保列识别正确）
        tables = tabula.read_pdf(
            filename,
            pages="all",
            multiple_tables=True,
            lattice=True,  # TIBOR PDF有网格线，用lattice模式
            guess=True,    # 自动识别表格结构
            pandas_options={'header': None},  # 不自动用第一行做表头
            encoding='utf-8'
        )
        
        # 打印表格结构（定位列数问题的核心！）
        print(f"读取到 {len(tables)} 个表格，各表格形状：")
        for i, table in enumerate(tables):
            print(f"  表格{i+1}：行数={table.shape[0]}, 列数={table.shape[1]}")
        
        # 3. 合并表格并查看列数
        df = pd.concat([pd.DataFrame(t) for t in tables], ignore_index=True)
        print(f"合并后DataFrame形状：行数={df.shape[0]}, 列数={df.shape[1]}")
        
        # 4. 关键：根据实际列数调整列名（解决长度不匹配的核心！）
        # 假设合并后列数是13，需删除一个多余的列名（比如12MONTH，根据PDF实际结构调整）
        # 先打印前几行数据，确认列对应的内容
        print("\n合并后DataFrame前3行数据：")
        print(df.head(3))
        
        # ！！根据实际列数修改列名列表！！
        # 示例1：如果列数=14（Date+13个期限），用下面这行
        # df.columns = ['Date', '1WEEK', '1MONTH', '2MONTH', '3MONTH', '4MONTH', '5MONTH', '6MONTH', '7MONTH', '8MONTH', '9MONTH', '10MONTH', '11MONTH', '12MONTH']
        # 示例2：如果列数=13（无12MONTH，或无Date），用下面这行（根据实际数据调整）
        df.columns = ['Date', '1WEEK', '1MONTH', '2MONTH', '3MONTH', '4MONTH', '5MONTH', '6MONTH', '7MONTH', '8MONTH', '9MONTH', '10MONTH', '11MONTH']
        
        # 5. 删除无效行（前两行通常是表头/空行，根据实际数据调整）
        df = df.drop([0, 1], errors='ignore').reset_index(drop=True)
        print(f"\n删除前2行后形状：行数={df.shape[0]}, 列数={df.shape[1]}")
        
        # 6. 处理换行单元格（如日期分行）
        df = split_row_to_rows(df)
        print(f"处理换行后形状：行数={df.shape[0]}, 列数={df.shape[1]}")
        
        # 7. 日期列处理
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')  # 无效日期转为NaT
        df = df.dropna(subset=['Date']).reset_index(drop=True)  # 删除无有效日期的行
        df.set_index('Date', inplace=True)
        df.index.rename('date', inplace=True)
        
        # 8. 数据清洗（NaN填充为0，避免后续计算错误）
        df_numeric = df.fillna(0)
        print("\n最终数据预览：")
        print(df_numeric.head())
        
        # 9. 构建邮件内容
        html_table = df.fillna('').to_html(index=True, border=0)  # 空值显示为空字符串
        body = f"<p>PDF下载链接：<a href='{pdf_url}' target='_blank'>点击下载</a></p><br/><div>{html_table}</div><br/>"
        
        # 10. 检测变化超0.1%的列，添加警告
        change_list = calculate_change(df_numeric)
        if change_list:
            change_msg = ", ".join(change_list) + " 变动超过0.1%"
            body = f"<h3 style='color:red;'>注意：{change_msg}</h3><br/>" + body
        
        # 11. 生成图表（至少5行数据才绘图）
        chart_data = None
        if len(df_numeric) >= 5:
            chart_data = create_line_chart(df_numeric)
            print("图表生成成功")
        else:
            print("数据行数不足5行，不生成图表")
        
        # 12. 发送邮件（替换为你的邮箱配置）
        sender_email = "chengguoyu_82@163.com"
        sender_password = "DUigKtCtMXw34MnB"  # 注意：163邮箱需用"授权码"，不是登录密码！
        recipient_emails = ["zling@jenseninvest.com","hwang@jenseninvest.com", "yqguo@jenseninvest.com", "13889632722@163.com"]
        subject = f"Japanese Yen TIBOR - {datetime.now().strftime('%Y-%m-%d')}"
        
        send_email(sender_email, sender_password, recipient_emails, subject, body, chart_data)
        
    except Exception as e:
        print(f"\n运行时错误：{str(e)}")
        # 可选：错误时发送告警邮件
        # send_email(sender_email, sender_password, ["你的邮箱"], "TIBOR程序报错", f"错误信息：{str(e)}")
