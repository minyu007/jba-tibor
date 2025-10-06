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
import matplotlib.dates as mdates  # 添加这行导入
import io
import os

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


def create_line_chart(df):
    """Create a line chart from the DataFrame and return it as a bytes object"""
    plt.figure(figsize=(12, 6))
    
    # 筛选需要绘制的列
    plot_columns = [
        col for col in df.columns 
        if (pd.api.types.is_numeric_dtype(df[col]) and 
            not all(df[col].fillna(0) == 0))
    ]
    
    if not plot_columns:
        return None
    
    # 确保索引是datetime类型并按日期升序排列
    if not isinstance(df.index, pd.DatetimeIndex):
        try:
            df.index = pd.to_datetime(df.index)
        except Exception as e:
            print(f"日期转换错误: {e}")
            return None
    
    # 确保数据按日期升序排列（从左到右时间递增）
    df = df.sort_index()
    
    # 创建图表
    for column in plot_columns:
        line = plt.plot(df.index, df[column], marker='o', label=column)
        
        # 计算变化并标注（从左到右）
        if len(df) >= 2:  # 至少需要两个点才能计算变化
            changes = df[column].diff()  # 计算与前一个点的差值
            for i in range(1, len(df)):  # 从第二个点开始检查
                change = changes.iloc[i]
                if abs(change) > 0.001:  # 变化超过0.1%
                    date = df.index[i]
                    y_val = df[column].iloc[i]
                    prev_val = df[column].iloc[i-1]
                    change_pct = (change / prev_val) * 100  # 计算变化百分比
                    
                    # 确定箭头方向和颜色
                    arrow_direction = '↑' if change > 0 else '↓'
                    arrow_color = 'red' if change > 0 else 'blue'
                    bg_color = 'lightcoral' if change > 0 else 'lightblue'
                    
                    # 添加标注
                    plt.annotate(f'{arrow_direction}{abs(change_pct):.2f}%', 
                                xy=(date, y_val),
                                xytext=(0, 15 if change > 0 else -15),
                                textcoords='offset points',
                                ha='center',
                                va='center',
                                bbox=dict(boxstyle='round,pad=0.5', 
                                         fc=bg_color, 
                                         alpha=0.8),
                                arrowprops=dict(arrowstyle='->', 
                                               color=arrow_color,
                                               linewidth=1.5))
    
    plt.title('Japanese Yen TIBOR Rates with Daily Changes')
    plt.ylabel('Rate (%)')
    plt.xlabel('Date')
    
    # 格式化x轴
    ax = plt.gca()
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    ax.xaxis.set_major_locator(mdates.AutoDateLocator())
    plt.xticks(rotation=45, ha='right')
    
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.grid(True)
    plt.tight_layout()
    
    # 保存图表
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
    buf.seek(0)
    plt.close()
    
    return buf

def send_email(sender_email, sender_password, recipient_emails, subject, body, chart_data=None, attachments=None):
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
    
    # Add chart if provided
    if chart_data:
        chart_html = '<h3>TIBOR Rates Trend</h3><img src="cid:chart">'
        body = chart_html + body
    
    msg.attach(MIMEText(css + body, 'html'))
    
    # Attach chart image
    if chart_data:
        image = MIMEImage(chart_data.read())
        image.add_header('Content-ID', '<chart>')
        msg.attach(image)
    
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

def split_row_to_rows(df):
    """
    遍历DataFrame的每一行，将包含换行符(\r)的单元格拆分成多行。
    """
    if df.empty:
        return pd.DataFrame()
    
    # 用于存储所有拆分后行的列表
    all_split_rows = []
    
    # 遍历原始DataFrame的每一行
    for index, row in df.iterrows():
        row = row.replace({np.nan: ''})
        split_data = {}
        
        # 处理当前行的每一列
        for col in row.index:
            cell_value = row[col]
            
            if pd.api.types.is_numeric_dtype(type(cell_value)):
                split_data[col] = [cell_value]
            else:
                # 确保是字符串再进行 split
                split_values = [v.strip() for v in str(cell_value).split('\r')]
                split_data[col] = split_values
        
        max_length = max(len(v) for v in split_data.values())
        
        # 统一长度
        for col in split_data:
            current_length = len(split_data[col])
            if current_length < max_length:
                # 使用np.nan填充，而不是空字符串，方便后续处理
                split_data[col].extend([np.nan] * (max_length - current_length))
        
        # 将当前行拆分后的结果转换为DataFrame并添加到列表中
        split_df = pd.DataFrame(split_data)
        all_split_rows.append(split_df)
    
    # 将所有拆分后的DataFrame合并成一个
    if not all_split_rows:
        return pd.DataFrame()
        
    new_df = pd.concat(all_split_rows, ignore_index=True)
    
    # 尝试将所有非'Date'列转换为数值类型
    for col in new_df.columns:
        if col != 'Date':
            new_df[col] = pd.to_numeric(new_df[col], errors='coerce')
    
    return new_df

if __name__ == "__main__":
    if not check_file_exists():
        try:
            # current_date = datetime.now().strftime("%y%m%d")
            current_date='250924'
            pdf_url = f"https://www.jbatibor.or.jp/rate/pdf/JAPANESEYENTIBOR{current_date}.pdf"
            filename = f"{current_date}.pdf"
            
            save_file(pdf_url, filename)
            
            tables = tabula.read_pdf(
                filename,
                pages="all"
            )
            
            dfs = [pd.DataFrame(table) for table in tables]
            print(dfs)
            # print(df)
            df = pd.concat(dfs, ignore_index=True)
            print(df.columns)
            print(df)
            df.columns=['Date',
                '1WEEK',
                '1MONTH',
                '2MONTH',
                '3MONTH',
                '4MONTH',
                '5MONTH',
                '6MONTH',
                '7MONTH',
                '8MONTH',
                '9MONTH',
                '10MONTH',
                '11MONTH',
                '12MONTH'
               ]

            print(df.columns)
            print(df)
            # df = df.drop([0, 1])
            df = split_row_to_rows(df)

            # print(df.columns)
            # print(df)
            
            # Convert Date column to datetime if it's not already
            df['Date'] = pd.to_datetime(df['Date'])
            
            # Set Date as index after conversion
            df.set_index('Date', inplace=True)
            df.index.rename('date', inplace=True)
            
            html_table = df.fillna('').to_html(border=1)
            df.fillna(0, inplace=True)
            
            sender_email = "chengguoyu_82@163.com"
            sender_password = "DUigKtCtMXw34MnB"
            # recipient_emails = ["zling@jenseninvest.com","hwang@jenseninvest.com", "yqguo@jenseninvest.com", "13889632722@163.com"]
            recipient_emails = ["wo_oplove@163.com"]
            subject = "Japanese Yen TIBOR"
            
            body = f"<p>Download PDF <a href='{pdf_url}' target='_blank'>click me!</a></p><br/><div>{html_table}</div><br/>"
            
            change_list = calculate_change(df)
            if change_list:
                change_message = ", ".join(change_list) + " changed by more than 0.1%"
                body = f"**<h3><font color='red'><b>Please note that {change_message}</b></font></h3>**<br/>" + body
            
            # Create and attach chart if there are more than 4 rows
            chart_data = None
            if len(df) >= 5:
                chart_data = create_line_chart(df)
            
            send_email(sender_email, sender_password, recipient_emails, subject, body, chart_data)
            
            print(df)
        except Exception as e:
            print("运行时错误:", e)
    else:
        print("文件已存在，跳过程序执行。")
