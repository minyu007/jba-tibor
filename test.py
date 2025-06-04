import requests
import pandas as pd
from datetime import datetime
import tabula
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import logging
import os
import sys
from retry import retry

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler('japan_yen_tibor.log'), logging.StreamHandler()]
)
logging.getLogger("org.apache.fontbox").setLevel(logging.ERROR)

def check_file_exists(filename):
    """检查文件是否存在"""
    return os.path.exists(filename)

@retry(tries=3, delay=2, backoff=2)
def save_file(url, filename):
    """下载文件并添加重试机制"""
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()  # 检查HTTP状态码
        with open(filename, 'wb') as f:
            f.write(response.content)
        logging.info(f"文件 '{filename}' 下载成功")
        return True
    except requests.exceptions.RequestException as e:
        logging.error(f"下载失败: {e}")
        raise

def parse_pdf(filename):
    """解析PDF文件并提取表格数据"""
    try:
        # 尝试自动解析
        tables = tabula.read_pdf(
            filename,
            pages="all",
            stream=True,  # 流式解析适合连续表格
            multiple_tables=True
        )
        
        # 如果自动解析失败，尝试指定区域
        if not tables:
            logging.warning("自动解析失败，尝试指定区域解析")
            tables = tabula.read_pdf(
                filename,
                pages="1",  # 只解析第一页
                area=(60, 50, 750, 650),  # 根据实际PDF调整区域
                stream=True,
                guess=False
            )
            
        if not tables:
            raise ValueError("无法从PDF中提取表格数据")
            
        # 合并所有表格
        dfs = [pd.DataFrame(table) for table in tables if not table.empty]
        if not dfs:
            raise ValueError("提取的表格数据为空")
            
        df = pd.concat(dfs, ignore_index=True)
        return df
        
    except Exception as e:
        logging.error(f"PDF解析失败: {e}")
        raise

def send_email(sender_email, sender_password, recipient_emails, subject, body, attachments=None):
    """发送邮件并添加附件支持"""
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = ", ".join(recipient_emails)
        msg['Subject'] = subject
        
        # 简化CSS样式以提高兼容性
        css = """
        <style>
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: center; }
        th { background-color: #f2f2f2; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        </style>
        """
        msg.attach(MIMEText(css + body, 'html'))
        
        # 添加附件
        if attachments:
            for attachment in attachments:
                if os.path.exists(attachment):
                    with open(attachment, 'rb') as file:
                        part = MIMEApplication(file.read(), Name=os.path.basename(attachment))
                        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment)}"'
                        msg.attach(part)
                else:
                    logging.warning(f"附件不存在: {attachment}")
        
        # 连接SMTP服务器并发送邮件
        with smtplib.SMTP_SSL('smtp.163.com', 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_emails, msg.as_string())
        logging.info("邮件发送成功")
        
    except smtplib.SMTPAuthenticationError:
        logging.error("SMTP认证失败，请检查邮箱账号或授权码")
        raise
    except Exception as e:
        logging.error(f"邮件发送失败: {e}")
        raise

def calculate_change(df):
    """计算变化超过阈值的列"""
    change_list = []
    if len(df) < 2:
        return change_list
        
    # 检查数据类型并转换为数值
    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        except Exception:
            continue
            
    # 计算变化
    for col in df.columns:
        try:
            if pd.api.types.is_numeric_dtype(df[col]):
                change = df[col].iloc[0] - df[col].iloc[1]
                if abs(change) > 0.001:  # 阈值0.1%
                    change_list.append(f"{col} (变化: {change:.4f})")
        except (IndexError, TypeError):
            continue
            
    return change_list

def main():
    """主函数"""
    try:
        # 获取当前日期并生成文件名
        current_date = datetime.now().strftime("%y%m%d")
        filename = f"{current_date}.pdf"
        pdf_url = f"https://www.jbatibor.or.jp/rate/pdf/JAPANESEYENTIBOR{current_date}.pdf"
        
        # 检查文件是否已存在
        if check_file_exists(filename):
            logging.info(f"文件 '{filename}' 已存在，跳过下载")
        else:
            logging.info(f"开始下载文件: {pdf_url}")
            save_file(pdf_url, filename)
            
        # 解析PDF
        logging.info("开始解析PDF文件")
        df = parse_pdf(filename)
        
        # 数据处理
        if df.empty:
            raise ValueError("解析后的数据为空")
            
        # 设置索引并转换为HTML
        df.set_index(df.columns[0], inplace=True)
        df.index.rename('date', inplace=True)
        html_table = df.fillna('nan').to_html(border=1)
        df.fillna(0, inplace=True)
        
        # 计算变化
        change_list = calculate_change(df)
        # "zling@jenseninvest.com",
        # "hwang@jenseninvest.com",
        # "yqguo@jenseninvest.com",
        # "13889632722@163.com"
        # 邮件配置
        sender_email = "chengguoyu_82@163.com"
        sender_password = "DUigKtCtMXw34MnB"
        recipient_emails = [
           "wo_oplove@163.com"
        ]
        subject = f"Japanese Yen TIBOR ({current_date})"
        
        # 构建邮件正文
        body = f"<p>Download PDF <a href='{pdf_url}' target='_blank'>click me!</a></p><br/><div>{html_table}</div><br/>"
        
        # 添加变化警告
        if change_list:
            change_message = ", ".join(change_list)
            body = f"<h3><font color='red'><b>注意：以下项目变化超过0.1%：{change_message}</b></font></h3><br/>" + body
        
        # 发送邮件
        logging.info("准备发送邮件")
        send_email(sender_email, sender_password, recipient_emails, subject, body, [filename])
        
        logging.info("程序执行成功")
        return 0
        
    except Exception as e:
        logging.error(f"程序执行失败: {e}", exc_info=True)
        print(f"运行时错误: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
