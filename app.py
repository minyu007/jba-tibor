import requests
import pandas as pd
from datetime import datetime
import tabula

# 获取今天的日期，并将其格式化为字符串
current_date = datetime.now().strftime("%y%m%d")

# 定义动态生成 PDF 文件的 URL
base_url = "https://www.jbatibor.or.jp/rate/pdf/JAPANESEYENTIBOR{}.pdf"

# 构建完整的 URL
pdf_url = base_url.format(current_date)

# 使用 tabula 读取 PDF 文件中的表格数据
tables = tabula.read_pdf(pdf_url, pages="all")

# 将表格数据存储到 DataFrame 中
dfs = [pd.DataFrame(table) for table in tables]

# 合并所有 DataFrame
final_df = pd.concat(dfs, ignore_index=True)

# 将 DataFrame 写入 Excel 文件
excel_filename = "output.xlsx"
final_df.to_excel(excel_filename, index=False)

print("DataFrame 已保存到 Excel 文件:", excel_filename)
