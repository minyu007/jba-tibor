import requests
import pandas as pd
from datetime import datetime
import tabula

current_date = datetime.now().strftime("%y%m%d")

base_url = "https://www.jbatibor.or.jp/rate/pdf/JAPANESEYENTIBOR{}.pdf"

pdf_url = base_url.format(current_date)

tables = tabula.read_pdf(pdf_url, pages="all")

dfs = [pd.DataFrame(table) for table in tables]

final_df = pd.concat(dfs, ignore_index=True)

# excel_filename = "output.xlsx"
# final_df.to_excel(excel_filename, index=False)

print(final_df)
