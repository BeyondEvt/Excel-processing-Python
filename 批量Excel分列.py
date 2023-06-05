import xlwings as xw
import pandas as pd
import os

app = xw.App(visible=True, add_book=False)
directory = "产品记录表" # 要批量处理的文件所在目录

for fname in os.listdir(directory):
    if fname.endswith(".xlsx"):
        workbook = app.books.open(os.path.join(directory, fname))
        worksheet = workbook.sheets["产品规格表"]
        df = worksheet.range("A1").options(pd.DataFrame, expand="table").value
        split_columns = df["规格"].str.split("*", expand=True)
        df["长"] = split_columns[0]
        df["宽"] = split_columns[1]
        df["高"] = split_columns[2]
        worksheet.range("A1").value = df
        workbook.save()
app.quit()