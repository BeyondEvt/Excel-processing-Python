import pandas as pd
import xlwings as xw

df_list = pd.read_excel("游戏产品表.xlsx",sheet_name=None)
df_all = pd.concat(df_list.values())

app = xw.App(visible=False, add_book =False)
workbook = app.books.open("游戏产品表.xlsx")
workbook.sheets.add("汇总表", before= workbook.sheets[0])
workbook.sheets["汇总表"].range("A1").options(index=False).value=df_all

workbook.save()
workbook.close()
app.quit()