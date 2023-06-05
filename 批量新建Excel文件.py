import xlwings as xw

app = xw.App(visible=True, add_book=False) # 可视，不添加book

for dept in ["技术部", "销售部", "运营部", "财务部", "人事部"]:
    workbook = app.books.add()
    workbook .save(f"./部门业绩-{dept}.xlsx")
