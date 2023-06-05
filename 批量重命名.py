import xlwings as xw

app = xw.App(visible=True, add_book=False)

workbook = app.books.open("部门业绩-技术部.xlsx")
for sheet in workbook.sheets:
    sheet.name = sheet.name.replace("技术","")

workbook.save()
app.quit()

