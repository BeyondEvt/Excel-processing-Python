import os
import xlwings as xw
app = xw.App(visible=True, add_book=False)

for file in os.listdir("."):
    if file.endswith(".xlsx"):
        app.books.open(file)
