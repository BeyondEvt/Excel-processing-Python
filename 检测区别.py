import xlwings as xw

app = xw.App(visible=True, add_book=False)
books = app.books.open("产品统计表2.xlsx")
books_backup = app.books.open("产品统计表.xlsx")

for row in books.sheets[0].range("A1").expand():
    for cell in row:
        backup_cell = books_backup.sheets[0].range(cell.address)
        if cell.value != backup_cell.value:
            cell.color = backup_cell.color = (255, 0, 0)

books.save()
books.close()
books_backup.save()
books_backup.close()
app.quit()
