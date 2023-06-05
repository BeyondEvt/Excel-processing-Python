import xlwings as xw
import os
import pandas as pd

app = xw.App(visible=True, add_book=False)
workbook = app.books.open("学生成绩表.xlsx")

df_total = workbook.sheets[0].range("A1").options(pd.DataFrame
    , expand="table", index=False, numbers=int).value
print(df_total)

df_student_list = []

for file_name in os.listdir("学生信息"):
    if "信息" not in file_name:
        continue
    workbook_student = app.books.open("学生信息/"+file_name)

    df_student  = workbook_student.sheets[0].range("A1").options(
        pd.DataFrame, expand="table", index=False, numbers=int
    ).value
    df_student["班级"] = file_name.replace("信息.xlsx","")
    df_student_list.append(df_student)

    workbook_student.close()
df_student_all = pd.concat(df_student_list)

# 数据合并
df_merge = pd.merge(
    left=df_total,
    right = df_student_all,
    left_on=["班级","姓名"],
    right_on=["班级","姓名"]
)

df_merge["电话号码"] = df_merge["电话"]
df_merge.drop(columns="电话", inplace=True)

workbook.sheets[0].range("A1").options(index=False).value = df_merge
workbook.save()
workbook.close()