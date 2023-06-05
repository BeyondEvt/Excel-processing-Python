import  pandas as pd
import os

data_list = []
for fname in os.listdir("."):
    if fname.startswith("产品-") and fname.endswith(".xlsx"):
        data_list.append(pd.read_excel(fname))

data_all = pd.concat(data_list)
data_all.to_excel("产品统计表.xlsx", index=False)