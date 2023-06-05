import pandas as pd

df = pd.read_excel("产品统计表.xlsx")

products = df["产品名"].unique()
print(products)

for product in products:
    df_product = df[df["产品名"] == product]
    df_product.to_excel(f"产品统计表-{product}.xlsx")
