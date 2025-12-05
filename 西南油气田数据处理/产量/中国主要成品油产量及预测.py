import pandas as pd
import os
import re

df = pd.read_excel(r'C:\Users\Administrator\Desktop\西南油气田数据\数据整理\中国主要成品油产量及预测.xlsx', engine='openpyxl', header=2)
unit = "万吨"
datasource = re.sub('[数据来源：。]', '', df.iloc[26, 0])
print(datasource)
country = "中国"
df = df.drop(df.index[0])
df = df.assign(国家=country, 单位=unit, 数据来源=datasource)
df = df.iloc[:23, :]
df["年 份"] = [re.sub(r"\D", "", str(col)) for col in df["年 份"]]
print(df)
os.makedirs('../处理后数据', exist_ok=True)
output_path = '../处理后数据/产量/中国主要成品油产量及预测处理后数据.xlsx'
with pd.ExcelWriter(output_path) as writer:
    df.to_excel(writer, sheet_name='sheet1', index=False)
    print('数据处理完成')
