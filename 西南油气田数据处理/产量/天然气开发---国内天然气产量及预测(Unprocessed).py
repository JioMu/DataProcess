import pandas as pd
import re
import os

df = pd.read_excel(r'C:\Users\Administrator\Desktop\西南油气田数据\数据整理\天然气开发---国内天然气产量及预测.xlsx', engine='openpyxl', header=2)
unit = "亿立方米"
datasource = df.iloc[40, 1]
datasource = re.sub('[数据来源：。]', '', datasource)
df = df.iloc[:40, :]
df = df.assign(数据来源=datasource, 单位=unit)
print(df)
