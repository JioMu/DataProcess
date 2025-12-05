import pandas as pd
import os
import re

df = df = pd.read_excel(r'C:\Users\Administrator\Desktop\西南油气田数据\数据整理\中国煤油进出口量.xlsx', engine='openpyxl',
                        header=None)
unit = re.sub('单位：', '', df.iloc[1, 0])
print(unit)
datasource = re.sub('[资料来源：。]', '', df.iloc[26, 0])
print(datasource)
country = "中国"
df = df.iloc[2:25, :]
df.columns = df.iloc[0]
df.drop(df.index[0], inplace=True)
df['进口'] = pd.to_numeric(df['进口'], errors='coerce')
df['出口'] = pd.to_numeric(df['出口'], errors='coerce')
df = df.assign(国家=country, 单位=unit, 数据来源=datasource)
print(df)
os.makedirs("../处理后数据/进出口", exist_ok=True)
output_path = "../处理后数据/进出口/中国煤油进出口量处理后数据.xlsx"
with pd.ExcelWriter(output_path) as writer:
    df.to_excel(writer, sheet_name='中国煤油进出口量', index=False)
    print('数据处理完成')
