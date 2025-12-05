import pandas as pd
import os
import re

df = pd.read_excel(r'C:\Users\Administrator\Desktop\西南油气田数据\数据整理\天然气进出口----世界主要地区和国家天然气进口量.xlsx', engine='openpyxl', header=None)
unit = re.sub("单位：", "", df.iloc[1, 0])
print(unit)
datasource = re.sub('[资料来源：。]', '', df.iloc[75, 0])
print(datasource)
df = df.iloc[2:75, :]
df.columns = df.iloc[0]
df.drop(df.index[0], inplace=True)
years = df.columns[1:]
print(years)
result = df.melt(
    id_vars=['国家/地区'],
    value_vars=years,
    var_name='年份',
    value_name='天然气进口量'
).sort_values(['国家/地区', '年份'], ascending=[True, True]).assign(
    单位=unit,
    数据来源=datasource
)
result['年份'] = [re.sub(r'\D', '', str(col)) for col in result['年份']]
result['天然气进口量'] = pd.to_numeric(result['天然气进口量'], errors='coerce')
print(result)
os.makedirs('../处理后数据/进出口', exist_ok=True)
output_path = '../处理后数据/进出口/天然气进出口----世界主要地区和国家天然气进口量处理后数据.xlsx'
with pd.ExcelWriter(output_path) as writer:
    result.to_excel(writer, sheet_name='sheet1', index=False)
    print('数据处理完成')

