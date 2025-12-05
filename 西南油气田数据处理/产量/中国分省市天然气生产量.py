import pandas as pd
import os
import re

df = pd.read_excel(r'C:\Users\Administrator\Desktop\西南油气田数据\数据整理\中国分省市天然气生产量.xlsx', engine='openpyxl', header=2)
print(df)
datasource = df.iloc[31, 0]
datasource = re.sub('[资料来源：。]', '', datasource)
print(datasource)
df.columns = [str(col) for col in df.columns]
df = df.iloc[:31, :]
years = df.columns[1:]

result = df.melt(
    id_vars=['地区'],
    value_vars=years,
    var_name='年份',
    value_name='天然气产量'
).assign(单位='亿立方米', 国家='中国', 数据来源=datasource).sort_values(['地区', '年份'])
result['年份'] = [re.sub(r'\D', '', str(col)) for col in result['年份']]
result['年份'] = result['年份'].astype(int)
result['天然气产量'] = pd.to_numeric(result['天然气产量'], errors='coerce')
print(result)
# output_path = './处理后数据/中国分省市天然气生产量处理后数据.xlsx'
# os.makedirs('/处理后数据', exist_ok=True)
# with pd.ExcelWriter(output_path) as writer:
#     result.to_excel(writer, index=False)
#     print(f"数据处理完成，结果已保存到--->{output_path}")
