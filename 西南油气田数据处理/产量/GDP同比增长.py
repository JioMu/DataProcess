import pandas as pd
import re
import os

old_df = pd.read_excel(r'C:\Users\Administrator\Desktop\西南油气田数据\数据整理\GDP_实际同比增长_G20.xlsx', engine='openpyxl')
old_df.columns = [str(col).strip() for col in old_df.columns]
df = old_df.iloc[0:42, 0:]
countryList = df.columns[1:]
result = df.melt(
    id_vars=['指标名称'],
    value_vars=countryList,
    var_name='国家',
    value_name='GDP同比增长'
).sort_values(by=['国家', '指标名称'], ascending=[True, True])
for country in result['国家']:
    result['国家'] = result['国家'].replace(country, country.replace('GDP:实际同比增长:', ''))
result['指标名称'] = result['指标名称'].dt.strftime('%Y')
print(result)
result.rename(columns={'指标名称': '年份'}, inplace=True)
result['年份'] = result['年份'].astype(int)
result['GDP同比增长'] = pd.to_numeric(result['GDP同比增长'], errors='coerce')
os.makedirs('../处理后数据', exist_ok=True)
out_path = '../处理后数据/产量/GDP_实际同比增长_G20处理后数据.xlsx'
with pd.ExcelWriter(out_path) as writer:
    result.to_excel(writer, sheet_name='sheet1', index=False)
print("数据处理完成,文件输出路径", out_path)
