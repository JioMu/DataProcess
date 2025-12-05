import pandas as pd
import re
import os

df = pd.read_excel(r'C:\Users\Administrator\Desktop\西南油气田数据\数据整理\世界主要地区和国家煤油产量.xlsx', engine='openpyxl',
                   header=2)
df.columns = [str(col).strip() for col in df.columns]
datasource = df.iloc[72, 0]
datasource = re.sub('[资料来源：]', '', datasource).strip()
print(datasource)
df = df.iloc[:72, :]
years = [re.sub(r'\D', '', str(col)) for col in df.columns[1:]]
years = df.columns[1:]
result = df.melt(
    id_vars=['国家/地区'],
    value_vars=years,
    var_name='年份',
    value_name='数值'
).sort_values(by=['国家/地区', '年份']).assign(单位="万吨", 数据来源=datasource)
result['年份'] = [re.sub(r'\D', '', str(col)) for col in result['年份']]
result['年份'] = result['年份'].astype(int)
result['数值'] = pd.to_numeric(result['数值'], errors='coerce')
print(result)
os.makedirs('../处理后数据', exist_ok=True)
output_path = '../处理后数据/产量/世界主要地区和国家煤油产量处理后数据.xlsx'
with pd.ExcelWriter(output_path) as writer:
    result.to_excel(writer, sheet_name='sheet1', index=False)
    pd.DataFrame(
        {
            '统计量': [
                '最大值',
                '最小值',
                '平均值'
            ],
            '数值': [
                result['数值'].max(),
                result['数值'].min(),
                result['数值'].mean()
            ]
        }
    ).to_excel(writer, sheet_name='数据摘要', index=False)
print('数据处理完成,输出路径为:', output_path)
