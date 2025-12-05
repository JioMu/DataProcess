import pandas as pd
import re
import os

# 读取文件
df = pd.read_excel(r'C:\Users\Administrator\Desktop\西南油气田数据\数据整理\天然气开发---世界分地区和主要国家天然气产量.xlsx', engine='openpyxl',
                   header=2)

df.columns = [str(col).strip() for col in df.columns]
years = [re.sub(r'\D', '', str(col)) for col in df.columns[1:]]

result = df.melt(
    id_vars=['国家/地区'],
    value_vars=years,
    var_name='年份',
    value_name='数值'
).sort_values(by=['国家/地区', '年份'])

result['年份'] = result['年份'].astype(int)
result['数值'] = pd.to_numeric(result['数值'], errors='coerce')

os.makedirs('../处理后数据', exist_ok=True)
out_path = '../处理后数据/产量/天然气开发---世界分地区和主要国家天然气产量处理后数据.xlsx'
with pd.ExcelWriter(out_path) as writer:
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
        }).to_excel(writer, sheet_name='数据摘要', index=False)

print("数据处理完成,文件输出路径\n"+out_path)
