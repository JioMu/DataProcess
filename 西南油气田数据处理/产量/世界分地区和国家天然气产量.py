import pandas as pd
import re
import os

# 读取文件
df = pd.read_excel(r'C:\Users\Administrator\Desktop\西南油气田数据\数据整理\世界分地区和国家天然气产量.xlsx', engine='openpyxl')
# 去除首尾空格
df.columns = [str(col).strip() for col in df.columns]
print(df)
# 获取年份
years = df.columns[1:13]
# 获取单位
unit = df.iloc[1, 13]
print("单位", unit)
# 获取来源
source = df.iloc[1, 14]
print("数据来源", source)
result = df.melt(
    id_vars=['国家/地区'],
    value_vars=years,
    var_name='年份',
    value_name='数值'
).sort_values(by=['国家/地区', '年份']).assign(单位=unit, 来源=source)
# 规范数据格式
result['年份'] = result['年份'].astype(int)
result['数值'] = pd.to_numeric(result['数值'], errors='coerce')
print(result)
# 写入文件
output_path = r'../处理后数据/产量/世界分地区和国家天然气产量处理后数据.xlsx'
os.makedirs('../处理后数据', exist_ok=True)
with pd.ExcelWriter(output_path) as writer:
    result.to_excel(writer, sheet_name='Sheet1', index=False)
    pd.DataFrame(
        {
            '统计量': ['最大值', '最小值', '平均值'],
            '数值': [
                result['数值'].max(),
                result['数值'].min(),
                result['数值'].mean()
            ]
        }
    ).to_excel(writer, sheet_name='数据摘要', index=False)

print('数据处理完成,存储路径在' + output_path)
