import pandas as pd
import re
import os
df = pd.read_excel(r'C:\Users\Administrator\Desktop\西南油气田数据\数据整理\中国分省市煤油产量及预测.xlsx', engine='openpyxl', header=2)
years = df.columns[1:]
unit = "万吨"
datasource = df.iloc[33, 0]
datasource = re.sub('[资料来源：。]', '', datasource)
df = df.iloc[:32, :]
print(df)
result = df.melt(
    id_vars=['省份'],
    value_vars=years,
    var_name='年份',
    value_name='煤油产量（万吨）'
).assign(单位=unit, 来源=datasource, 国家='中国').sort_values(['省份', '年份'], ascending=[True, True])
result['年份'] = [re.sub(r'\D', '', str(col)) for col in result['年份']]
result['煤油产量（万吨）'] = pd.to_numeric(result['煤油产量（万吨）'], errors='coerce')
print(result)
os.makedirs('../处理后数据', exist_ok=True)
out_path = '../处理后数据/产量/中国分省市煤油产量及预测处理后数据.xlsx'
with pd.ExcelWriter(out_path) as writer:
    result.to_excel(writer, index=False)
    print('写入完成')
