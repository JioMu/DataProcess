import pandas as pd
import os
import re

df = pd.read_excel(r'C:\Users\Administrator\Desktop\西南油气田数据\数据整理\中国分省市汽油产量及预测.xlsx', engine='openpyxl', header=None)
df.columns = [str(col) for col in df.columns]
unit = re.sub("单位：", "", df.iloc[1, 0])
print(unit)
datasource = re.sub("[资料来源：。]", "", df.iloc[36, 0])
print(datasource)
df = df.iloc[2:35, :]
df.columns = df.iloc[0]
df = df.drop(df.index[0])
result = df.melt(
    id_vars=['省份'],
    value_vars=df.columns[1:],
    var_name='年份',
    value_name='汽油产量'
).sort_values(['省份', '年份'], ascending=[True, True]).assign(单位=unit, 数据来源=datasource, 国家='中国')
result['年份'] = [re.sub(r"\D", "", str(col)) for col in result['年份']]
result['汽油产量'] = pd.to_numeric(result['汽油产量'], errors='coerce')
print(result)
os.makedirs('../处理后数据', exist_ok=True)
output_path = '../处理后数据/产量/中国分省市汽油产量及预测处理后数据.xlsx'
with pd.ExcelWriter(output_path) as writer:
    result.to_excel(writer, sheet_name='中国分省市汽油产量及预测处理后数据', index=False)
    print('数据处理完成')
