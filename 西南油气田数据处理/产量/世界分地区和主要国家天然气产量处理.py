# 正则表达式模块
import re
import pandas as pd
import os

# 读取Excel文件, 从第三行开始读取数据, 跳过了标题
df = pd.read_excel(r'C:\Users\Administrator\Desktop\西南油气田数据\数据整理\世界分地区和主要国家天然气产量.xlsx', engine='openpyxl', header=2)
df.columns = [str(col).strip() for col in df.columns]
datasource = df.iloc[72, 0]
datasource = re.sub('[资料来源：]', '', datasource)
# 选择前72行数据
df = df.iloc[:72, :]
# 用正则表达式的sub方法删除非数字字符
years = [re.sub(r'\D', '', str(col)) for col in df.columns[1:]]  # 假设第一列是“国家/地区”，年份列从第二列开始
result = df.melt(
    id_vars=['国家/地区'],
    value_vars=years,
    var_name='年份',
    value_name='数值'
).assign(单位="亿立方米", 来源=datasource).sort_values(by=['国家/地区', '年份'], ascending=[True, True])
# 规范数据格式
result['年份'] = result['年份'].astype(int)
# 将数值列转换为数值类型, 忽略无法转换的行, 并使用coerce选项, 将无法转换的数据默认为NaN
result['数值'] = pd.to_numeric(result['数值'], errors='coerce')
print(result)

# 创建输出目录并保存
os.makedirs('../处理后数据', exist_ok=True)
output_path = '../处理后数据/产量/世界分地区和主要国家天然气产量处理处理后数据.xlsx'
with pd.ExcelWriter(output_path) as writer:
    result.to_excel(writer, sheet_name='Sheet1', index=False)

print("数据整合完成,文件已保存至: output_path" + output_path)
