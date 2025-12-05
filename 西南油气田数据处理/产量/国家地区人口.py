import pandas as pd
import os

# 读取Excel文件, 从第三行开始读取数据, 跳过了标题
df = pd.read_excel(r'C:\Users\Administrator\Desktop\西南油气田数据\数据整理\世界主要地区和国家人口.xlsx', engine='openpyxl', header=2)
df.columns = [str(col).strip() for col in df.columns]
df = df.iloc[:72, :]
print(df)
years = df.columns[1:]  # 假设第一列是“国家/地区”，年份列从第二列开始
print("年份列名:\n", years)

result = df.melt(
    id_vars=['国家/地区'],
    value_vars=years,
    var_name='年份',
    value_name='数值'
).sort_values(by=['国家/地区', '年份'], ascending=[True, True])
print(result)
# 规范数据格式
result['年份'] = result['年份'].astype(int)
# 将数值列转换为数值类型, 忽略无法转换的行, 并使用coerce选项, 将无法转换的数据默认为NaN
result['数值'] = pd.to_numeric(result['数值'], errors='coerce')
print(result['年份'])

# 创建输出目录并保存
os.makedirs('../处理后数据', exist_ok=True)
output_path = '../处理后数据/产量/国家地区人口处理后数据.xlsx'
with pd.ExcelWriter(output_path) as writer:
    result.to_excel(writer, sheet_name='Sheet1', index=False)
    # 添加数据摘要表
    pd.DataFrame({
        '统计量': ['最大值', '最小值', '平均值'],
        '数值': [
            result['数值'].max(),
            result['数值'].min(),
            result['数值'].mean()
        ]
    }).to_excel(writer, sheet_name='数据摘要', index=False)

print("数据整合完成,文件已保存至: output_path" + output_path)

