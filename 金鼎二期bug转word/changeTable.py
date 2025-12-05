# r"C:\Users\Administrator\Desktop\处理测试方案表格.docx"

from docx import Document


# 加载文档
doc = Document(r"C:\Users\Administrator\Desktop\处理测试方案表格v3.0.docx")

# 遍历文档中的每个表格
for table in doc.tables:
    # 遍历表格的每一行
    for row_index, row in enumerate(table.rows):
        # 遍历每一行中的单元格
        for cell_index, cell in enumerate(row.cells):
            # 替换单元格内容
            if "期望结果与实际测试结果一致可正常终止执行测试用例" in cell.text:
                cell.text = '是'
            if "测试用例期望结果与实际结果一致，测试用例通过" in cell.text:
                cell.text = ''

# 保存修改后的文档
doc.save('处理测试方案表格_已修改.docx')


