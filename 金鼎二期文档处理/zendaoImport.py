from docx import Document
import pandas as pd
import re

def extract_test_cases_from_docx(docx_path):
    doc = Document(docx_path)
    test_cases = []

    # 遍历文档中的所有表格
    for table in doc.tables:
        # 判断是否是测试用例表：检查第一行或任意单元格是否包含“测试编号”
        is_test_case_table = False
        headers = []
        rows = []

        for i, row in enumerate(table.rows):
            cells_text = [cell.text.strip() for cell in row.cells]
            # 如果某一行包含“测试编号”，则认为是测试用例表
            if any("测试编号" in text for text in cells_text):
                is_test_case_table = True
                # 假设这一行是表头（键值对形式）
                # 实际可能是两列：字段名 | 值
                # 我们将按两列解析
                break

        if not is_test_case_table:
            continue

        # 解析该表格为键值对
        case_dict = {}
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if len(cells) >= 2:
                key = cells[0].replace(' ', '').replace('｜', '|').strip()
                value = cells[1].strip()
                # 合并重复字段（如有）
                if key in case_dict:
                    case_dict[key] += "；" + value
                else:
                    case_dict[key] = value

        if case_dict:
            test_cases.append(case_dict)

    return test_cases

# 主程序
if __name__ == "__main__":
    docx_file = r"F:\金鼎二期二阶段验收材料\【参考】上期验收文档\09功能测试报告\金融数据智能分析和展示平台二期二阶段-功能测试报告.docx"
    test_cases = extract_test_cases_from_docx(docx_file)

    if not test_cases:
        print("未找到符合的测试用例表格。")
    else:
        # 获取所有字段名（列名）
        all_keys = set()
        for case in test_cases:
            all_keys.update(case.keys())
        all_keys = sorted(all_keys)

        # 补齐缺失字段
        df_data = []
        for case in test_cases:
            row = {key: case.get(key, "") for key in all_keys}
            df_data.append(row)

        df = pd.DataFrame(df_data, columns=all_keys)
        output_excel = "测试用例_导出.xlsx"
        df.to_excel(output_excel, index=False)
        print(f"成功导出 {len(test_cases)} 条测试用例到 {output_excel}")