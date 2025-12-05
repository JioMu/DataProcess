from docx import Document
import pandas as pd
import re
import os


def extract_interface_test_cases(docx_path):
    doc = Document(docx_path)
    test_cases = []

    # 遍历所有表格
    for table in doc.tables:
        # 判断是否为接口测试用例表：包含“接口地址”或“请求方式”等关键词
        is_interface_case = False
        for row in table.rows:
            cells_text = [cell.text.strip() for cell in row.cells]
            if any(kw in ''.join(cells_text) for kw in ["接口地址", "请求方式", "用例名称"]):
                is_interface_case = True
                break

        if not is_interface_case:
            continue

        # 解析键值对
        case_dict = {}
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if len(cells) >= 2:
                key = re.sub(r'[：:]', '', cells[0]).strip()  # 去掉中文/英文冒号
                value = cells[1].strip()
                # 合并重复字段（如有）
                if key in case_dict:
                    case_dict[key] += "；" + value
                else:
                    case_dict[key] = value

        # 只保留非空且有意义的用例
        if case_dict.get("用例名称"):
            test_cases.append(case_dict)

    return test_cases


def generate_test_step(row):
    method = row.get("请求方式", "").strip()
    url = row.get("接口地址", "").strip()
    headers = row.get("请求头部", "").strip() or row.get("请求头", "").strip()
    params = row.get("请求参数", "").strip()

    # 构造测试步骤字符串
    step = f"发送{method}请求至{url}"
    if headers:
        step += f"，请求头：{headers}"
    if params:
        step += f"，请求体：{params}"
    return step


# 主程序
if __name__ == "__main__":
    docx_file = r"F:\金鼎二期二阶段验收材料\【参考】上期验收文档\11接口测试报告\金融数据智能分析和展示平台二期二阶段-接口测试报告.docx"

    if not os.path.exists(docx_file):
        print(f"❌ 文件 {docx_file} 不存在，请确认路径。")
        exit(1)

    cases = extract_interface_test_cases(docx_file)

    if not cases:
        print("⚠️ 未找到接口测试用例表格。")
    else:
        # 标准化字段名（兼容不同写法）
        standardized_cases = []
        for case in cases:
            new_case = {}
            # 字段映射（处理可能的别名）
            field_map = {
                "用例名称": ["用例名称"],
                "用例编号": ["用例编号"],
                "接口地址": ["接口地址"],
                "请求方式": ["请求方式"],
                "请求头部": ["请求头部", "请求头"],
                "请求参数": ["请求参数"],
                "状态码": ["状态码"],
                "预期返回结果": ["预期返回结果"],
                "实际结果": ["实际结果"],
                "前置条件": ["前置条件"],
                "描述": ["描述"]
            }
            for std_key, possible_keys in field_map.items():
                for k in possible_keys:
                    if k in case:
                        new_case[std_key] = case[k]
                        break
                else:
                    new_case[std_key] = ""
            standardized_cases.append(new_case)

        # 转为 DataFrame
        df = pd.DataFrame(standardized_cases)

        # 生成“测试步骤”列
        df["测试步骤"] = df.apply(generate_test_step, axis=1)

        # 调整列顺序（把测试步骤放前面）
        cols = ["用例名称", "用例编号", "测试步骤", "状态码", "预期返回结果", "实际结果", "前置条件", "描述",
                "接口地址", "请求方式", "请求头部", "请求参数"]
        existing_cols = [col for col in cols if col in df.columns]
        other_cols = [col for col in df.columns if col not in existing_cols]
        df = df[existing_cols + other_cols]

        # 导出 Excel
        output = "接口测试用例_导出.xlsx"
        df.to_excel(output, index=False, engine='openpyxl')
        print(f"✅ 成功导出 {len(cases)} 条接口测试用例到 {output}")