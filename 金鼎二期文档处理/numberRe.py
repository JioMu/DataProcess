from docx import Document
import re


def renumber_test_cases(input_path, output_path, start_number=43385):
    doc = Document(input_path)
    current_number = start_number
    total_renumbered = 0

    # 遍历文档中的每一个表格
    for table in doc.tables:
        try:
            # 遍历表格中的每一行
            for row in table.rows:
                if len(row.cells) < 2:
                    continue

                first_cell_text = row.cells[0].text.strip()

                # 检查该行是否为“用例编号”行（支持模糊匹配）
                if re.search(r'用例编号|测试编号', first_cell_text, re.IGNORECASE):
                    # 更新第二列的值
                    second_cell = row.cells[1]
                    second_cell.text = str(current_number)

                    print(f"✅ 找到 '{first_cell_text}'，已更新为: {current_number}")
                    current_number += 1
                    total_renumbered += 1
                    break  # 假设每个表格只有一个用例编号，找到后跳出该表格

        except Exception as e:
            print(f"⚠️ 处理表格时出错: {e}")

    doc.save(output_path)
    print(f"\n🎉 处理完成！")
    print(f"📄 新文档已保存至: {output_path}")
    print(f"🔢 共重编号了 {total_renumbered} 个测试用例。")


if __name__ == "__main__":
    input_file = r"C:\Users\Captain\Downloads\金融数据智能分析和展示平台二期二阶段-接口测试报告.docx"
    output_file = "金融数据智能分析和展示平台二期二阶段-接口测试报告_重编号版.docx"
    renumber_test_cases(input_file, output_file, start_number=43385)