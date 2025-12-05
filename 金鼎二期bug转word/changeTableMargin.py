from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from docx.enum.table import WD_TABLE_ALIGNMENT


def process_single_table(table):
    """处理单个表格的宽度、缩进、对齐（核心逻辑）"""
    # ========== 1. 强制关闭「自动调整」（防止Word覆盖宽度） ==========
    table.autofit = False

    # ========== 2. 高层API设置宽度（97.5%）- 兼容所有版本 ==========
    table.preferred_width = 97.5  # 直接设置百分比数值
    try:
        from docx.enum.table import WD_PREFERRED_WIDTH_TYPE
        table.preferred_width_type = WD_PREFERRED_WIDTH_TYPE.PERCENT
    except ImportError:
        pass  # 旧版本无此枚举，不影响，底层XML兜底

    # ========== 3. 底层XML：先删旧宽度，再强制设新值 ==========
    tbl = table._tbl  # 表格XML根元素
    tblPr = tbl.tblPr  # 表格属性容器

    # ❶ 删除旧的<w:tblW>（彻底清除历史属性）
    old_tblW = tblPr.xpath('w:tblW')
    for elem in old_tblW:
        elem.getparent().remove(elem)

    # ❷ 新建<w:tblW>，强制设97.5%（千分比9750）
    new_tblW = OxmlElement('w:tblW')
    new_tblW.set(qn('w:w'), '9750')  # 97.5% = 9750（100%=10000）
    new_tblW.set(qn('w:type'), 'pct')  # 单位：百分比
    tblPr.append(new_tblW)

    # ========== 4. 左缩进：先删旧属性，再设0.2厘米 ==========
    # ❶ 删除旧的<w:tblInd>
    old_tblInd = tblPr.xpath('w:tblInd')
    for elem in old_tblInd:
        elem.getparent().remove(elem)

    # ❷ 新建<w:tblInd>，0.2厘米→113缇
    new_tblInd = OxmlElement('w:tblInd')
    indent_twips = int(Cm(0.2).twips)  # 0.2cm = 113.4缇 → 取整113
    new_tblInd.set(qn('w:w'), str(indent_twips))
    new_tblInd.set(qn('w:type'), 'dxa')  # 单位：缇（绝对长度）
    tblPr.append(new_tblInd)

    # ========== 5. 确保表格左对齐（缩进生效前提） ==========
    table.alignment = WD_TABLE_ALIGNMENT.LEFT


def set_table_styles(doc_path, output_path):
    """处理文档中所有顶层表格（无嵌套表格错误）"""
    doc = Document(doc_path)

    # 遍历所有顶层表格（`python-docx`不支持嵌套表格的Shape遍历，故仅处理顶层）
    for table in doc.tables:
        process_single_table(table)

    doc.save(output_path)
    print(f"表格样式修改完成！保存至：{output_path}")


# ========== 执行测试 ==========
if __name__ == "__main__":
    input_doc = r"C:\Users\Administrator\Desktop\数智公卫 -数据资源中心验收资料  - 副本 (2).docx"
    output_doc = r"C:\Users\Administrator\Desktop\样式统一_数智公卫验收资料.docx"
    set_table_styles(input_doc, output_doc)