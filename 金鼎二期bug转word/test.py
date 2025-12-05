from docx import Document
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn

doc = Document(r'C:\Users\Administrator\Downloads\9.金融数据智能分析和展示平台-功能测试报告v1.1.docx')


def set_table_border_black(table):
    tbl = table._tbl
    tbl_pr = tbl.tblPr

    # 获取或创建 <w:tblBorders>
    tbl_borders = tbl_pr.first_child_found_in('w:tblBorders')
    if not tbl_borders:
        tbl_borders = OxmlElement('w:tblBorders')
        tbl_pr.append(tbl_borders)

    borders = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']
    for border_type in borders:
        tag = f'w:{border_type}'
        elm = tbl_borders.find(qn(tag))
        if elm is None:
            elm = OxmlElement(tag)
            tbl_borders.append(elm)
        elm.set(qn('w:val'), 'single')
        elm.set(qn('w:sz'), '4')
        elm.set(qn('w:space'), '0')
        elm.set(qn('w:color'), '000000')


for table in doc.tables:
    table.style = None  # 避免样式冲突
    set_table_border_black(table)

doc.save('./修改后的文档.docx')
