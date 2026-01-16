"""
创建测试模板文件

这个脚本创建一个包含模板变量的 Word 文档用于测试。
"""

from spire.doc import *
from spire.doc.common import *


def create_test_template():
    """创建包含模板变量的测试文档"""

    # 创建新文档
    doc = Document()

    # 添加节
    section = doc.AddSection()

    # 添加段落
    paragraph = section.AddParagraph()

    # 添加标题
    title_run = paragraph.AppendText("模板变量测试文档\n\n")
    title_format = title_run.CharacterFormat
    title_format.FontName = "黑体"
    title_format.FontSize = 18
    title_format.Bold = True

    # 添加各种变量示例
    paragraph.AppendText("姓名：${name}\n")
    paragraph.AppendText("年龄：${age}\n")
    paragraph.AppendText("部门：${dept|未知}\n")  # 带默认值
    paragraph.AppendText("日期：${date}\n")
    paragraph.AppendText("金额：${amount}\n\n")

    # 添加说明
    paragraph.AppendText("变量说明：\n")
    paragraph.AppendText("- ${name}: 姓名变量\n")
    paragraph.AppendText("- ${age}: 年龄变量\n")
    paragraph.AppendText("- ${dept|未知}: 部门变量（带默认值）\n")
    paragraph.AppendText("- ${date}: 日期变量\n")
    paragraph.AppendText("- ${amount}: 金额变量\n\n")

    # 添加表格
    table = section.AddTable(True)
    table.ResetCells(3, 2)  # 3行2列

    # 填充表头
    row = table.Rows.get_Item(0)
    cell = row.Cells.get_Item(0)
    para = cell.AddParagraph()
    para.AppendText("字段")
    cell = row.Cells.get_Item(1)
    para = cell.AddParagraph()
    para.AppendText("值")

    # 填充数据行
    row = table.Rows.get_Item(1)
    cell = row.Cells.get_Item(0)
    para = cell.AddParagraph()
    para.AppendText("姓名")
    cell = row.Cells.get_Item(1)
    para = cell.AddParagraph()
    para.AppendText("${name}")

    row = table.Rows.get_Item(2)
    cell = row.Cells.get_Item(0)
    para = cell.AddParagraph()
    para.AppendText("年龄")
    cell = row.Cells.get_Item(1)
    para = cell.AddParagraph()
    para.AppendText("${age}")

    # 保存文档
    output_path = "fixtures/templates/template_vars.docx"
    doc.SaveToFile(output_path, FileFormat.Docx)
    print(f"测试模板已创建：{output_path}")

    # 再创建一个简单的示例文档
    create_simple_template()


def create_simple_template():
    """创建简单的示例模板"""

    doc = Document()
    section = doc.AddSection()
    paragraph = section.AddParagraph()

    # 添加内容
    paragraph.AppendText("合同编号：${contract_no}\n")
    paragraph.AppendText("甲方：${party_a}\n")
    paragraph.AppendText("乙方：${party_b}\n")
    paragraph.AppendText("金额：${amount}\n")
    paragraph.AppendText("日期：${date}\n\n")

    paragraph.AppendText("签字：${signature|未签字}")

    # 保存文档
    output_path = "fixtures/templates/contract_template.docx"
    doc.SaveToFile(output_path, FileFormat.Docx)
    print(f"合同模板已创建：{output_path}")


if __name__ == "__main__":
    import os
    from pathlib import Path

    # 确保目录存在
    output_dir = Path("fixtures/templates")
    output_dir.mkdir(parents=True, exist_ok=True)

    # 创建模板
    create_test_template()
    print("\n所有测试模板创建完成！")
