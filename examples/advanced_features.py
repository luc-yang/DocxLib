"""
DocxLib 高级功能示例

演示图片填充、网格数据填充、文档转换等高级功能。
"""

from docxlib import (
    load_docx,
    fill_image,
    fill_grid,
    fill_date,
    to_pdf,
    find_text,
    iterate_cells,
    get_table_dimensions,
    save_docx,
)


def main():
    """高级功能示例"""

    print("加载模板...")
    doc = load_docx("fixtures/templates/simple.docx")

    # 获取表格信息
    print("\n=== 表格信息 ===")
    rows, cols = get_table_dimensions(doc, 1, 1)
    print(f"第一个表格大小: {rows} 行 x {cols} 列")

    # 查找文本
    print("\n=== 查找文本 ===")
    positions = find_text(doc, "姓名")
    print(f"找到 '姓名' 的位置: {positions}")

    # 遍历单元格
    print("\n=== 遍历单元格 ===")
    count = 0
    for sec, tbl, row, col, cell in iterate_cells(doc):
        text = cell.Range.Text.strip()
        if text:
            count += 1
            if count <= 5:  # 只显示前5个
                print(f"({sec}, {tbl}, {row}, {col}): {text}")

    # 填充图片
    print("\n=== 填充图片 ===")
    try:
        fill_image(doc, "照片：", "fixtures/images/logo.png",
                   mode="match_right",
                   width=80,
                   height=80)
        print("图片填充成功")
    except Exception as e:
        print(f"图片填充失败: {e}")

    # 填充日期
    print("\n=== 填充日期 ===")
    fill_date(doc, (1, 1, 4, 2), "2024年1月15日")
    print("日期填充成功")

    # 填充网格数据
    print("\n=== 填充网格数据 ===")
    data = [
        ["序号", "项目", "金额"],
        ["1", "设备费", "50000"],
        ["2", "人工费", "30000"],
        ["3", "材料费", "20000"],
    ]
    try:
        fill_grid(doc, data, position=(1, 1, 7, 1))
        print("网格数据填充成功")
    except Exception as e:
        print(f"网格数据填充失败: {e}")

    # 保存文档
    print("\n=== 保存文档 ===")
    save_docx(doc, "output/advanced_output.docx")
    print("文档保存成功")

    # 转换为 PDF
    print("\n=== 转换为 PDF ===")
    try:
        pdf_bytes = to_pdf(doc)
        with open("output/advanced_output.pdf", "wb") as f:
            f.write(pdf_bytes)
        print("PDF 转换成功")
    except Exception as e:
        print(f"PDF 转换失败: {e}")

    print("\n完成！")


if __name__ == "__main__":
    main()
