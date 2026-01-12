"""
DocxLib 基础用法示例

演示如何加载文档、填充字段、保存文档。
"""

from docxlib import fill_date, fill_text, load_docx, save_docx


def main():
    """基础用法示例"""

    # 加载模板文档
    print("加载模板文档...")
    doc = load_docx("fixtures/templates/simple.docx")

    # 直接定位填充
    print("填充字段...")
    fill_text(doc, (1, 1, 2, 2), "测试文本", h_align="center", v_align="center")

    # 右侧填充
    fill_text(
        doc, "姓名", "张三", mode="match_right", h_align="center", v_align="center"
    )

    # 下方填充
    fill_text(
        doc, "项目", "智慧城市", mode="match_down", h_align="center", v_align="center"
    )

    # 日期填充
    fill_date(doc, (1, 1, 3, 2), "2024年1月15日", h_align="center", v_align="center")

    # 带样式填充
    fill_text(
        doc,
        "年龄",
        "32",
        mode="match_right",
        font_name="黑体",
        font_size=16,
        bold=True,
        color="red",
        h_align="center",
        v_align="center",
    )

    # 保存文档
    print("保存文档...")
    save_docx(doc, "output/basic_output.docx")

    print("完成！")


if __name__ == "__main__":
    main()
