"""
DocxLib 基础用法示例

演示如何加载文档、填充字段、保存文档。
"""

from docxlib import (
    Alignment,
    Options,
    Style,
    fill_date,
    fill_text,
    load_docx,
    save_docx,
)


def main():
    """基础用法示例"""

    # 加载模板文档
    print("加载模板文档...")
    doc = load_docx("fixtures/templates/sample.docx")

    # 直接定位填充
    print("填充字段...")
    fill_text(
        doc,
        (1, 1, 2, 2),
        "测试文本",
        alignment=Alignment(h_align="center", v_align="middle"),
    )

    # 右侧填充
    fill_text(
        doc,
        "姓名",
        "张三",
        options=Options(mode="match_right"),
        alignment=Alignment(h_align="center", v_align="middle"),
    )

    # 下方填充
    fill_text(
        doc,
        "证明人",
        "李四",
        options=Options(mode="match_down"),
        alignment=Alignment(h_align="center", v_align="middle"),
    )

    # 日期填充
    fill_date(
        doc,
        (1, 1, 4, 2),
        "2024年1月15日",
        alignment=Alignment(h_align="center", v_align="middle"),
    )

    # 带样式填充
    fill_text(
        doc,
        "工作单位",
        "某公司",
        options=Options(mode="match_right"),
        style=Style(font_family="黑体", font_size=16, bold=True, color="red"),
        alignment=Alignment(h_align="center", v_align="middle"),
    )

    # 保存文档
    print("保存文档...")
    save_docx(doc, "output/basic_output.docx")

    print("完成！")


if __name__ == "__main__":
    main()
