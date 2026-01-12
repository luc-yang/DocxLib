"""
DocxLib 基础用法示例

演示如何加载文档、填充字段、保存文档。
"""

from docxlib import load_docx, fill_text, save_docx


def main():
    """基础用法示例"""

    # 加载模板文档
    print("加载模板文档...")
    doc = load_docx("fixtures/templates/simple.docx")

    # 直接定位填充
    print("填充字段...")
    fill_text(doc, (1, 1, 2, 2), "测试文本")

    # 右侧填充
    fill_text(doc, "姓名：", "张三", mode="match_right")

    # 下方填充
    fill_text(doc, "项目1", "智慧城市", mode="match_down")

    # 带样式填充
    fill_text(doc, "标题", "内容",
              font_name="黑体",
              font_size=16,
              bold=True,
              color="red")

    # 保存文档
    print("保存文档...")
    save_docx(doc, "output/basic_output.docx")

    print("完成！")


if __name__ == "__main__":
    main()
