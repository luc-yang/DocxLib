"""
DocxLib 批量处理示例

演示如何批量生成多个文档。
"""

from docxlib import load_docx, fill_text, fill_date, save_docx


def main():
    """批量处理示例"""

    # 模板文件路径
    template_path = "fixtures/templates/sample.docx"

    # 准备数据
    data = [
        {
            "name": "张三",
            "age": "25",
            "date": "2024年1月15日",
            "project": "智慧城市建设项目",
        },
        {
            "name": "李四",
            "age": "30",
            "date": "2024年1月16日",
            "project": "数字化改造项目",
        },
        {
            "name": "王五",
            "age": "32",
            "date": "2024年1月17日",
            "project": "信息系统集成项目",
        },
    ]

    # 批量生成文档
    print(f"开始批量生成 {len(data)} 个文档...")

    for i, item in enumerate(data):
        print(f"正在处理第 {i+1} 个文档...")

        # 每次重新加载模板（避免复制文档对象的问题）
        doc = load_docx(template_path)

        # 填充数据
        fill_text(doc, "姓名", item["name"], mode="match_right")
        fill_text(doc, "年龄", item["age"], mode="match_right")
        fill_text(doc, "项目", item["project"], mode="match_right")

        # 填充日期
        fill_date(doc, "日期", item["date"])

        # 保存文档
        save_docx(doc, f"output/batch_output_{i+1}.docx")

    print("批量生成完成！")


if __name__ == "__main__":
    main()
