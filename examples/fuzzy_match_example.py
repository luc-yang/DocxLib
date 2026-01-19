"""
DocxLib 文本匹配示例

演示如何处理模板中的空格、换行等格式问题。
"""

from docxlib import (
    Alignment,
    Options,
    Style,
    find_text,
    fill_date,
    fill_text,
    load_docx,
    save_docx,
)


def example_text_normalization():
    """示例：文本规范化匹配（推荐方式）"""
    print("=" * 60)
    print("示例 1: 文本规范化匹配")
    print("=" * 60)

    doc = load_docx("fixtures/templates/sample.docx")

    # 默认情况下，find_text 会自动规范化文本
    # 这意味着 "姓名" 可以匹配：
    # - "姓名"
    # - "姓    名" (中间有多个空格)
    # - "姓\n名" (中间有换行)
    # - "姓\t名" (中间有制表符)

    print("\n查找 '姓名'（自动规范化）：")
    positions = find_text(doc, "姓名")
    print(f"找到位置: {positions}")

    print("\n查找 '性别'（自动规范化）：")
    positions = find_text(doc, "性别")
    print(f"找到位置: {positions}")


def example_exact_match():
    """示例：精确匹配"""
    print("\n" + "=" * 60)
    print("示例 2: 精确匹配")
    print("=" * 60)

    doc = load_docx("fixtures/templates/sample.docx")

    # 精确匹配：不规范化文本
    # 只有完全匹配才会找到

    print("\n精确查找 '姓名'（不规范化）：")
    positions = find_text(doc, "姓名", normalize=False)
    print(f"找到位置: {positions}")

    # 如果模板中是 "姓    名"，精确匹配会失败
    # 但规范化匹配可以成功
    print("\n规范化查找可以匹配 '姓    名' 等情况")
    positions = find_text(doc, "姓名", normalize=True)
    print(f"找到位置: {positions}")


def example_fill_with_normalization():
    """示例：使用规范化匹配填充数据"""
    print("\n" + "=" * 60)
    print("示例 3: 填充数据（自动处理空格/换行）")
    print("=" * 60)

    doc = load_docx("fixtures/templates/sample.docx")

    # fill_text 使用的 match_right/match_down 模式
    # 默认启用规范化匹配（normalize=True）
    # 无需担心模板中的空格或换行

    print("\n填充数据...")

    # 这些调用会自动匹配，即使模板中有空格或换行
    fill_text(
        doc,
        "姓名",  # 会匹配 "姓名"、"姓    名"、"姓\n名" 等
        "张三",
        options=Options.match_right(),  # 默认 normalize=True
        style=Style(font_family="宋体", font_size=12),
        alignment=Alignment(h_align="left", v_align="middle"),
    )

    fill_text(
        doc,
        "性别",  # 会匹配 "性别"、"性  别"、"性\n别" 等
        "男",
        options=Options.match_right(),
        alignment=Alignment(h_align="center", v_align="middle"),
    )

    fill_text(
        doc,
        "工作单位",
        "某公司",
        options=Options.match_right(),
        alignment=Alignment(h_align="center", v_align="middle"),
    )

    print("填充完成！")

    # 保存文档
    save_docx(doc, "output/fuzzy_match_output.docx")
    print("已保存到: output/fuzzy_match_output.docx")


def example_fill_with_exact_match():
    """示例：禁用规范化匹配（精确匹配）"""
    print("\n" + "=" * 60)
    print("示例 4: 禁用规范化匹配（精确匹配）")
    print("=" * 60)

    doc = load_docx("fixtures/templates/sample.docx")

    print("\n使用精确匹配填充...")

    # 禁用规范化匹配，只匹配完全相同的文本
    try:
        fill_text(
            doc,
            "姓名",  # 只会匹配 "姓名"，不会匹配 "姓    名"
            "李四",
            options=Options.match_right(normalize=False),
            alignment=Alignment(h_align="left", v_align="middle"),
        )
    except Exception as e:
        print(f"填充姓名时出错: {e}")

    try:
        fill_date(
            doc,
            "日期",
            "2024年1月15日",
            normalize=False,  # 精确匹配
        )
    except Exception as e:
        print(f"填充日期时出错: {e}")

    print("填充完成！")


def example_fill_date_with_normalization():
    """示例：填充日期（规范化匹配）"""
    print("\n" + "=" * 60)
    print("示例 5: 填充日期（规范化匹配）")
    print("=" * 60)

    doc = load_docx("fixtures/templates/sample.docx")

    print("\n填充日期...")

    # fill_date 支持字符串模式（查找文本并填充到右侧）
    # 默认启用规范化匹配
    fill_date(
        doc,
        "入党时间",  # 会匹配 "日期"、"日  期"、"日\n期" 等
        "2024年1月15日",
        normalize=True,  # 可以显式指定，默认就是 True
    )

    print("日期填充完成！")


def example_template_scenarios():
    """示例：常见模板场景"""
    print("\n" + "=" * 60)
    print("示例 6: 常见模板匹配场景")
    print("=" * 60)

    doc = load_docx("fixtures/templates/sample.docx")

    scenarios = [
        ("姓名", "标准格式：姓名"),
        ("姓  名", "中间有空格：姓  名"),
        ("姓\n名", "中间有换行：姓\\n名"),
        ("姓\t名", "中间有制表符：姓\\t名"),
        ("姓名：", "带标点：姓名："),
        ("姓　名", "中间有全角空格：姓　名"),
    ]

    print("\n规范化匹配可以处理以下情况：")
    for search_text, description in scenarios:
        positions = find_text(doc, "姓名", normalize=True)
        status = "✓" if positions else "✗"
        print(f"  {status} {description}")

    print("\n" + "=" * 60)
    print("用户操作指南：")
    print("=" * 60)
    print()
    print("1. 使用 fill_text/fill_image/fill_date 的匹配模式时：")
    print()
    print("   # 默认行为（推荐）- 自动规范化")
    print('   fill_text(doc, "姓名", "张三",')
    print('            options=Options.match_right())')
    print()
    print("   # 禁用规范化 - 精确匹配")
    print('   fill_text(doc, "姓名", "张三",')
    print('            options=Options.match_right(normalize=False))')
    print()
    print("2. 使用 find_text 查找文本时：")
    print()
    print("   # 默认行为（推荐）- 自动规范化")
    print('   positions = find_text(doc, "姓名")')
    print()
    print("   # 禁用规范化 - 精确匹配")
    print('   positions = find_text(doc, "姓名", normalize=False)')
    print()
    print("3. 使用 fill_date 填充日期时：")
    print()
    print("   # 默认行为（推荐）- 自动规范化")
    print('   fill_date(doc, "日期", "2024年1月15日")')
    print()
    print("   # 禁用规范化 - 精确匹配")
    print('   fill_date(doc, "日期", "2024年1月15日", normalize=False)')
    print()


if __name__ == "__main__":
    # 确保输出目录存在
    from pathlib import Path
    Path("output").mkdir(exist_ok=True)

    # 运行示例
    example_text_normalization()
    example_exact_match()
    example_fill_with_normalization()
    example_fill_with_exact_match()
    example_fill_date_with_normalization()
    example_template_scenarios()

    print("\n" + "=" * 60)
    print("所有示例运行完成！")
    print("=" * 60)
