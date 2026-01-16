"""
DocxLib 模板变量功能使用示例

演示如何使用模板变量功能批量生成文档。
"""

from docxlib import load_docx, fill_template, save_docx, extract_template_vars, validate_template_data


def example_basic_usage():
    """基础用法示例"""
    print("=" * 50)
    print("示例1: 基础用法")
    print("=" * 50)

    # 1. 加载模板
    doc = load_docx("fixtures/templates/template_vars.docx")

    # 2. 准备数据
    data = {
        "name": "张三",
        "age": "28",
        "dept": "研发部",
        "date": "2024年1月15日",
        "amount": "50000"
    }

    # 3. 填充模板
    stats = fill_template(doc, data)

    # 4. 查看统计信息
    print(f"总变量数: {stats['total']}")
    print(f"成功替换: {stats['replaced']}")

    # 5. 保存文档
    save_docx(doc, "output/example_basic.docx")
    print("已保存到: output/example_basic.docx\n")


def example_extract_variables():
    """提取模板变量示例"""
    print("=" * 50)
    print("示例2: 提取模板变量")
    print("=" * 50)

    # 加载模板
    doc = load_docx("fixtures/templates/template_vars.docx")

    # 提取所有变量
    vars = extract_template_vars(doc)
    print(f"模板需要以下变量: {vars}")
    print(f"共 {len(vars)} 个变量\n")


def example_validate_data():
    """验证数据完整性示例"""
    print("=" * 50)
    print("示例3: 验证数据完整性")
    print("=" * 50)

    # 加载模板
    doc = load_docx("fixtures/templates/contract_template.docx")

    # 准备不完整的数据
    data = {
        "contract_no": "HT-2024-001",
        "party_a": "某某科技公司"
        # 缺少 party_b, amount, date, signature
    }

    # 验证数据
    result = validate_template_data(doc, data)

    print(f"数据是否完整: {result['is_valid']}")
    if not result['is_valid']:
        print(f"缺失的变量: {result['missing_vars']}")
        print(f"需要的变量: {result['required_vars']}")

    # 使用 ignore 模式填充（不会报错）
    stats = fill_template(doc, data, missing_var_action="ignore")
    save_docx(doc, "output/example_partial.docx")
    print("已保存部分填充的文档\n")


def example_default_values():
    """使用默认值示例"""
    print("=" * 50)
    print("示例4: 使用默认值")
    print("=" * 50)

    # 加载模板
    doc = load_docx("fixtures/templates/template_vars.docx")

    # 不提供 dept 变量，模板中的 ${dept|未知} 会使用默认值
    data = {
        "name": "李四",
        "age": "30",
        # dept 未提供，会使用默认值 "未知"
        "date": "2024年1月15日",
        "amount": "80000"
    }

    stats = fill_template(doc, data)
    save_docx(doc, "output/example_default.docx")
    print("已保存使用默认值的文档\n")


def example_batch_generation():
    """批量生成文档示例"""
    print("=" * 50)
    print("示例5: 批量生成文档")
    print("=" * 50)

    # 准备多条数据
    records = [
        {"name": "张三", "age": "25", "dept": "技术部", "date": "2024-01-10", "amount": "50000"},
        {"name": "李四", "age": "30", "dept": "市场部", "date": "2024-01-11", "amount": "60000"},
        {"name": "王五", "age": "28", "dept": "财务部", "date": "2024-01-12", "amount": "55000"},
    ]

    # 批量生成（每次重新加载模板）
    template_path = "fixtures/templates/template_vars.docx"
    for i, data in enumerate(records, 1):
        # 重新加载模板（避免修改原始模板）
        doc = load_docx(template_path)

        # 填充数据
        stats = fill_template(doc, data)

        # 保存文档
        output_file = f"output/batch_{i}_{data['name']}.docx"
        save_docx(doc, output_file)
        print(f"已生成: {output_file} (替换 {stats['replaced']} 个变量)")

    print()


def example_contract_generation():
    """合同生成示例"""
    print("=" * 50)
    print("示例6: 合同生成")
    print("=" * 50)

    # 加载合同模板
    doc = load_docx("fixtures/templates/contract_template.docx")

    # 填充合同信息
    contract_data = {
        "contract_no": "HT-2024-001",
        "party_a": "甲方科技有限公司",
        "party_b": "乙方服务公司",
        "amount": "1000000",
        "date": "2024年1月15日",
        "signature": "张三"
    }

    stats = fill_template(doc, contract_data)
    save_docx(doc, "output/contract_HT-2024-001.docx")
    print(f"合同已生成: contract_HT-2024-001.docx")
    print(f"替换了 {stats['replaced']} 个变量\n")


if __name__ == "__main__":
    # 确保输出目录存在
    from pathlib import Path
    Path("output").mkdir(exist_ok=True)

    # 运行所有示例
    try:
        example_basic_usage()
        example_extract_variables()
        example_validate_data()
        example_default_values()
        example_batch_generation()
        example_contract_generation()

        print("=" * 50)
        print("所有示例运行完成！")
        print("请查看 output/ 目录中的生成文档")
        print("=" * 50)

    except Exception as e:
        print(f"运行示例时出错: {e}")
        import traceback
        traceback.print_exc()
