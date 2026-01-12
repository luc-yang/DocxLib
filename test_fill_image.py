"""测试 fill_image 函数的改进功能"""

from docxlib import load_docx, fill_image, save_docx


def main():
    """测试图片填充功能"""

    print("测试 fill_image 函数...")

    # 加载模板文档
    doc = load_docx("fixtures/templates/sample.docx")

    # 测试1: 从文件路径填充图片
    print("\n测试1: 从文件路径填充图片")
    try:
        fill_image(doc, (1, 1, 2, 2), "fixtures/images/logo.png", width=80, height=80)
        print("✓ 从文件路径填充成功")
    except FileNotFoundError as e:
        print(f"✗ 图片文件不存在: {e}")
    except Exception as e:
        print(f"✗ 填充失败: {e}")

    # 测试2: 从字节数据填充图片
    print("\n测试2: 从字节数据填充图片")
    try:
        with open("fixtures/images/logo.png", "rb") as f:
            image_data = f.read()
        fill_image(doc, (1, 1, 3, 2), image_data, width=60, height=60)
        print("✓ 从字节数据填充成功")
    except FileNotFoundError as e:
        print(f"✗ 图片文件不存在: {e}")
    except Exception as e:
        print(f"✗ 填充失败: {e}")

    # 保存文档
    print("\n保存测试文档...")
    save_docx(doc, "output/test_fill_image.docx")
    print("✓ 文档已保存到 output/test_fill_image.docx")


if __name__ == "__main__":
    main()
