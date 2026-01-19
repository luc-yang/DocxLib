"""示例：从 Word 文档中提取图片

演示如何使用 read_images() 函数提取文档中的图片元数据和字节数据。
"""

from docxlib import load_docx, read_images
import os


def read_images_metadata_only():
    """仅提取图片元数据（不包含图片数据）"""
    print("=" * 60)
    print("示例 1: 提取图片元数据")
    print("=" * 60)

    doc = load_docx("fixtures/templates/sample.docx")
    images = read_images(doc)

    print(f"找到 {len(images)} 张图片\n")

    for i, img in enumerate(images):
        print(f"图片 {i + 1}:")
        print(f"  位置: 节{img['position'][0]}, 表格{img['position'][1]}, "
              f"行{img['position'][2]}, 列{img['position'][3]}")
        print(f"  尺寸: {img['width']:.2f} x {img['height']:.2f} 磅")
        print(f"  格式: {img['format']}")
        print(f"  索引: {img['index']}")
        print()


def read_images_with_data():
    """提取图片并包含字节数据"""
    print("=" * 60)
    print("示例 2: 提取图片并保存到文件")
    print("=" * 60)

    doc = load_docx("fixtures/templates/sample.docx")
    images = read_images(doc, include_data=True)

    print(f"找到 {len(images)} 张图片\n")

    # 创建输出目录
    output_dir = "output/extracted_images"
    os.makedirs(output_dir, exist_ok=True)

    for i, img in enumerate(images):
        # 生成文件名
        position = img["position"]
        filename = f"image_s{position[0]}_t{position[1]}_r{position[2]}_c{position[3]}.{img['format']}"
        filepath = os.path.join(output_dir, filename)

        # 保存图片
        if img["data"]:
            with open(filepath, "wb") as f:
                f.write(img["data"])

            file_size = len(img["data"])
            print(f"保存图片 {i + 1}: {filename} ({file_size} 字节)")
        else:
            print(f"图片 {i + 1}: 无数据（可能是占位符或损坏）")

    print(f"\n所有图片已保存到: {output_dir}")


def read_images_from_specific_table():
    """仅提取特定表格中的图片"""
    print("=" * 60)
    print("示例 3: 提取特定表格的图片")
    print("=" * 60)

    doc = load_docx("fixtures/templates/sample.docx")

    # 仅提取第1节第1个表格的图片
    images = read_images(doc, section=1, table=1)

    print(f"第1节第1个表格中找到 {len(images)} 张图片\n")

    for i, img in enumerate(images):
        print(f"图片 {i + 1}:")
        print(f"  位置: 行{img['position'][2]}, 列{img['position'][3]}")
        print(f"  尺寸: {img['width']:.2f} x {img['height']:.2f} 磅")
        print(f"  格式: {img['format']}")


if __name__ == "__main__":
    # 运行示例
    read_images_metadata_only()
    print("\n")

    read_images_with_data()
    print("\n")

    read_images_from_specific_table()
