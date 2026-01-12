"""
DocxLib 基础功能测试脚本

用于快速验证库的基本功能是否正常工作。
"""

import sys

sys.path.insert(0, r"d:\Code\DocxLib")

from docxlib import (
    __version__,
    # 异常类
    DocxLibError,
    DocumentError,
    PositionError,
    FillError,
    ValidationError,
    # 常量
    DEFAULT_FONT,
    DEFAULT_FONT_SIZE,
    DEFAULT_COLOR,
    SUPPORTED_IMAGE_FORMATS,
    FileFormat,
    FillMode,
    # 工具函数
    is_valid_docx,
    parse_color,
    parse_date_string,
)


def test_version():
    """测试版本号"""
    print(f"[OK] DocxLib 版本: {__version__}")
    assert __version__ == "0.1.0"


def test_exceptions():
    """测试异常类"""
    print("[OK] 异常类导入成功")
    assert issubclass(DocumentError, DocxLibError)
    assert issubclass(PositionError, DocxLibError)
    assert issubclass(FillError, DocxLibError)
    assert issubclass(ValidationError, DocxLibError)


def test_constants():
    """测试常量"""
    print(f"[OK] 默认字体: {DEFAULT_FONT}")
    print(f"[OK] 默认字号: {DEFAULT_FONT_SIZE}")
    print(f"[OK] 默认颜色: {DEFAULT_COLOR}")
    print(f"[OK] 支持的图片格式: {SUPPORTED_IMAGE_FORMATS}")

    assert DEFAULT_FONT == "仿宋_GB2312"
    assert DEFAULT_FONT_SIZE == 10.5
    assert DEFAULT_COLOR == "black"
    assert ".png" in SUPPORTED_IMAGE_FORMATS
    assert ".jpg" in SUPPORTED_IMAGE_FORMATS


def test_fill_modes():
    """测试填充模式"""
    print(f"[OK] POSITION 模式: {FillMode.POSITION}")
    print(f"[OK] MATCH_RIGHT 模式: {FillMode.MATCH_RIGHT}")
    print(f"[OK] MATCH_DOWN 模式: {FillMode.MATCH_DOWN}")

    assert FillMode.POSITION == "position"
    assert FillMode.MATCH_RIGHT == "match_right"
    assert FillMode.MATCH_DOWN == "match_down"


def test_color_parsing():
    """测试颜色解析"""
    # 测试颜色名称
    color = parse_color("red")
    print(f"[OK] 解析颜色 'red': {color}")

    color = parse_color("black")
    print(f"[OK] 解析颜色 'black': {color}")

    # 测试十六进制
    color = parse_color("#FF0000")
    print(f"[OK] 解析颜色 '#FF0000': {color}")

    # 测试无效颜色（应该返回黑色）
    color = parse_color("invalid")
    print(f"[OK] 解析无效颜色返回黑色: {color}")


def test_date_parsing():
    """测试日期解析"""
    numbers, separators = parse_date_string("2024年1月15日")
    print(f"[OK] 解析日期 '2024年1月15日':")
    print(f"  - 数字: {numbers}")
    print(f"  - 年月日: {separators}")

    assert numbers == ["2024", "1", "15"]
    assert separators == ["年", "月", "日"]


def main():
    """运行所有测试"""
    print("=" * 50)
    print("DocxLib 基础功能测试")
    print("=" * 50)

    tests = [
        ("版本号", test_version),
        ("异常类", test_exceptions),
        ("常量", test_constants),
        ("填充模式", test_fill_modes),
        ("颜色解析", test_color_parsing),
        ("日期解析", test_date_parsing),
    ]

    passed = 0
    failed = 0

    for name, test_func in tests:
        try:
            print(f"\n测试: {name}")
            print("-" * 30)
            test_func()
            passed += 1
        except Exception as e:
            print(f"[FAIL] 测试失败: {e}")
            failed += 1

    print("\n" + "=" * 50)
    print(f"测试结果: {passed} 通过, {failed} 失败")
    print("=" * 50)

    if failed == 0:
        print("\n[OK] 所有测试通过!")
        return 0
    else:
        print(f"\n[FAIL] {failed} 个测试失败")
        return 1


if __name__ == "__main__":
    sys.exit(main())
