"""
DocxLib 命令行工具

提供简单的命令行接口用于快速测试和文档处理。
"""

import sys
import argparse


def cmd_version(args):
    """显示版本信息"""
    from docxlib import __version__

    print(f"DocxLib version {__version__}")
    return 0


def cmd_info(args):
    """显示库信息"""
    from docxlib import (
        __version__,
        DEFAULT_FONT,
        DEFAULT_FONT_SIZE,
        DEFAULT_COLOR,
        SUPPORTED_IMAGE_FORMATS,
        FillMode,
    )

    print("=" * 50)
    print("DocxLib Information")
    print("=" * 50)
    print(f"Version: {__version__}")
    print(f"Default Font: {DEFAULT_FONT}")
    print(f"Default Font Size: {DEFAULT_FONT_SIZE}")
    print(f"Default Color: {DEFAULT_COLOR}")
    print(f"Supported Image Formats: {', '.join(SUPPORTED_IMAGE_FORMATS)}")
    print(
        f"Fill Modes: {FillMode.POSITION}, {FillMode.MATCH_RIGHT}, {FillMode.MATCH_DOWN}"
    )
    print("=" * 50)
    return 0


def cmd_test(args):
    """运行测试"""
    import subprocess

    print("Running basic tests...")
    result = subprocess.run(
        [sys.executable, "tests/test_basic.py"], cwd=args.project_dir
    )
    return result.returncode


def cmd_validate(args):
    """验证 DOCX 文件"""
    from pathlib import Path
    from docxlib import is_valid_docx

    file_path = Path(args.file)
    if not file_path.exists():
        print(f"Error: File not found: {args.file}")
        return 1

    if is_valid_docx(file_path):
        print(f"[OK] {args.file} is a valid DOCX file")
        return 0
    else:
        print(f"[FAIL] {args.file} is NOT a valid DOCX file")
        return 1


def cmd_inspect(args):
    """检查文档信息"""
    from pathlib import Path
    from docxlib import load_docx, get_section_count, get_table_dimensions

    file_path = Path(args.file)
    if not file_path.exists():
        print(f"Error: File not found: {args.file}")
        return 1

    try:
        doc = load_docx(file_path)

        print("=" * 50)
        print(f"Document: {file_path.name}")
        print("=" * 50)

        # 节数量
        section_count = get_section_count(doc)
        print(f"Sections: {section_count}")

        # 遍历所有节和表格
        for sec_idx in range(1, section_count + 1):
            print(f"\nSection {sec_idx}:")
            print("-" * 30)

            try:
                from docxlib import get_section_table_count

                table_count = get_section_table_count(doc, sec_idx)
                print(f"  Tables: {table_count}")

                for tbl_idx in range(1, table_count + 1):
                    try:
                        rows, cols = get_table_dimensions(doc, sec_idx, tbl_idx)
                        print(f"    Table {tbl_idx}: {rows} rows x {cols} cols")
                    except Exception as e:
                        print(f"    Table {tbl_idx}: Error - {e}")
            except Exception as e:
                print(f"  Error: {e}")

        print("=" * 50)
        return 0

    except Exception as e:
        print(f"Error loading document: {e}")
        return 1


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        prog="docxlib", description="DocxLib - Word document processing library"
    )

    parser.add_argument(
        "--version", action="store_true", help="Show version information"
    )

    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # version 命令
    subparsers.add_parser("version", help="Show version information")

    # info 命令
    subparsers.add_parser("info", help="Show library information")

    # test 命令
    test_parser = subparsers.add_parser("test", help="Run tests")
    test_parser.add_argument(
        "--project-dir",
        default=".",
        help="Project directory (default: current directory)",
    )

    # validate 命令
    validate_parser = subparsers.add_parser("validate", help="Validate DOCX file")
    validate_parser.add_argument("file", help="DOCX file to validate")

    # inspect 命令
    inspect_parser = subparsers.add_parser("inspect", help="Inspect document structure")
    inspect_parser.add_argument("file", help="DOCX file to inspect")

    args = parser.parse_args()

    # 处理 --version 参数
    if args.version:
        return cmd_version(args)

    # 处理子命令
    if args.command == "version":
        return cmd_version(args)
    elif args.command == "info":
        return cmd_info(args)
    elif args.command == "test":
        return cmd_test(args)
    elif args.command == "validate":
        return cmd_validate(args)
    elif args.command == "inspect":
        return cmd_inspect(args)
    else:
        parser.print_help()
        return 0


if __name__ == "__main__":
    sys.exit(main())
