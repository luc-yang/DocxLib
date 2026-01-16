"""
DocxLib 命令行工具

提供简单的命令行接口用于快速测试和文档处理。
"""

import argparse
import sys


def cmd_version(args: argparse.Namespace) -> int:
    """显示版本信息"""
    from docxlib import __version__

    print(f"DocxLib version {__version__}")
    return 0


def cmd_info(args: argparse.Namespace) -> int:
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


def cmd_test(args: argparse.Namespace) -> int:
    """运行测试"""
    import pytest

    print(f"Running tests from: {args.project_dir}")
    # 使用 pytest API 运行测试
    exit_code = pytest.main([args.project_dir, "-v"])
    return exit_code


def cmd_validate(args: argparse.Namespace) -> int:
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


def cmd_inspect(args: argparse.Namespace) -> int:
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


def cmd_extract_vars(args: argparse.Namespace) -> int:
    """提取模板变量"""
    from pathlib import Path
    from docxlib import load_docx, extract_template_vars
    import json

    file_path = Path(args.file)
    if not file_path.exists():
        print(f"Error: File not found: {args.file}")
        return 1

    try:
        doc = load_docx(file_path)

        # 提取变量
        vars_list = extract_template_vars(doc, unique=True)

        print("=" * 50)
        print(f"Template: {file_path.name}")
        print(f"Found {len(vars_list)} variables:")
        print("=" * 50)
        for var in vars_list:
            print(f"  - {var}")
        print("=" * 50)

        # 可选：保存到 JSON 文件
        if args.output:
            output_path = Path(args.output)
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump({"variables": vars_list}, f, ensure_ascii=False, indent=2)
            print(f"Variables saved to: {output_path}")

        return 0

    except Exception as e:
        print(f"Error extracting variables: {e}")
        return 1


def cmd_fill(args: argparse.Namespace) -> int:
    """填充模板文档"""
    from pathlib import Path
    from docxlib import load_docx, fill_template, save_docx, parse_json

    template_path = Path(args.template)
    data_path = Path(args.data)
    output_path = Path(args.output) if args.output else None

    if not template_path.exists():
        print(f"Error: Template file not found: {args.template}")
        return 1

    if not data_path.exists():
        print(f"Error: Data file not found: {args.data}")
        return 1

    try:
        # 加载模板和数据
        doc = load_docx(template_path)

        # 解析数据
        if data_path.suffix.lower() == ".json":
            data = parse_json(data_path)
        else:
            print(f"Error: Unsupported data format: {data_path.suffix}")
            return 1

        # 填充模板
        result = fill_template(doc, data)

        print("=" * 50)
        print("Template Fill Result:")
        print(f"  Total variables: {result.get('total', 0)}")
        print(f"  Replaced: {result.get('replaced', 0)}")
        if result.get("missing"):
            print(f"  Missing: {', '.join(result['missing'])}")
        print("=" * 50)

        # 保存文档
        if not output_path:
            output_path = template_path.parent / f"{template_path.stem}_filled{template_path.suffix}"

        save_docx(doc, output_path)
        print(f"Document saved to: {output_path}")

        return 0

    except Exception as e:
        print(f"Error filling template: {e}")
        return 1


def cmd_convert(args: argparse.Namespace) -> int:
    """转换文档格式"""
    from pathlib import Path
    from docxlib import load_docx, to_pdf_file

    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else None

    if not input_path.exists():
        print(f"Error: Input file not found: {args.input}")
        return 1

    try:
        doc = load_docx(input_path)

        # 确定输出格式
        if args.format:
            fmt = args.format.lower()
        elif output_path:
            ext = output_path.suffix.lower()
            fmt = "pdf" if ext == ".pdf" else "unknown"
        else:
            print("Error: Must specify either --format or --output")
            return 1

        if fmt != "pdf":
            print(f"Error: Unsupported format: {fmt}")
            return 1

        # 确定输出路径
        if not output_path:
            output_path = input_path.parent / f"{input_path.stem}.{fmt}"

        # 转换
        print(f"Converting {input_path.name} to {fmt.upper()}...")
        to_pdf_file(doc, str(output_path))
        print(f"Document saved to: {output_path}")

        return 0

    except Exception as e:
        print(f"Error converting document: {e}")
        return 1


def main() -> int:
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

    # extract-vars 命令
    extract_vars_parser = subparsers.add_parser("extract-vars", help="Extract template variables")
    extract_vars_parser.add_argument("file", help="Template DOCX file")
    extract_vars_parser.add_argument("-o", "--output", help="Output JSON file path (optional)")

    # fill 命令
    fill_parser = subparsers.add_parser("fill", help="Fill template with data")
    fill_parser.add_argument("template", help="Template DOCX file")
    fill_parser.add_argument("data", help="Data file (JSON)")
    fill_parser.add_argument("-o", "--output", help="Output DOCX file path (optional)")

    # convert 命令
    convert_parser = subparsers.add_parser("convert", help="Convert document format")
    convert_parser.add_argument("input", help="Input DOCX file")
    convert_parser.add_argument("-f", "--format", choices=["pdf"], help="Output format")
    convert_parser.add_argument("-o", "--output", help="Output file path")

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
    elif args.command == "extract-vars":
        return cmd_extract_vars(args)
    elif args.command == "fill":
        return cmd_fill(args)
    elif args.command == "convert":
        return cmd_convert(args)
    else:
        parser.print_help()
        return 0


if __name__ == "__main__":
    sys.exit(main())
