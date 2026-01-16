"""
DocxLib CLI 模块测试
"""

import argparse
import json
from pathlib import Path
import tempfile
import pytest

from docxlib.cli import (
    cmd_version,
    cmd_info,
    cmd_test,
    cmd_validate,
    cmd_inspect,
    cmd_extract_vars,
    cmd_fill,
    cmd_convert,
    main,
)


class TestCmdVersion:
    """测试 version 命令"""

    def test_cmd_version_success(self):
        """测试成功显示版本信息"""
        args = argparse.Namespace()
        result = cmd_version(args)
        assert result == 0


class TestCmdInfo:
    """测试 info 命令"""

    def test_cmd_info_success(self):
        """测试成功显示库信息"""
        args = argparse.Namespace()
        result = cmd_info(args)
        assert result == 0


class TestCmdValidate:
    """测试 validate 命令"""

    def test_cmd_valid_docx(self):
        """测试验证有效的 DOCX 文件"""
        args = argparse.Namespace(file="fixtures/templates/sample.docx")
        result = cmd_validate(args)
        assert result == 0

    def test_cmd_invalid_file(self):
        """测试文件不存在"""
        args = argparse.Namespace(file="nonexistent.docx")
        result = cmd_validate(args)
        assert result == 1

    def test_cmd_invalid_format(self):
        """测试无效的文件格式"""
        # 创建一个临时非 DOCX 文件
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as f:
            f.write(b"Not a DOCX file")
            temp_path = f.name

        try:
            args = argparse.Namespace(file=temp_path)
            result = cmd_validate(args)
            assert result == 1
        finally:
            Path(temp_path).unlink()


class TestCmdInspect:
    """测试 inspect 命令"""

    def test_cmd_inspect_success(self):
        """测试成功检查文档信息"""
        args = argparse.Namespace(file="fixtures/templates/sample.docx")
        result = cmd_inspect(args)
        assert result == 0

    def test_cmd_inspect_file_not_found(self):
        """测试文件不存在"""
        args = argparse.Namespace(file="nonexistent.docx")
        result = cmd_inspect(args)
        assert result == 1


class TestCmdExtractVars:
    """测试 extract-vars 命令"""

    def test_cmd_extract_vars_success(self):
        """测试成功提取模板变量"""
        args = argparse.Namespace(file="fixtures/templates/sample.docx", output=None)
        result = cmd_extract_vars(args)
        # 命令应该成功执行，即使没有找到变量
        assert result == 0

    def test_cmd_extract_vars_with_output(self):
        """测试提取变量并保存到文件"""
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
            temp_path = f.name

        try:
            args = argparse.Namespace(file="fixtures/templates/sample.docx", output=temp_path)
            result = cmd_extract_vars(args)
            assert result == 0

            # 验证输出文件存在且是有效的 JSON
            output = Path(temp_path)
            assert output.exists()
            with open(output, "r", encoding="utf-8") as f:
                data = json.load(f)
                assert "variables" in data
                assert isinstance(data["variables"], list)
        finally:
            Path(temp_path).unlink()

    def test_cmd_extract_vars_file_not_found(self):
        """测试文件不存在"""
        args = argparse.Namespace(file="nonexistent.docx", output=None)
        result = cmd_extract_vars(args)
        assert result == 1


class TestCmdFill:
    """测试 fill 命令"""

    def test_cmd_fill_template_not_found(self):
        """测试模板文件不存在"""
        args = argparse.Namespace(template="nonexistent.docx", data="fixtures/data/sample.json", output=None)
        result = cmd_fill(args)
        assert result == 1

    def test_cmd_fill_data_not_found(self):
        """测试数据文件不存在"""
        args = argparse.Namespace(
            template="fixtures/templates/sample.docx",
            data="nonexistent.json",
            output=None,
        )
        result = cmd_fill(args)
        assert result == 1


class TestCmdConvert:
    """测试 convert 命令"""

    def test_cmd_convert_file_not_found(self):
        """测试输入文件不存在"""
        args = argparse.Namespace(input="nonexistent.docx", format=None, output=None)
        result = cmd_convert(args)
        assert result == 1

    def test_cmd_convert_missing_format_and_output(self):
        """测试缺少格式和输出参数"""
        args = argparse.Namespace(input="fixtures/templates/sample.docx", format=None, output=None)
        result = cmd_convert(args)
        assert result == 1

    def test_cmd_convert_unsupported_format(self):
        """测试不支持的格式"""
        # 注意：由于 choices=["pdf"]，argparse 会自动拒绝无效值
        # 这里测试格式验证逻辑
        args = argparse.Namespace(
            input="fixtures/templates/sample.docx",
            format=None,
            output=Path("output.docx")  # 不是 .pdf
        )
        result = cmd_convert(args)
        assert result == 1


class TestMain:
    """测试 main 函数"""

    def test_main_no_args(self, capsys):
        """测试无参数时显示帮助"""
        # 模拟无参数调用
        import sys
        old_argv = sys.argv
        try:
            sys.argv = ["docxlib"]
            result = main()
            # 应该显示帮助并返回 0
            assert result == 0
            captured = capsys.readouterr()
            assert "usage:" in captured.out or "show this help message" in captured.out.lower()
        finally:
            sys.argv = old_argv

    def test_main_version_flag(self, capsys):
        """测试 --version 参数"""
        import sys
        old_argv = sys.argv
        try:
            sys.argv = ["docxlib", "--version"]
            result = main()
            assert result == 0
            captured = capsys.readouterr()
            assert "DocxLib version" in captured.out
        finally:
            sys.argv = old_argv

    def test_main_version_command(self, capsys):
        """测试 version 子命令"""
        import sys
        old_argv = sys.argv
        try:
            sys.argv = ["docxlib", "version"]
            result = main()
            assert result == 0
            captured = capsys.readouterr()
            assert "DocxLib version" in captured.out
        finally:
            sys.argv = old_argv

    def test_main_info_command(self, capsys):
        """测试 info 子命令"""
        import sys
        old_argv = sys.argv
        try:
            sys.argv = ["docxlib", "info"]
            result = main()
            assert result == 0
            captured = capsys.readouterr()
            assert "DocxLib Information" in captured.out
        finally:
            sys.argv = old_argv
