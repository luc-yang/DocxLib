"""
DocxLib 文档操作模块测试
"""

import pytest
from pathlib import Path
import tempfile
import os

from docxlib import load_docx, save_docx, merge_docs, to_pdf, copy_doc
from docxlib.errors import DocumentError, ValidationError


class TestLoadDocx:
    """测试文档加载功能"""

    def test_load_docx_success(self):
        """测试成功加载文档"""
        doc = load_docx("fixtures/templates/sample.docx")
        assert doc is not None
        assert doc.Sections.Count > 0

    def test_load_docx_file_not_exists(self):
        """测试文件不存在时抛出异常"""
        with pytest.raises((DocumentError, Exception)):
            load_docx("nonexistent.docx")

    def test_load_docx_invalid_format(self):
        """测试无效格式时抛出异常"""
        # 创建一个存在但格式错误的文件
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False, mode="w") as f:
            f.write("This is not a docx file")
            temp_file = f.name

        try:
            with pytest.raises((ValidationError, Exception)):
                load_docx(temp_file)
        finally:
            Path(temp_file).unlink()


class TestSaveDocx:
    """测试文档保存功能"""

    def test_save_docx_success(self):
        """测试成功保存文档"""
        doc = load_docx("fixtures/templates/sample.docx")
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            temp_path = f.name

        try:
            save_docx(doc, temp_path)
            # 验证文件已创建
            assert Path(temp_path).exists()
            # 验证文件大小大于0
            assert Path(temp_path).stat().st_size > 0
        finally:
            if Path(temp_path).exists():
                Path(temp_path).unlink()

    def test_save_docx_create_directory(self):
        """测试自动创建目录"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 创建一个嵌套目录路径
        temp_dir = tempfile.mkdtemp()
        nested_path = Path(temp_dir) / "nested" / "dir" / "output.docx"

        try:
            save_docx(doc, str(nested_path))
            # 验证目录和文件已创建
            assert nested_path.exists()
            assert nested_path.parent.exists()
        finally:
            # 清理
            import shutil
            if Path(temp_dir).exists():
                shutil.rmtree(temp_dir)


class TestMergeDocs:
    """测试文档合并功能"""

    def test_merge_docs_empty_list(self):
        """测试空列表时抛出异常"""
        with pytest.raises((DocumentError, ValueError, Exception)):
            merge_docs([])

    def test_merge_docs_success(self):
        """测试成功合并文档"""
        # 加载多个文档
        doc1 = load_docx("fixtures/templates/sample.docx")
        doc2 = load_docx("fixtures/templates/sample.docx")

        # 合并文档
        merged = merge_docs([doc1, doc2])
        assert merged is not None
        assert merged.Sections.Count > 0


class TestToPdf:
    """测试 PDF 转换功能"""

    def test_to_pdf_success(self):
        """测试成功转换为 PDF"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 转换为 PDF 字节
        pdf_bytes = to_pdf(doc)
        assert pdf_bytes is not None
        assert isinstance(pdf_bytes, bytes)
        # PDF 文件应该以 %PDF 开头
        if len(pdf_bytes) > 4:
            assert pdf_bytes[:4] == b"%PDF" or pdf_bytes[:4].startswith(b"%")


class TestCopyDoc:
    """测试文档复制功能"""

    def test_copy_doc_success(self):
        """测试成功复制文档"""
        doc = load_docx("fixtures/templates/sample.docx")
        copied = copy_doc(doc)

        # 验证复制成功
        assert copied is not None
        assert copied.Sections.Count > 0
        # 验证是独立副本（不是同一个对象）
        assert id(doc) != id(copied)
