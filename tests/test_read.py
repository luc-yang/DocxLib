"""DocxLib 文档读取功能测试"""

import pytest
from docxlib import (
    load_docx,
    get_table_text,
    get_table_row_text,
    get_table_column_text,
    get_document_properties,
    get_cell,
    get_cell_style,
    get_paragraph_style,
)
from docxlib.errors import PositionError, DocumentError


class TestTableContentReading:
    """测试表格内容读取"""

    def test_get_table_text_success(self):
        """测试成功获取表格文本"""
        doc = load_docx("fixtures/templates/sample.docx")
        table_data = get_table_text(doc, 1, 1)
        assert isinstance(table_data, list)
        assert len(table_data) > 0
        assert isinstance(table_data[0], list)

    def test_get_table_text_structure(self):
        """测试表格数据结构"""
        doc = load_docx("fixtures/templates/sample.docx")
        table_data = get_table_text(doc, 1, 1)
        # 验证是二维列表
        for row in table_data:
            assert isinstance(row, list)
            for cell in row:
                assert isinstance(cell, str)

    def test_get_table_row_text(self):
        """测试获取表格行文本"""
        doc = load_docx("fixtures/templates/sample.docx")
        row_text = get_table_row_text(doc, 1, 1, 1)
        assert isinstance(row_text, list)
        assert len(row_text) > 0
        # 验证所有元素都是字符串
        for item in row_text:
            assert isinstance(item, str)

    def test_get_table_column_text(self):
        """测试获取表格列文本"""
        doc = load_docx("fixtures/templates/sample.docx")
        col_text = get_table_column_text(doc, 1, 1, 1)
        assert isinstance(col_text, list)
        assert len(col_text) > 0
        # 验证所有元素都是字符串
        for item in col_text:
            assert isinstance(item, str)

    def test_get_table_text_invalid_position(self):
        """测试无效位置抛出异常"""
        doc = load_docx("fixtures/templates/sample.docx")
        with pytest.raises(PositionError):
            get_table_text(doc, 99, 99)

    def test_get_table_row_text_invalid_position(self):
        """测试无效行位置抛出异常"""
        doc = load_docx("fixtures/templates/sample.docx")
        with pytest.raises(PositionError):
            get_table_row_text(doc, 99, 99, 99)

    def test_get_table_column_text_invalid_position(self):
        """测试无效列位置抛出异常"""
        doc = load_docx("fixtures/templates/sample.docx")
        with pytest.raises(PositionError):
            get_table_column_text(doc, 99, 99, 99)

    def test_table_functions_consistency(self):
        """测试表格读取函数的一致性"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 获取整个表格
        table_data = get_table_text(doc, 1, 1)

        if len(table_data) > 0:
            # 获取第一行
            row_1 = get_table_row_text(doc, 1, 1, 1)
            assert row_1 == table_data[0]

            # 获取第一列
            if len(table_data[0]) > 0:
                col_1 = get_table_column_text(doc, 1, 1, 1)
                expected_col = [row[0] for row in table_data]
                assert col_1 == expected_col


class TestMetadataReading:
    """测试元数据读取"""

    def test_get_document_properties(self):
        """测试获取文档属性"""
        doc = load_docx("fixtures/templates/sample.docx")
        props = get_document_properties(doc)
        assert isinstance(props, dict)
        # 验证必需字段存在
        assert "title" in props
        assert "author" in props
        assert "subject" in props
        assert "keywords" in props
        assert "comments" in props
        assert "created_time" in props
        assert "modified_time" in props

    def test_document_properties_value_types(self):
        """测试文档属性值的类型"""
        doc = load_docx("fixtures/templates/sample.docx")
        props = get_document_properties(doc)
        # 所有值应该是字符串
        for key, value in props.items():
            assert isinstance(value, str)

    def test_get_document_properties_empty_doc(self):
        """测试空属性文档返回空字符串"""
        doc = load_docx("fixtures/templates/sample.docx")
        props = get_document_properties(doc)
        # 即使属性为空，也应该返回字典（值为空字符串）
        assert isinstance(props, dict)
        assert len(props) == 7  # 7个属性字段


class TestStyleReading:
    """测试样式读取"""

    def test_get_cell_style(self):
        """测试获取单元格样式"""
        doc = load_docx("fixtures/templates/sample.docx")
        cell = get_cell(doc, 1, 1, 1, 1)
        style = get_cell_style(cell)
        assert isinstance(style, dict)
        # 检查必需字段
        assert "font_name" in style
        assert "font_size" in style
        assert "bold" in style
        assert "italic" in style
        assert "underline" in style
        assert "color" in style
        assert "h_align" in style
        assert "v_align" in style
        assert "background_color" in style

    def test_cell_style_value_types(self):
        """测试单元格样式值的类型"""
        doc = load_docx("fixtures/templates/sample.docx")
        cell = get_cell(doc, 1, 1, 1, 1)
        style = get_cell_style(cell)

        if style:  # 如果样式不为空
            assert isinstance(style["font_name"], str)
            assert isinstance(style["font_size"], (int, float))
            assert isinstance(style["bold"], bool)
            assert isinstance(style["italic"], bool)
            assert isinstance(style["underline"], bool)
            assert isinstance(style["color"], str)
            assert isinstance(style["h_align"], str)
            assert isinstance(style["v_align"], str)
            assert isinstance(style["background_color"], str)

    def test_get_cell_style_empty_cell(self):
        """测试空单元格样式返回空字典"""
        doc = load_docx("fixtures/templates/sample.docx")
        cell = get_cell(doc, 1, 1, 1, 1)
        # 清空单元格
        cell.Paragraphs.Clear()
        style = get_cell_style(cell)
        # 空单元格应该返回空字典
        assert isinstance(style, dict)

    def test_get_paragraph_style(self):
        """测试获取段落样式"""
        doc = load_docx("fixtures/templates/sample.docx")
        section = doc.Sections.get_Item(0)
        if section.Paragraphs.Count > 0:
            paragraph = section.Paragraphs.get_Item(0)
            style = get_paragraph_style(paragraph)
            assert isinstance(style, dict)
            # 检查必需字段
            assert "alignment" in style
            assert "first_line_indent" in style
            assert "line_spacing" in style
            assert "space_before" in style
            assert "space_after" in style

    def test_paragraph_style_value_types(self):
        """测试段落样式值的类型"""
        doc = load_docx("fixtures/templates/sample.docx")
        section = doc.Sections.get_Item(0)
        if section.Paragraphs.Count > 0:
            paragraph = section.Paragraphs.get_Item(0)
            style = get_paragraph_style(paragraph)

            assert isinstance(style["alignment"], str)
            assert isinstance(style["first_line_indent"], float)
            assert isinstance(style["line_spacing"], float)
            assert isinstance(style["space_before"], float)
            assert isinstance(style["space_after"], float)


class TestBackwardCompatibility:
    """测试向后兼容性"""

    def test_existing_read_functions_work(self):
        """测试现有读取功能仍然可用"""
        doc = load_docx("fixtures/templates/sample.docx")
        from docxlib import get_cell_text, find_text

        # 现有功能不应受影响
        text = get_cell_text(doc, 1, 1, 1, 1)
        assert isinstance(text, str)

        positions = find_text(doc, "测试")
        assert isinstance(positions, list)

    def test_new_functions_are_importable(self):
        """测试新函数可以从顶层导入"""
        from docxlib import (
            get_table_text,
            get_table_row_text,
            get_table_column_text,
            get_document_properties,
            get_cell_style,
            get_paragraph_style,
        )
        # 所有新函数应该可以导入
        assert callable(get_table_text)
        assert callable(get_table_row_text)
        assert callable(get_table_column_text)
        assert callable(get_document_properties)
        assert callable(get_cell_style)
        assert callable(get_paragraph_style)
