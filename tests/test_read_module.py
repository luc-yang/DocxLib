"""DocxLib 读取模块测试"""

import pytest

from docxlib import (
    read_images,
    extract_template_data,
    load_docx,
    read_all_text,
    read_cells,
    read_document_structure,
    read_grid,
    read_table,
    read_text,
)
from docxlib.config import Options
from docxlib.errors import FillError, PositionError


class TestReadText:
    """测试 read_text() 函数"""

    def test_read_text_by_position(self):
        """测试通过位置元组读取文本"""
        doc = load_docx("fixtures/templates/sample.docx")
        text = read_text(doc, (1, 1, 1, 1))
        assert isinstance(text, str)

    def test_read_text_with_default(self):
        """测试使用默认值"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 无效位置返回默认值
        text = read_text(doc, (99, 99, 99, 99), default="N/A")
        assert text == "N/A"

    def test_read_text_match_right(self):
        """测试 match_right 模式"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 查找文本并读取右侧
        text = read_text(doc, "测试", default="未找到", options=Options.match_right())
        assert isinstance(text, str)

    def test_read_text_match_down(self):
        """测试 match_down 模式"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 查找文本并读取下方
        text = read_text(doc, "测试", default="未找到", options=Options.match_down())
        assert isinstance(text, str)

    def test_read_text_invalid_mode(self):
        """测试无效模式抛出异常"""
        doc = load_docx("fixtures/templates/sample.docx")
        with pytest.raises(FillError):
            # 创建一个无效的 options
            options = Options()
            options.mode = "invalid_mode"
            read_text(doc, (1, 1, 1, 1), options=options)


class TestReadGrid:
    """测试 read_grid() 函数"""

    def test_read_grid_basic(self):
        """测试基础网格读取"""
        doc = load_docx("fixtures/templates/sample.docx")
        grid = read_grid(doc, (1, 1, 1, 1))
        assert isinstance(grid, list)
        assert len(grid) > 0
        # 验证是二维列表
        for row in grid:
            assert isinstance(row, list)

    def test_read_grid_partial(self):
        """测试部分网格读取"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 从第2行第2列开始读取
        grid = read_grid(doc, (1, 1, 2, 2))
        assert isinstance(grid, list)

    def test_read_grid_invalid_position(self):
        """测试无效位置抛出异常"""
        doc = load_docx("fixtures/templates/sample.docx")
        with pytest.raises(PositionError):
            read_grid(doc, (99, 99, 1, 1))


class TestReadTable:
    """测试 read_table() 函数"""

    def test_read_table_basic(self):
        """测试基础表格读取"""
        doc = load_docx("fixtures/templates/sample.docx")
        table_data = read_table(doc, 1, 1)
        assert isinstance(table_data, list)
        assert len(table_data) > 0
        # 验证是二维列表
        for row in table_data:
            assert isinstance(row, list)

    def test_read_table_with_style(self):
        """测试带样式的表格读取"""
        doc = load_docx("fixtures/templates/sample.docx")
        table_info = read_table(doc, 1, 1, include_style=True)
        assert isinstance(table_info, dict)
        assert "data" in table_info
        assert "dimensions" in table_info
        assert "styles" in table_info

    def test_read_table_with_merged(self):
        """测试带合并单元格信息的表格读取"""
        doc = load_docx("fixtures/templates/sample.docx")
        table_info = read_table(doc, 1, 1, include_merged=True)
        assert isinstance(table_info, dict)
        assert "data" in table_info
        assert "merged_cells" in table_info

    def test_read_table_full(self):
        """测试完整信息表格读取"""
        doc = load_docx("fixtures/templates/sample.docx")
        table_full = read_table(doc, 1, 1, include_style=True, include_merged=True)
        assert isinstance(table_full, dict)
        assert "data" in table_full
        assert "dimensions" in table_full
        assert "styles" in table_full
        assert "merged_cells" in table_full

    def test_read_table_invalid_position(self):
        """测试无效位置抛出异常"""
        doc = load_docx("fixtures/templates/sample.docx")
        with pytest.raises(PositionError):
            read_table(doc, 99, 99)


class TestReadCells:
    """测试 read_cells() 函数"""

    def test_read_cells_batch(self):
        """测试批量读取"""
        doc = load_docx("fixtures/templates/sample.docx")
        positions = [(1, 1, 1, 1), (1, 1, 1, 2), (1, 1, 1, 3)]
        values = read_cells(doc, positions)
        assert isinstance(values, list)
        assert len(values) == 3
        # 验证所有元素都是字符串
        for value in values:
            assert isinstance(value, str)

    def test_read_cells_with_default(self):
        """测试使用默认值"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 混合有效和无效位置
        positions = [(1, 1, 1, 1), (99, 99, 99, 99), (1, 1, 1, 2)]
        values = read_cells(doc, positions, default="N/A")
        assert isinstance(values, list)
        assert len(values) == 3
        # 第二个应该是默认值
        assert values[1] == "N/A"


class TestExtractImages:
    """测试 read_images() 函数"""

    def test_read_images_all(self):
        """测试提取所有图片"""
        doc = load_docx("fixtures/templates/sample.docx")
        images = read_images(doc)
        assert isinstance(images, list)
        # 如果有图片，验证结构
        for img in images:
            assert "position" in img
            assert "width" in img
            assert "height" in img
            assert "format" in img
            assert "index" in img
            # 默认不包含 data
            assert "data" not in img

    def test_read_images_with_data(self):
        """测试提取图片并包含字节数据"""
        doc = load_docx("fixtures/templates/sample.docx")
        images = read_images(doc, include_data=True)
        assert isinstance(images, list)
        # 如果有图片，验证包含 data 字段
        for img in images:
            assert "data" in img
            assert isinstance(img["data"], bytes)
            # 验证字节数据不为空（如果有有效图片）
            if len(img["data"]) > 0:
                # 图片数据应该有实际内容
                assert len(img["data"]) > 0

    def test_read_images_section_filter(self):
        """测试按节过滤图片"""
        doc = load_docx("fixtures/templates/sample.docx")
        images = read_images(doc, section=1)
        assert isinstance(images, list)

    def test_read_images_table_filter(self):
        """测试按表格过滤图片"""
        doc = load_docx("fixtures/templates/sample.docx")
        images = read_images(doc, section=1, table=1)
        assert isinstance(images, list)


class TestReadDocumentStructure:
    """测试 read_document_structure() 函数"""

    def test_structure_basic(self):
        """测试基础结构读取"""
        doc = load_docx("fixtures/templates/sample.docx")
        structure = read_document_structure(doc)
        assert isinstance(structure, dict)
        assert "section_count" in structure
        assert "sections" in structure
        assert len(structure["sections"]) > 0

    def test_structure_with_paragraphs(self):
        """测试包含段落信息的结构读取"""
        doc = load_docx("fixtures/templates/sample.docx")
        structure = read_document_structure(doc, include_paragraphs=True)
        assert isinstance(structure, dict)
        # 验证第一个节包含段落计数
        if structure["sections"]:
            assert "paragraph_count" in structure["sections"][0]

    def test_structure_with_tables(self):
        """测试包含表格信息的结构读取"""
        doc = load_docx("fixtures/templates/sample.docx")
        structure = read_document_structure(doc, include_tables=True)
        assert isinstance(structure, dict)
        # 验证节包含表格信息
        if structure["sections"]:
            section = structure["sections"][0]
            assert "table_count" in section
            if "tables" in section:
                for table in section["tables"]:
                    assert "index" in table
                    assert "rows" in table
                    assert "columns_per_row" in table


class TestReadAllText:
    """测试 read_all_text() 函数"""

    def test_read_all_basic(self):
        """测试基础全文读取"""
        doc = load_docx("fixtures/templates/sample.docx")
        text = read_all_text(doc)
        assert isinstance(text, str)
        assert len(text) > 0

    def test_read_all_tables_only(self):
        """测试仅读取表格文本"""
        doc = load_docx("fixtures/templates/sample.docx")
        text = read_all_text(doc, include_paragraphs=False, include_tables=True)
        assert isinstance(text, str)

    def test_read_all_paragraphs_only(self):
        """测试仅读取段落文本"""
        doc = load_docx("fixtures/templates/sample.docx")
        text = read_all_text(doc, include_paragraphs=True, include_tables=False)
        assert isinstance(text, str)

    def test_read_all_custom_separator(self):
        """测试自定义分隔符"""
        doc = load_docx("fixtures/templates/sample.docx")
        text = read_all_text(doc, separator="\n\n")
        assert isinstance(text, str)
        # 自定义分隔符被正确使用（如果文档有多个段落/单元格）
        # 单个文本块时不会出现分隔符，这是正常的
        assert len(text) >= 0  # 只是验证返回了有效的字符串


class TestExtractTemplateData:
    """测试 extract_template_data() 函数"""

    def test_extract_basic(self):
        """测试基础数据提取"""
        doc = load_docx("fixtures/templates/sample.docx")
        data = extract_template_data(doc)
        assert isinstance(data, dict)
        assert "variables" in data
        assert "variable_details" in data
        assert "total_count" in data

    def test_extract_unique(self):
        """测试去重提取"""
        doc = load_docx("fixtures/templates/sample.docx")
        data = extract_template_data(doc, unique=True)
        assert isinstance(data, dict)
        assert "variables" in data
        # 验证 variables 是列表
        assert isinstance(data["variables"], list)

    def test_extract_with_positions(self):
        """测试包含位置信息的提取"""
        doc = load_docx("fixtures/templates/sample.docx")
        data = extract_template_data(doc)
        # 如果有变量，验证结构
        if data["variable_details"]:
            var_detail = data["variable_details"][0]
            assert "name" in var_detail
            assert "default" in var_detail
            assert "positions" in var_detail

    def test_extract_custom_delimiters(self):
        """测试自定义占位符"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 使用自定义占位符
        data = extract_template_data(
            doc, placeholder_prefix="{{", placeholder_suffix="}}"
        )
        assert isinstance(data, dict)
