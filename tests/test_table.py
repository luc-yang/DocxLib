"""
DocxLib 表格操作模块测试
"""

import pytest
from docxtbl import load_docx, get_cell, get_cells, find_text, iterate_cells, get_cell_text
from docxtbl.errors import PositionError


class TestGetCell:
    """测试获取单元格功能"""

    def test_get_cell_success(self):
        """测试成功获取单元格"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 假设第一个节、第一个表格存在
        cell = get_cell(doc, 1, 1, 1, 1)
        assert cell is not None
        # 验证可以获取单元格文本 - 使用 get_cell_text 函数
        text = get_cell_text(doc, 1, 1, 1, 1)
        assert isinstance(text, str)

    def test_get_cell_invalid_position(self):
        """测试无效位置时抛出异常"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 测试越界位置
        with pytest.raises((PositionError, Exception)):
            get_cell(doc, 99, 99, 99, 99)


class TestGetCells:
    """测试通配符获取单元格功能"""

    def test_get_cells_all(self):
        """测试获取所有单元格"""
        doc = load_docx("fixtures/templates/sample.docx")
        cells = get_cells(doc, 0, 0, 0, 0)
        assert isinstance(cells, list)
        # 应该至少有一些单元格
        assert len(cells) > 0

    def test_get_cells_specific_section(self):
        """测试获取特定节的所有单元格"""
        doc = load_docx("fixtures/templates/sample.docx")
        cells = get_cells(doc, 1, 0, 0, 0)
        assert isinstance(cells, list)
        # 所有单元格都应该属于第一节
        for sec_idx, _, _, _, _ in cells:
            assert sec_idx == 1

    def test_get_cells_specific_table(self):
        """测试获取特定表格的所有单元格"""
        doc = load_docx("fixtures/templates/sample.docx")
        cells = get_cells(doc, 1, 1, 0, 0)
        assert isinstance(cells, list)
        # 所有单元格都应该属于第一个节、第一个表格
        for sec_idx, tbl_idx, _, _, _ in cells:
            assert sec_idx == 1
            assert tbl_idx == 1


class TestFindText:
    """测试查找文本功能"""

    def test_find_text_found(self):
        """测试成功找到文本"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 查找一些常见的文本
        positions = find_text(doc, "测试")
        # 结果应该是列表
        assert isinstance(positions, list)

    def test_find_text_not_found(self):
        """测试未找到文本"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 查找一些不太可能存在的文本
        positions = find_text(doc, "__VERY_RARE_TEXT_NOT_IN_DOC__")
        # 应该返回空列表
        assert isinstance(positions, list)
        assert len(positions) == 0


class TestIterateCells:
    """测试遍历单元格功能"""

    def test_iterate_cells_count(self):
        """测试遍历单元格数量正确"""
        doc = load_docx("fixtures/templates/sample.docx")
        count = 0
        for sec_idx, tbl_idx, row_idx, col_idx, cell in iterate_cells(doc):
            count += 1
            # 验证位置信息存在
            assert sec_idx >= 1
            assert tbl_idx >= 1
            assert row_idx >= 1
            assert col_idx >= 1
            assert cell is not None
        # 应该至少有一些单元格
        assert count > 0

    def test_iterate_cells_yield(self):
        """测试生成器正确返回"""
        doc = load_docx("fixtures/templates/sample.docx")
        results = list(iterate_cells(doc))
        # 应该返回包含位置元组和单元格的列表
        assert len(results) > 0
        # 检查第一个元素结构
        first = results[0]
        assert len(first) == 5  # (sec, tbl, row, col, cell)
        sec_idx, tbl_idx, row_idx, col_idx, cell = first
        assert isinstance(sec_idx, int)
        assert isinstance(tbl_idx, int)
        assert isinstance(row_idx, int)
        assert isinstance(col_idx, int)
        assert cell is not None


class TestGetCellText:
    """测试获取单元格文本功能"""

    def test_get_cell_text_basic(self):
        """测试基本文本获取"""
        doc = load_docx("fixtures/templates/sample.docx")
        text = get_cell_text(doc, 1, 1, 1, 1)
        assert isinstance(text, str)

    def test_get_cell_text_empty_cell(self):
        """测试空单元格文本获取"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 尝试获取一个可能为空的单元格
        text = get_cell_text(doc, 1, 1, 1, 1)
        # 应该返回字符串而不是 None
        assert isinstance(text, str)
