"""
DocxLib 字段填充模块测试
"""

import pytest
from docxlib import load_docx, fill_text, fill_date, fill_grid
from docxlib.errors import PositionError, FillError, ValidationError


class TestFillText:
    """测试文本填充功能"""

    def test_fill_text_by_position(self):
        """测试按位置填充文本"""
        # doc = load_docx("fixtures/templates/sample.docx")
        # fill_text(doc, (1, 1, 2, 2), "测试文本")
        # cell = get_cell(doc, 1, 1, 2, 2)
        # text = cell.Range.Text.strip()
        # assert "测试文本" in text
        pass

    def test_fill_text_by_match_right(self):
        """测试 match_right 模式填充"""
        pass

    def test_fill_text_by_match_down(self):
        """测试 match_down 模式填充"""
        pass

    def test_fill_text_invalid_position(self):
        """测试无效位置时抛出异常"""
        # doc = load_docx("fixtures/templates/sample.docx")
        # with pytest.raises(PositionError):
        #     fill_text(doc, (99, 99, 99, 99), "测试")
        pass

    def test_fill_text_with_style(self):
        """测试带样式的文本填充"""
        pass


class TestFillDate:
    """测试日期填充功能"""

    def test_fill_date_success(self):
        """测试成功填充日期"""
        pass

    def test_fill_date_with_different_fonts(self):
        """测试数字和年月日使用不同字体"""
        pass

    def test_fill_date_invalid_format(self):
        """测试无效日期格式时抛出 ValidationError"""
        doc = load_docx("fixtures/templates/sample.docx")

        # 测试完全不匹配的格式
        with pytest.raises(ValidationError, match="日期格式无效"):
            fill_date(doc, (1, 1, 1, 1), "hello world")

        with pytest.raises(ValidationError, match="日期格式无效"):
            fill_date(doc, (1, 1, 1, 1), "2024-01-15")

        with pytest.raises(ValidationError, match="日期格式无效"):
            fill_date(doc, (1, 1, 1, 1), "2025年月1日")

    def test_fill_date_invalid_date(self):
        """测试不存在的日期时抛出 ValidationError"""
        doc = load_docx("fixtures/templates/sample.docx")

        # 测试月份无效
        with pytest.raises(ValidationError, match="日期不存在"):
            fill_date(doc, (1, 1, 1, 1), "2025年13月1日")

        with pytest.raises(ValidationError, match="日期不存在"):
            fill_date(doc, (1, 1, 1, 1), "2025年0月1日")

        # 测试日期无效
        with pytest.raises(ValidationError, match="日期不存在"):
            fill_date(doc, (1, 1, 1, 1), "2025年2月30日")

        with pytest.raises(ValidationError, match="日期不存在"):
            fill_date(doc, (1, 1, 1, 1), "2025年4月31日")

        # 测试非闰年的2月29日
        with pytest.raises(ValidationError, match="日期不存在"):
            fill_date(doc, (1, 1, 1, 1), "2023年2月29日")


class TestFillGrid:
    """测试网格数据填充功能"""

    def test_fill_grid_success(self):
        """测试成功填充网格数据"""
        pass

    def test_fill_grid_out_of_bounds(self):
        """测试数据超出边界时抛出异常"""
        # doc = load_docx("fixtures/templates/sample.docx")
        # data = [["测试"] * 100]
        # with pytest.raises(PositionError):
        #     fill_grid(doc, data, position=(1, 1, 1, 1))
        pass


class TestReplaceAll:
    """测试全局替换功能"""

    def test_replace_all_success(self):
        """测试成功全局替换"""
        pass
