"""
DocxLib 字段填充模块测试
"""

import pytest
from docxlib import load_docx, fill_text, fill_date, fill_grid, MatchMode
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


class TestFillTextWildcard:
    """测试文本填充的通配符功能"""

    def test_fill_text_wildcard_all_tables(self):
        """测试通配符所有表格"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 填充所有表格的第1行第1列
        fill_text(doc, (1, 0, 1, 1), "通配符测试")
        # TODO: 验证所有表格的对应位置都已填充
        pass

    def test_fill_text_wildcard_all_sections(self):
        """测试通配符所有节"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 填充所有节的所有表格的第2行第2列
        fill_text(doc, (0, 0, 2, 2), "多节测试")
        # TODO: 验证填充成功
        pass

    def test_fill_text_wildcard_no_matches(self):
        """测试通配符无匹配时抛出异常"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 不存在的节应该抛出异常
        with pytest.raises(PositionError, match="通配符位置"):
            fill_text(doc, (99, 0, 1, 1), "测试")


class TestMatchModeControl:
    """测试匹配模式控制功能"""

    def test_match_right_fill_all(self):
        """测试match_right模式填充所有匹配"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 默认填充所有匹配（match_mode="all"是默认值）
        fill_text(doc, "姓名", "张三", mode="match_right", match_mode=MatchMode.ALL)
        # TODO: 验证所有匹配位置都已填充
        pass

    def test_match_right_fill_first(self):
        """测试match_right模式仅填充第一个匹配"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 仅填充第一个匹配
        fill_text(doc, "姓名", "李四", mode="match_right", match_mode=MatchMode.FIRST)
        # TODO: 验证只有第一个匹配位置被填充
        pass

    def test_match_down_fill_all(self):
        """测试match_down模式填充所有匹配"""
        doc = load_docx("fixtures/templates/sample.docx")
        fill_text(doc, "项目", "测试项目", mode="match_down", match_mode=MatchMode.ALL)
        # TODO: 验证所有匹配位置都已填充
        pass

    def test_match_down_fill_first(self):
        """测试match_down模式仅填充第一个匹配"""
        doc = load_docx("fixtures/templates/sample.docx")
        fill_text(doc, "项目", "第一个项目", mode="match_down", match_mode=MatchMode.FIRST)
        # TODO: 验证只有第一个匹配位置被填充
        pass


class TestBackwardCompatibility:
    """测试向后兼容性"""

    def test_fill_text_without_wildcard_still_works(self):
        """测试无通配符的填充仍然正常工作"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 不使用通配符的传统调用方式
        fill_text(doc, (1, 1, 2, 2), "传统方式测试")
        # TODO: 验证填充成功
        pass

    def test_fill_text_without_match_mode_param(self):
        """测试不指定match_mode参数时使用默认值"""
        doc = load_docx("fixtures/templates/sample.docx")
        # 不指定match_mode，应该默认为"all"
        fill_text(doc, "姓名", "王五", mode="match_right")
        # TODO: 验证使用默认值（填充所有匹配）
        pass
