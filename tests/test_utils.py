"""
DocxLib 工具函数测试
"""

import pytest
from docxlib.utils import validate_date_string, parse_date_string
from docxlib.errors import ValidationError


class TestValidateDateString:
    """测试日期字符串验证功能"""

    def test_validate_date_valid_dates(self):
        """测试有效日期"""
        # 标准格式
        validate_date_string("2024年1月15日")
        validate_date_string("2024年12月31日")

        # 补零格式
        validate_date_string("2024年01月15日")
        validate_date_string("2024年12月31日")

        # 闰年日期
        validate_date_string("2024年2月29日")

        # 不同月份的最后一天
        validate_date_string("2024年1月31日")
        validate_date_string("2024年3月31日")
        validate_date_string("2024年4月30日")
        validate_date_string("2024年5月31日")
        validate_date_string("2024年6月30日")

    def test_validate_date_invalid_format(self):
        """测试无效格式"""
        # 完全不匹配
        with pytest.raises(ValidationError, match="日期格式无效"):
            validate_date_string("hello")

        with pytest.raises(ValidationError, match="日期格式无效"):
            validate_date_string("2024-01-15")

        with pytest.raises(ValidationError, match="日期格式无效"):
            validate_date_string("2024/01/15")

        with pytest.raises(ValidationError, match="日期格式无效"):
            validate_date_string("2025年月1日")  # 缺少月份

        with pytest.raises(ValidationError, match="日期格式无效"):
            validate_date_string("2025年1月日")  # 缺少日期

        with pytest.raises(ValidationError, match="日期格式无效"):
            validate_date_string("年1月1日")  # 缺少年份

        # 格式不完整
        with pytest.raises(ValidationError, match="日期格式无效"):
            validate_date_string("2025年1月")

        with pytest.raises(ValidationError, match="日期格式无效"):
            validate_date_string("2025年")

    def test_validate_date_invalid_month(self):
        """测试无效月份"""
        # 月份为 0
        with pytest.raises(ValidationError, match="日期不存在"):
            validate_date_string("2025年0月1日")

        # 月份大于 12
        with pytest.raises(ValidationError, match="日期不存在"):
            validate_date_string("2025年13月1日")

        with pytest.raises(ValidationError, match="日期不存在"):
            validate_date_string("2025年24月1日")

    def test_validate_date_invalid_day(self):
        """测试无效日期"""
        # 2月没有 30 日
        with pytest.raises(ValidationError, match="日期不存在"):
            validate_date_string("2025年2月30日")

        # 非闰年的 2月没有 29 日
        with pytest.raises(ValidationError, match="日期不存在"):
            validate_date_string("2023年2月29日")

        # 4月只有 30 日
        with pytest.raises(ValidationError, match="日期不存在"):
            validate_date_string("2025年4月31日")

        # 6月只有 30 日
        with pytest.raises(ValidationError, match="日期不存在"):
            validate_date_string("2025年6月31日")

        # 9月只有 30 日
        with pytest.raises(ValidationError, match="日期不存在"):
            validate_date_string("2025年9月31日")

        # 11月只有 30 日
        with pytest.raises(ValidationError, match="日期不存在"):
            validate_date_string("2025年11月31日")

        # 日期为 0
        with pytest.raises(ValidationError, match="日期不存在"):
            validate_date_string("2025年1月0日")

    def test_validate_date_edge_cases(self):
        """测试边界情况"""
        # 1月1日（年初）
        validate_date_string("2024年1月1日")

        # 12月31日（年末）
        validate_date_string("2024年12月31日")

        # 小月30日
        validate_date_string("2024年4月30日")
        validate_date_string("2024年6月30日")
        validate_date_string("2024年9月30日")
        validate_date_string("2024年11月30日")

        # 大月31日
        validate_date_string("2024年1月31日")
        validate_date_string("2024年3月31日")
        validate_date_string("2024年5月31日")
        validate_date_string("2024年7月31日")
        validate_date_string("2024年8月31日")
        validate_date_string("2024年10月31日")
        validate_date_string("2024年12月31日")


class TestParseDateString:
    """测试日期字符串解析功能"""

    def test_parse_date_standard_format(self):
        """测试标准格式解析"""
        numbers, separators = parse_date_string("2024年1月15日")

        assert numbers == ["2024", "01", "15"]
        assert separators == ["年", "月", "日"]

    def test_parse_date_with_zero_padding(self):
        """测试补零"""
        numbers, separators = parse_date_string("2024年1月5日")

        assert numbers == ["2024", "01", "05"]
        assert separators == ["年", "月", "日"]

    def test_parse_date_already_padded(self):
        """测试已经补零的日期"""
        numbers, separators = parse_date_string("2024年01月05日")

        assert numbers == ["2024", "01", "05"]
        assert separators == ["年", "月", "日"]

    def test_parse_date_invalid_format(self):
        """测试无效格式返回空列表"""
        # 完全不匹配
        numbers, separators = parse_date_string("hello")
        assert numbers == []
        assert separators == []

        # 错误的分隔符
        numbers, separators = parse_date_string("2024-01-15")
        assert numbers == []
        assert separators == []

        # 不完整的格式
        numbers, separators = parse_date_string("2024年1月")
        assert numbers == ["2024", "01"]
        assert separators == ["年", "月"]
