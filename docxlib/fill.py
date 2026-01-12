"""
DocxLib 字段填充模块

提供文本、图片、日期、网格数据填充等功能。
"""

from pathlib import Path
from typing import Union, List, Tuple

from spire.doc import *
from spire.doc.common import *

from .table import get_cell, find_text
from .style import apply_font_style, parse_color
from .errors import PositionError, FillError
from .constants import DEFAULT_FONT, DEFAULT_FONT_SIZE, DEFAULT_COLOR, FillMode, Position


def fill_text(
    doc: Document,
    position: Union[Position, str],
    value: str,
    mode: str = FillMode.POSITION,
    font_name: str = DEFAULT_FONT,
    font_size: float = DEFAULT_FONT_SIZE,
    color: str = DEFAULT_COLOR,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False
) -> None:
    """填充文本到文档

    Args:
        doc: Document 对象
        position: 位置元组 (section, table, row, col) 或查找文本
        value: 要填充的文本
        mode: 填充模式
            - "position": 直接定位
            - "match_right": 查找文本，填充到右侧
            - "match_down": 查找文本，填充到下方
        font_name: 字体名称
        font_size: 字体大小（磅）
        color: 颜色（名称或十六进制）
        bold: 是否粗体
        italic: 是否斜体
        underline: 是否下划线

    Raises:
        PositionError: 位置无效
        FillError: 填充失败

    Examples:
        >>> # 直接定位
        >>> fill_text(doc, (1, 1, 2, 2), "测试文本")

        >>> # 右侧填充
        >>> fill_text(doc, "姓名：", "张三", mode="match_right")

        >>> # 下方填充
        >>> fill_text(doc, "项目1", "智慧城市", mode="match_down")

        >>> # 带样式
        >>> fill_text(doc, "标题", "内容", font_name="黑体", font_size=16, bold=True)
    """
    try:
        # 确定目标单元格位置
        if mode == FillMode.POSITION:
            if isinstance(position, str):
                raise PositionError("position 模式需要位置元组，不是字符串")
            target_pos = position
        elif mode == FillMode.MATCH_RIGHT:
            if not isinstance(position, str):
                raise PositionError("match_right 模式需要查找文本字符串")
            positions = find_text(doc, position)
            if not positions:
                raise PositionError(f"未找到文本: {position}")
            # 取第一个匹配位置，列索引+1
            pos = positions[0]
            target_pos = (pos[0], pos[1], pos[2], pos[3] + 1)
        elif mode == FillMode.MATCH_DOWN:
            if not isinstance(position, str):
                raise PositionError("match_down 模式需要查找文本字符串")
            positions = find_text(doc, position)
            if not positions:
                raise PositionError(f"未找到文本: {position}")
            # 取第一个匹配位置，行索引+1
            pos = positions[0]
            target_pos = (pos[0], pos[1], pos[2] + 1, pos[3])
        else:
            raise FillError(f"不支持的填充模式: {mode}")

        # 获取单元格
        cell = get_cell(doc, *target_pos)

        # 清空单元格内容
        cell.Paragraphs.Clear()

        # 添加段落并设置文本
        paragraph = cell.AddParagraph()
        run = paragraph.AppendText(value)

        # 应用样式
        apply_font_style(
            run, font_name, font_size, color,
            bold, italic, underline
        )

    except (PositionError, FillError):
        raise
    except Exception as e:
        raise FillError(f"填充文本失败: {e}")


def fill_image(
    doc: Document,
    position: Union[Position, str],
    image_path: str,
    mode: str = FillMode.POSITION,
    width: float = None,
    height: float = None,
    maintain_ratio: bool = True
) -> None:
    """填充图片到文档

    Args:
        doc: Document 对象
        position: 位置元组或查找文本
        image_path: 图片文件路径
        mode: 填充模式（同 fill_text）
        width: 宽度（磅）
        height: 高度（磅）
        maintain_ratio: 是否保持宽高比

    Raises:
        FillError: 图片文件不存在或格式不支持

    Examples:
        >>> fill_image(doc, (1, 1, 2, 2), "logo.png")

        >>> fill_image(doc, "印章：", "seal.png", mode="match_right", width=100, height=100)
    """
    try:
        # 验证图片文件
        img_path = Path(image_path)
        if not img_path.exists():
            raise FillError(f"图片文件不存在: {image_path}")

        # 确定目标单元格位置
        if mode == FillMode.POSITION:
            if isinstance(position, str):
                raise PositionError("position 模式需要位置元组")
            target_pos = position
        elif mode == FillMode.MATCH_RIGHT:
            positions = find_text(doc, position)
            if not positions:
                raise PositionError(f"未找到文本: {position}")
            pos = positions[0]
            target_pos = (pos[0], pos[1], pos[2], pos[3] + 1)
        elif mode == FillMode.MATCH_DOWN:
            positions = find_text(doc, position)
            if not positions:
                raise PositionError(f"未找到文本: {position}")
            pos = positions[0]
            target_pos = (pos[0], pos[1], pos[2] + 1, pos[3])
        else:
            raise FillError(f"不支持的填充模式: {mode}")

        # 获取单元格
        cell = get_cell(doc, *target_pos)

        # 清空单元格
        cell.Paragraphs.Clear()

        # 添加段落
        paragraph = cell.AddParagraph()

        # 加载图片
        from spire.doc import DocPicture
        picture = paragraph.AppendPicture()

        # 加载图片文件
        picture.LoadImage(image_path)

        # 调整图片大小
        if width is not None or height is not None:
            original_width = picture.Width
            original_height = picture.Height

            if width is None and height is not None:
                # 只指定高度
                if maintain_ratio:
                    width = original_width * (height / original_height)
                else:
                    width = original_width
            elif height is None and width is not None:
                # 只指定宽度
                if maintain_ratio:
                    height = original_height * (width / original_width)
                else:
                    height = original_height

            if width is not None:
                picture.Width = width
            if height is not None:
                picture.Height = height

    except (PositionError, FillError):
        raise
    except Exception as e:
        raise FillError(f"填充图片失败: {e}")


def fill_date(
    doc: Document,
    position: Union[Position, str],
    date_str: str,
    font_name: str = DEFAULT_FONT,
    font_size: float = DEFAULT_FONT_SIZE
) -> None:
    """填充日期

    数字和年月日使用不同字体：数字部分使用 font_name，
    年月日部分使用宋体。

    Args:
        doc: Document 对象
        position: 位置元组或查找文本
        date_str: 日期字符串，如 "2024年1月15日"
        font_name: 数字字体（年月日使用宋体）
        font_size: 字体大小

    Examples:
        >>> fill_date(doc, (1, 1, 4, 2), "2024年1月15日")
        >>> fill_date(doc, "日期：", "2024年1月15日", mode="match_right")
    """
    try:
        # 解析日期字符串
        from .utils import parse_date_string
        numbers, separators = parse_date_string(date_str)

        if not numbers or not separators:
            # 如果解析失败，直接填充文本
            return fill_text(doc, position, date_str, font_name=font_name, font_size=font_size)

        # 确定目标单元格位置
        if isinstance(position, str):
            positions = find_text(doc, position)
            if not positions:
                raise PositionError(f"未找到文本: {position}")
            # 默认右侧填充
            pos = positions[0]
            target_pos = (pos[0], pos[1], pos[2], pos[3] + 1)
        else:
            target_pos = position

        # 获取单元格
        cell = get_cell(doc, *target_pos)

        # 清空单元格
        cell.Paragraphs.Clear()

        # 添加段落
        paragraph = cell.AddParagraph()

        # 依次添加数字和年月日
        for i, (num, sep) in enumerate(zip(numbers, separators)):
            # 添加数字（使用指定字体）
            run_num = paragraph.AppendText(num)
            apply_font_style(run_num, font_name, font_size, DEFAULT_COLOR)

            # 添加年月日（使用宋体）
            run_sep = paragraph.AppendText(sep)
            apply_font_style(run_sep, "宋体", font_size, DEFAULT_COLOR)

    except (PositionError, FillError):
        raise
    except Exception as e:
        raise FillError(f"填充日期失败: {e}")


def fill_grid(
    doc: Document,
    data: List[List[str]],
    position: Position
) -> None:
    """填充网格数据

    从二维数组填充数据到表格。

    Args:
        doc: Document 对象
        data: 二维数组，每个元素代表一个单元格的值
        position: 起始位置 (section, table, row, col)

    Raises:
        PositionError: 数据超出表格边界

    Examples:
        >>> data = [
        ...     ["序号", "项目", "金额"],
        ...     ["1", "设备费", "50000"],
        ...     ["2", "人工费", "30000"],
        ... ]
        >>> fill_grid(doc, data, position=(1, 1, 7, 1))
    """
    try:
        section_idx, table_idx, start_row, start_col = position

        # 填充数据
        for row_idx, row_data in enumerate(data):
            for col_idx, cell_value in enumerate(row_data):
                # 计算目标位置
                target_row = start_row + row_idx
                target_col = start_col + col_idx

                # 填充单元格
                try:
                    cell = get_cell(
                        doc,
                        section_idx,
                        table_idx,
                        target_row,
                        target_col
                    )

                    # 清空并设置文本
                    cell.Paragraphs.Clear()
                    paragraph = cell.AddParagraph()
                    paragraph.AppendText(str(cell_value))

                except PositionError:
                    raise PositionError(
                        f"数据超出表格边界: "
                        f"无法填充到 ({section_idx}, {table_idx}, {target_row}, {target_col})"
                    )

    except PositionError:
        raise
    except Exception as e:
        raise FillError(f"填充网格数据失败: {e}")


def replace_all(doc: Document, old_text: str, new_text: str) -> None:
    """全局替换文档中的文本

    Args:
        doc: Document 对象
        old_text: 要查找的文本
        new_text: 替换的文本

    Examples:
        >>> replace_all(doc, "{合同编号}", "HT-2024-001")
        >>> replace_all(doc, "{甲方}", "某某公司")
    """
    try:
        doc.Replace(old_text, new_text, False, False)
    except Exception as e:
        raise FillError(f"全局替换失败: {e}")


def clear_cell(doc: Document, section: int, table: int,
               row: int, col: int) -> None:
    """清空单元格内容

    Args:
        doc: Document 对象
        section: 节索引（从1开始）
        table: 表格索引（从1开始）
        row: 行索引（从1开始）
        col: 列索引（从1开始）

    Examples:
        >>> clear_cell(doc, 1, 1, 2, 2)
    """
    try:
        cell = get_cell(doc, section, table, row, col)
        cell.Paragraphs.Clear()
    except Exception as e:
        raise FillError(f"清空单元格失败: {e}")
