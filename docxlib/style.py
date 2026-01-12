"""
DocxLib 样式管理模块

提供颜色解析、字体样式应用等功能。
"""

import re
from typing import Union

from spire.doc import Color
from spire.doc.common import *

from .constants import COLOR_MAP


def parse_color(color_str: str) -> Color:
    """解析颜色字符串为 Color 对象

    支持颜色名称和十六进制格式。解析失败时返回黑色。

    Args:
        color_str: 颜色字符串
            - 颜色名称: 'black', 'red', 'blue' 等
            - 十六进制: '#FF0000' 或 'FF0000'

    Returns:
        Color: Spire.Doc Color 对象

    Examples:
        >>> parse_color('red')
        <Color object>
        >>> parse_color('#FF0000')
        <Color object>
        >>> parse_color('FF0000')
        <Color object>
    """
    # 预处理：去除空格，转小写
    color_str = color_str.strip().lower()

    # 检查颜色名称映射
    if color_str in COLOR_MAP:
        hex_str = COLOR_MAP[color_str]
    else:
        hex_str = color_str

    # 去除 # 前缀
    if hex_str.startswith('#'):
        hex_str = hex_str[1:]

    # 验证十六进制格式
    if not re.match(r'^[0-9a-f]{6}$', hex_str):
        # 解析失败，返回黑色
        return Color.get_Black()

    # 解析 RGB
    try:
        r = int(hex_str[0:2], 16)
        g = int(hex_str[2:4], 16)
        b = int(hex_str[4:6], 16)
        # 使用 FromArgb 创建颜色 (alpha, red, green, blue)
        return Color.FromArgb(255, r, g, b)
    except (ValueError, IndexError):
        # 解析失败，返回黑色
        return Color.get_Black()


def apply_font_style(
    run,
    font_name: str,
    font_size: float,
    color: str,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False
) -> None:
    """应用字体样式到文本范围

    Args:
        run: Spire.Doc TextRange 对象
        font_name: 字体名称
        font_size: 字体大小（磅）
        color: 颜色（名称或十六进制）
        bold: 是否粗体
        italic: 是否斜体
        underline: 是否下划线

    Examples:
        >>> apply_font_style(
        ...     run,
        ...     font_name="黑体",
        ...     font_size=16,
        ...     color="red",
        ...     bold=True
        ... )
    """
    # 设置字体名称
    if font_name:
        run.CharacterFormat.FontName = font_name

    # 设置字体大小
    if font_size > 0:
        run.CharacterFormat.FontSize = font_size

    # 设置颜色
    if color:
        run.CharacterFormat.TextColor = parse_color(color)

    # 设置粗体
    run.CharacterFormat.Bold = bold

    # 设置斜体
    run.CharacterFormat.Italic = italic

    # 设置下划线
    if underline:
        from spire.doc import UnderlineStyle
        run.CharacterFormat.UnderlineStyle = UnderlineStyle.Single


def get_cell_format(cell):
    """获取单元格格式对象

    Args:
        cell: Spire.Doc Cell 对象

    Returns:
        CellFormat: 单元格格式对象
    """
    return cell.CellFormat


def set_cell_border(
    cell,
    border_style=None,
    border_color=None,
    border_width=None
) -> None:
    """设置单元格边框样式

    Args:
        cell: Spire.Doc Cell 对象
        border_style: 边框样式（可选）
        border_color: 边框颜色（可选）
        border_width: 边框宽度（可选）

    Examples:
        >>> set_cell_border(
        ...     cell,
        ...     border_color="black",
        ...     border_width=0.5
        ... )
    """
    borders = cell.CellFormat.Borders

    if border_color:
        color = parse_color(border_color)
        borders.BorderType.Left.Color = color
        borders.BorderType.Right.Color = color
        borders.BorderType.Top.Color = color
        borders.BorderType.Bottom.Color = color

    if border_width is not None:
        borders.BorderType.Left.LineWidth = border_width
        borders.BorderType.Right.LineWidth = border_width
        borders.BorderType.Top.LineWidth = border_width
        borders.BorderType.Bottom.LineWidth = border_width
