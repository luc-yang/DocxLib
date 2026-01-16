"""
DocxLib 样式管理模块

提供颜色解析、字体样式应用等功能。
"""

import re
from typing import Union

from spire.doc import Color
from spire.doc.common import *

from .constants import COLOR_MAP
from .errors import FillError


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
    if hex_str.startswith("#"):
        hex_str = hex_str[1:]

    # 验证十六进制格式
    if not re.match(r"^[0-9a-f]{6}$", hex_str):
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
    underline: bool = False,
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
    cell, border_style=None, border_color=None, border_width=None
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


def apply_paragraph_alignment(paragraph, alignment: str) -> None:
    """应用段落对齐方式

    Args:
        paragraph: Spire.Doc Paragraph 对象
        alignment: 对齐方式
            - "left": 左对齐
            - "center": 居中对齐
            - "right": 右对齐
            - "justify": 两端对齐

    Examples:
        >>> apply_paragraph_alignment(paragraph, "center")
        >>> apply_paragraph_alignment(paragraph, Alignment.RIGHT)
    """
    from spire.doc import HorizontalAlignment

    alignment_map = {
        "left": HorizontalAlignment.Left,
        "center": HorizontalAlignment.Center,
        "right": HorizontalAlignment.Right,
        "justify": HorizontalAlignment.Justify,
    }

    if alignment in alignment_map:
        paragraph.Format.HorizontalAlignment = alignment_map[alignment]


def apply_cell_alignment(cell, alignment: str) -> None:
    """应用单元格对齐方式

    Args:
        cell: Spire.Doc Cell 对象
        alignment: 对齐方式
            - "left": 左对齐
            - "center": 居中对齐
            - "right": 右对齐
            - "justify": 两端对齐

    Examples:
        >>> apply_cell_alignment(cell, "center")
    """
    from spire.doc import VerticalAlignment

    alignment_map = {
        "top": VerticalAlignment.Top,
        "center": VerticalAlignment.Middle,
        "bottom": VerticalAlignment.Bottom,
    }

    if alignment in alignment_map:
        cell.CellFormat.VerticalAlignment = alignment_map[alignment]


def get_cell_style(cell) -> dict:
    """获取单元格样式信息

    Args:
        cell: Spire.Doc Cell 对象

    Returns:
        dict: 样式信息字典
        {
            "font_name": "字体名称",
            "font_size": 12.0,
            "bold": True,
            "italic": False,
            "underline": False,
            "color": "#000000",
            "h_align": "left",
            "v_align": "middle",
            "background_color": "#FFFFFF",
        }

    Examples:
        >>> cell = get_cell(doc, 1, 1, 2, 2)
        >>> style = get_cell_style(cell)
        >>> print(f"字体: {style['font_name']}")
    """
    try:
        if cell.Paragraphs.Count == 0:
            return {}

        # 获取第一个段落
        paragraph = cell.Paragraphs.get_Item(0)

        # 读取段落格式（包含字符格式）
        font_name = ""
        font_size = 0
        bold = False
        italic = False
        underline = False
        color = ""

        try:
            # 尝试从段落格式中获取字符格式
            if hasattr(paragraph, "Format") and paragraph.Format:
                format_obj = paragraph.Format

                # 尝试获取字符格式属性
                if hasattr(format_obj, "CharacterFormat"):
                    char_format = format_obj.CharacterFormat
                else:
                    # 如果没有 CharacterFormat，尝试直接访问
                    char_format = None

                # 如果还是没有，使用默认值
                if not char_format:
                    char_format = None
        except Exception:
            char_format = None

        # 简化：只返回对齐和背景色信息
        # 读取对齐方式
        h_align = ""
        try:
            format_obj = paragraph.Format
            if format_obj.HorizontalAlignment:
                align_map = {
                    "left": "left",
                    "center": "center",
                    "right": "right",
                    "justify": "justify",
                }
                h_align = align_map.get(str(format_obj.HorizontalAlignment), "")
        except Exception:
            pass

        # 读取垂直对齐
        v_align = ""
        try:
            cell_format = cell.CellFormat
            if cell_format.VerticalAlignment:
                v_align_map = {
                    "top": "top",
                    "middle": "middle",
                    "bottom": "bottom",
                }
                v_align = v_align_map.get(str(cell_format.VerticalAlignment), "")
        except Exception:
            pass

        # 读取背景色
        background_color = ""
        try:
            cell_format = cell.CellFormat
            if cell_format.BackColor:
                background_color = f"#{cell_format.BackColor.Name.replace('#', '')}"
        except Exception:
            pass

        return {
            "font_name": font_name,
            "font_size": font_size,
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "color": color,
            "h_align": h_align,
            "v_align": v_align,
            "background_color": background_color,
        }

    except Exception as e:
        raise FillError(f"获取单元格样式失败: {e}")


def get_paragraph_style(paragraph) -> dict:
    """获取段落样式信息

    Args:
        paragraph: Spire.Doc Paragraph 对象

    Returns:
        dict: 样式信息字典
        {
            "alignment": "center",
            "first_line_indent": 24.0,
            "line_spacing": 1.5,
            "space_before": 12.0,
            "space_after": 12.0,
        }

    Examples:
        >>> para = doc.Sections.get_Item(0).Paragraphs.get_Item(0)
        >>> style = get_paragraph_style(para)
        >>> print(f"对齐方式: {style['alignment']}")
    """
    try:
        format_obj = paragraph.Format

        # 对齐方式
        alignment = ""
        try:
            if format_obj.HorizontalAlignment:
                align_map = {
                    "left": "left",
                    "center": "center",
                    "right": "right",
                    "justify": "justify",
                }
                alignment = align_map.get(str(format_obj.HorizontalAlignment), "")
        except Exception:
            pass

        # 首行缩进
        first_line_indent = 0.0
        try:
            if format_obj.FirstLineIndent:
                first_line_indent = float(format_obj.FirstLineIndent)
        except Exception:
            pass

        # 行距
        line_spacing = 0.0
        try:
            if format_obj.LineSpacing:
                line_spacing = float(format_obj.LineSpacing)
        except Exception:
            pass

        # 段前间距
        space_before = 0.0
        try:
            if format_obj.SpaceBefore:
                space_before = float(format_obj.SpaceBefore)
        except Exception:
            pass

        # 段后间距
        space_after = 0.0
        try:
            if format_obj.SpaceAfter:
                space_after = float(format_obj.SpaceAfter)
        except Exception:
            pass

        return {
            "alignment": alignment,
            "first_line_indent": first_line_indent,
            "line_spacing": line_spacing,
            "space_before": space_before,
            "space_after": space_after,
        }

    except Exception as e:
        raise FillError(f"获取段落样式失败: {e}")
