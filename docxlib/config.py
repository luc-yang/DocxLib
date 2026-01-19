"""
DocxLib 参数对象配置模块

使用 dataclass 实现参数对象模式，将函数的多个相关参数封装成配置对象。
这样可以提高代码的可读性、可维护性和可复用性。
"""

from dataclasses import dataclass
from typing import Literal, Optional

from docxlib.constants import (
    DEFAULT_COLOR,
    DEFAULT_FONT_FAMILY,
    DEFAULT_FONT_SIZE,
)

# 对齐方式的字面量类型
HorizontalAlignmentLiteral = Literal["left", "center", "right", "justify"]
VerticalAlignmentLiteral = Literal["top", "middle", "bottom"]
FillModeLiteral = Literal["position", "match_right", "match_down"]
MatchModeLiteral = Literal["all", "first"]


@dataclass
class FontStyle:
    """字体样式配置

    封装所有字体相关的样式属性，提供灵活的字体样式配置。

    Attributes:
        font_name: 字体名称，默认"仿宋_GB2312"
        font_size: 字体大小（磅），默认10.5
        color: 颜色（名称或十六进制），默认"black"
        bold: 是否粗体，默认False
        italic: 是否斜体，默认False
        underline: 是否下划线，默认False

    Examples:
        >>> # 使用默认样式
        >>> style = FontStyle()
        >>> # 自定义样式
        >>> style = FontStyle(font_name="黑体", font_size=16, bold=True)
        >>> # 使用预设样式
        >>> title_style = FontStyle.title()
        >>> heading_style = FontStyle.heading(level=2)
    """

    font_name: str = DEFAULT_FONT_FAMILY
    font_size: float = DEFAULT_FONT_SIZE
    color: str = DEFAULT_COLOR
    bold: bool = False
    italic: bool = False
    underline: bool = False

    @classmethod
    def title(cls) -> "FontStyle":
        """预设：标题样式（黑体 16磅 粗体）

        Returns:
            标题样式配置对象

        Examples:
            >>> fill_text(doc, (1, 1, 1, 1), "文档标题", style=FontStyle.title())
        """
        return cls(font_name="黑体", font_size=16, bold=True)

    @classmethod
    def heading(cls, level: int = 1) -> "FontStyle":
        """预设：分级标题样式

        Args:
            level: 标题级别（1-3），默认为1

        Returns:
            标题样式配置对象

        Examples:
            >>> # 一级标题
            >>> fill_text(doc, (1, 1, 1, 1), "第一章", style=FontStyle.heading(1))
            >>> # 二级标题
            >>> fill_text(doc, (1, 1, 2, 1), "1.1 概述", style=FontStyle.heading(2))
        """
        sizes = {1: 18, 2: 16, 3: 14}
        return cls(font_name="黑体", font_size=sizes.get(level, 16), bold=True)

    @classmethod
    def body(cls) -> "FontStyle":
        """预设：正文样式（仿宋_GB2312 12磅）

        Returns:
            正文样式配置对象

        Examples:
            >>> fill_text(doc, (1, 1, 3, 1), "正文内容...", style=FontStyle.body())
        """
        return cls(font_name="仿宋_GB2312", font_size=12)

    @classmethod
    def emphasis(cls) -> "FontStyle":
        """预设：强调样式（红色粗体）

        Returns:
            强调样式配置对象

        Examples:
            >>> fill_text(doc, (1, 1, 4, 1), "重要提示", style=FontStyle.emphasis())
        """
        return cls(bold=True, color="red")


@dataclass
class Alignment:
    """对齐方式配置

    封装水平和垂直对齐方式的配置。

    Attributes:
        h_align: 水平对齐方式（left/center/right/justify）
        v_align: 垂直对齐方式（top/middle/bottom）

    Examples:
        >>> # 使用默认对齐
        >>> align = Alignment()
        >>> # 居中对齐
        >>> align = Alignment(h_align="center", v_align="middle")
        >>> # 使用预设样式
        >>> center_align = Alignment.center()
    """

    h_align: Optional[HorizontalAlignmentLiteral] = None
    v_align: Optional[VerticalAlignmentLiteral] = None

    @classmethod
    def center(cls) -> "Alignment":
        """预设：居中对齐（水平居中，垂直居中）

        Returns:
            居中对齐配置对象

        Examples:
            >>> fill_text(doc, (1, 1, 1, 1), "居中文字", alignment=Alignment.center())
        """
        return cls(h_align="center", v_align="middle")

    @classmethod
    def left_top(cls) -> "Alignment":
        """预设：左上对齐

        Returns:
            左上对齐配置对象

        Examples:
            >>> fill_text(doc, (1, 1, 1, 1), "左上对齐", alignment=Alignment.left_top())
        """
        return cls(h_align="left", v_align="top")

    @classmethod
    def right_bottom(cls) -> "Alignment":
        """预设：右下对齐

        Returns:
            右下对齐配置对象

        Examples:
            >>> fill_text(doc, (1, 1, 1, 1), "右下对齐", alignment=Alignment.right_bottom())
        """
        return cls(h_align="right", v_align="bottom")


@dataclass
class FillOptions:
    """填充模式配置

    封装填充模式和匹配模式的配置，用于控制文本和图片填充的行为。

    Attributes:
        mode: 填充模式（position/match_right/match_down）
            - "position": 直接定位填充
            - "match_right": 查找文本，填充到右侧
            - "match_down": 查找文本，填充到下方
        match_mode: 匹配模式（all/first）
            - "all": 填充所有匹配位置（默认）
            - "first": 仅填充第一个匹配位置

    Examples:
        >>> # 使用默认模式（直接定位）
        >>> options = FillOptions()
        >>> # 向右匹配模式
        >>> options = FillOptions.match_right()
        >>> # 向下匹配，仅第一个
        >>> options = FillOptions.match_down(match_mode="first")
    """

    mode: FillModeLiteral = "position"
    match_mode: MatchModeLiteral = "all"

    @classmethod
    def match_right(cls, match_mode: MatchModeLiteral = "all") -> "FillOptions":
        """预设：向右匹配模式

        查找指定文本，将内容填充到其右侧单元格。

        Args:
            match_mode: 匹配模式，默认为"all"（填充所有匹配）

        Returns:
            向右匹配模式配置对象

        Examples:
            >>> fill_text(
            ...     doc, "姓名：", "张三",
            ...     options=FillOptions.match_right()
            ... )
        """
        return cls(mode="match_right", match_mode=match_mode)

    @classmethod
    def match_down(cls, match_mode: MatchModeLiteral = "all") -> "FillOptions":
        """预设：向下匹配模式

        查找指定文本，将内容填充到其下方单元格。

        Args:
            match_mode: 匹配模式，默认为"all"（填充所有匹配）

        Returns:
            向下匹配模式配置对象

        Examples:
            >>> fill_text(
            ...     doc, "项目", "智慧城市",
            ...     options=FillOptions.match_down(match_mode="first")
            ... )
        """
        return cls(mode="match_down", match_mode=match_mode)


@dataclass
class ImageConfig:
    """图片配置

    封装图片插入时的布局和尺寸配置。

    Attributes:
        h_align: 水平对齐方式（left/center/right/justify）
        v_align: 垂直对齐方式（top/middle/bottom）
        width: 宽度（磅），None表示自动
        height: 高度（磅），None表示自动
        maintain_ratio: 是否保持宽高比，默认True

    Examples:
        >>> # 使用默认配置
        >>> config = ImageConfig()
        >>> # 指定尺寸（保持比例）
        >>> config = ImageConfig(width=80, height=80)
        >>> # 固定尺寸（不保持比例）
        >>> config = ImageConfig.fixed_size(100, 100)
        >>> # 居中图片
        >>> config = ImageConfig.centered(width=80, height=80)
    """

    h_align: Optional[HorizontalAlignmentLiteral] = None
    v_align: Optional[VerticalAlignmentLiteral] = None
    width: Optional[float] = None
    height: Optional[float] = None
    maintain_ratio: bool = True

    @classmethod
    def centered(cls, width: float = None, height: float = None) -> "ImageConfig":
        """预设：居中图片

        创建一个居中对齐的图片配置。

        Args:
            width: 宽度（磅），可选
            height: 高度（磅），可选

        Returns:
            居中图片配置对象

        Examples:
            >>> fill_image(
            ...     doc, "照片", "logo.png",
            ...     config=ImageConfig.centered(width=80, height=80)
            ... )
        """
        return cls(h_align="center", v_align="middle", width=width, height=height)

    @classmethod
    def fixed_size(cls, width: float, height: float) -> "ImageConfig":
        """预设：固定尺寸（不保持比例）

        创建一个固定尺寸的图片配置，不保持宽高比。

        Args:
            width: 宽度（磅）
            height: 高度（磅）

        Returns:
            固定尺寸图片配置对象

        Examples:
            >>> fill_image(
            ...     doc, (1, 1, 1, 1), "logo.png",
            ...     config=ImageConfig.fixed_size(100, 100)
            ... )
        """
        return cls(width=width, height=height, maintain_ratio=False)
