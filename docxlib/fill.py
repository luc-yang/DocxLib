"""
DocxLib 字段填充模块

提供文本、图片、日期、网格数据填充等功能。
"""

from pathlib import Path
from typing import Any, Dict, List, Tuple, Union
import re

from spire.doc import *
from spire.doc.common import *

from .constants import (
    DEFAULT_COLOR,
    DEFAULT_FONT,
    DEFAULT_FONT_SIZE,
    DEFAULT_MISSING_VAR_ACTION,
    DEFAULT_VAR_PREFIX,
    DEFAULT_VAR_SUFFIX,
    FillMode,
    HorizontalAlignment,
    MatchMode,
    Position,
    VerticalAlignment,
)
from .errors import FillError, PositionError, ValidationError, VariableNotFoundError
from .style import apply_cell_alignment, apply_font_style, apply_paragraph_alignment
from .table import find_text, get_cell, get_cells


def _has_wildcard(position: Position) -> bool:
    """检查位置元组是否包含通配符

    Args:
        position: 位置元组 (section, table, row, col)

    Returns:
        bool: 如果任何维度为 0（通配符）则返回 True
    """
    return 0 in position


def _fill_single_cell_text(
    cell,
    value: str,
    font_name: str,
    font_size: float,
    color: str,
    bold: bool,
    italic: bool,
    underline: bool,
    h_align: HorizontalAlignment,
    v_align: VerticalAlignment,
) -> None:
    """填充文本到单个单元格（内部辅助函数）

    Args:
        cell: 单元格对象
        value: 要填充的文本
        font_name: 字体名称
        font_size: 字体大小
        color: 颜色
        bold: 是否粗体
        italic: 是否斜体
        underline: 是否下划线
        h_align: 水平对齐方式
        v_align: 垂直对齐方式
    """
    # 清空单元格内容
    cell.Paragraphs.Clear()

    # 添加段落并设置文本
    paragraph = cell.AddParagraph()
    run = paragraph.AppendText(value)

    # 应用样式
    apply_font_style(run, font_name, font_size, color, bold, italic, underline)

    # 应用对齐方式
    if h_align:
        apply_paragraph_alignment(paragraph, h_align)
    if v_align:
        apply_cell_alignment(cell, v_align)


def _fill_single_cell_image(
    cell,
    image_path: str,
    h_align: HorizontalAlignment,
    v_align: VerticalAlignment,
    width: float,
    height: float,
    maintain_ratio: bool,
    original_width_px: int = None,
    original_height_px: int = None,
) -> None:
    """填充图片到单个单元格（内部辅助函数）

    Args:
        cell: 单元格对象
        image_path: 图片文件路径
        h_align: 水平对齐方式
        v_align: 垂直对齐方式
        width: 宽度（磅）
        height: 高度（磅）
        maintain_ratio: 是否保持宽高比
        original_width_px: 原始宽度（像素）
        original_height_px: 原始高度（像素）
    """
    # 清空单元格
    cell.Paragraphs.Clear()

    # 添加段落
    paragraph = cell.AddParagraph()

    # 加载图片
    picture = paragraph.AppendPicture(image_path)

    # 设置图片为内联样式
    from spire.doc import TextWrappingStyle

    picture.TextWrappingStyle = TextWrappingStyle.Inline

    # 应用对齐方式
    if h_align:
        apply_paragraph_alignment(paragraph, h_align)
    if v_align:
        apply_cell_alignment(cell, v_align)

    # 调整图片大小
    if width is not None or height is not None:
        # 优先使用 PIL 获取的尺寸，否则使用 Spire.Doc 的尺寸
        if original_width_px and original_height_px:
            # 像素转换为磅（96 DPI: 1 磅 = 96/72 像素）
            px_to_points = 96.0 / 72.0
            original_width = original_width_px / px_to_points
            original_height = original_height_px / px_to_points
        else:
            original_width = picture.Width
            original_height = picture.Height

        if width is None and height is not None:
            # 只指定高度
            if maintain_ratio and original_width and original_height:
                width = original_width * (height / original_height)
            elif original_width:
                width = original_width
        elif height is None and width is not None:
            # 只指定宽度
            if maintain_ratio and original_width and original_height:
                height = original_height * (width / original_width)
            elif original_height:
                height = original_height

        if width is not None:
            picture.Width = width
        if height is not None:
            picture.Height = height


def _fill_single_cell_date(
    cell,
    numbers: list,
    separators: list,
    font_name: str,
    font_size: float,
    h_align: HorizontalAlignment,
    v_align: VerticalAlignment,
) -> None:
    """填充日期到单个单元格（内部辅助函数）

    Args:
        cell: 单元格对象
        numbers: 数字部分列表
        separators: 分隔符部分列表
        font_name: 数字字体（年月日使用宋体）
        font_size: 字体大小
        h_align: 水平对齐方式
        v_align: 垂直对齐方式
    """
    # 清空单元格
    cell.Paragraphs.Clear()

    # 添加段落
    paragraph = cell.AddParagraph()

    # 依次添加数字和年月日
    for num, sep in zip(numbers, separators):
        # 添加数字（使用指定字体）
        run_num = paragraph.AppendText(num)
        apply_font_style(run_num, font_name, font_size, DEFAULT_COLOR)

        # 添加年月日（使用宋体）
        run_sep = paragraph.AppendText(sep)
        apply_font_style(run_sep, "宋体", font_size, DEFAULT_COLOR)

    # 应用对齐方式
    if h_align:
        apply_paragraph_alignment(paragraph, h_align)
    if v_align:
        apply_cell_alignment(cell, v_align)


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
    underline: bool = False,
    h_align: HorizontalAlignment = None,
    v_align: VerticalAlignment = None,
    match_mode: MatchMode = MatchMode.ALL,
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
        h_align: 水平对齐方式
            - "left": 左对齐
            - "center": 居中对齐
            - "right": 右对齐
            - "justify": 两端对齐
        v_align: 垂直对齐方式
            - "top": 顶部对齐
            - "center": 居中对齐
            - "bottom": 底部对齐
        match_mode: 匹配模式（仅在 match_right/match_down 模式下有效）
            - "all": 填充所有匹配位置（默认）
            - "first": 仅填充第一个匹配位置

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
        >>> fill_text(doc, "标题", "内容", mode="match_right", font_name="黑体", font_size=16, bold=True)

        >>> # 通配符：所有表格的第2行第3列
        >>> fill_text(doc, (1, 0, 2, 3), "统一内容")

        >>> # 通配符：所有节的所有表格的第1行第1列
        >>> fill_text(doc, (0, 0, 1, 1), "标题", h_align="center")

        >>> # 匹配模式：仅填充第一个
        >>> fill_text(doc, "标签：", "值", mode="match_right", match_mode="first")
    """
    try:
        # 确定目标单元格位置
        if mode == FillMode.POSITION:
            if isinstance(position, str):
                raise PositionError("position 模式需要位置元组，不是字符串")

            # 检查是否包含通配符
            if _has_wildcard(position):
                # 使用 get_cells 获取所有匹配的单元格
                cells_list = get_cells(doc, *position)
                if not cells_list:
                    raise PositionError(f"通配符位置 {position} 未匹配到任何单元格")

                # 批量填充
                for _, _, _, _, cell in cells_list:
                    _fill_single_cell_text(
                        cell,
                        value,
                        font_name,
                        font_size,
                        color,
                        bold,
                        italic,
                        underline,
                        h_align,
                        v_align,
                    )
                return
            else:
                # 单个单元格填充
                target_pos = position

        elif mode == FillMode.MATCH_RIGHT:
            if not isinstance(position, str):
                raise PositionError("match_right 模式需要查找文本字符串")
            positions = find_text(doc, position)
            if not positions:
                raise PositionError(f"未找到文本: {position}")

            # 根据 match_mode 决定填充所有还是仅第一个
            target_positions = (
                positions if match_mode == MatchMode.ALL else [positions[0]]
            )

            # 批量填充所有匹配位置
            for pos in target_positions:
                target_pos = (pos[0], pos[1], pos[2], pos[3] + 1)
                cell = get_cell(doc, *target_pos)
                _fill_single_cell_text(
                    cell,
                    value,
                    font_name,
                    font_size,
                    color,
                    bold,
                    italic,
                    underline,
                    h_align,
                    v_align,
                )
            return

        elif mode == FillMode.MATCH_DOWN:
            if not isinstance(position, str):
                raise PositionError("match_down 模式需要查找文本字符串")
            positions = find_text(doc, position)
            if not positions:
                raise PositionError(f"未找到文本: {position}")

            # 根据 match_mode 决定填充所有还是仅第一个
            target_positions = (
                positions if match_mode == MatchMode.ALL else [positions[0]]
            )

            # 批量填充所有匹配位置
            for pos in target_positions:
                target_pos = (pos[0], pos[1], pos[2] + 1, pos[3])
                cell = get_cell(doc, *target_pos)
                _fill_single_cell_text(
                    cell,
                    value,
                    font_name,
                    font_size,
                    color,
                    bold,
                    italic,
                    underline,
                    h_align,
                    v_align,
                )
            return
        else:
            raise FillError(f"不支持的填充模式: {mode}")

        # 单个单元格填充（position 模式且无通配符）
        cell = get_cell(doc, *target_pos)
        _fill_single_cell_text(
            cell,
            value,
            font_name,
            font_size,
            color,
            bold,
            italic,
            underline,
            h_align,
            v_align,
        )

    except (PositionError, FillError):
        raise
    except Exception as e:
        raise FillError(f"填充文本失败: {e}")


def fill_image(
    doc: Document,
    position: Union[Position, str],
    source: Union[str, bytes, Path],
    mode: str = FillMode.POSITION,
    h_align: HorizontalAlignment = None,
    v_align: VerticalAlignment = None,
    width: float = None,
    height: float = None,
    maintain_ratio: bool = True,
    match_mode: MatchMode = MatchMode.ALL,
) -> None:
    """填充图片到文档

    Args:
        doc: Document 对象
        position: 位置元组或查找文本
        source: 图片文件路径（str/Path）或字节数据（bytes）
        mode: 填充模式
            - "position": 直接定位
            - "match_right": 查找文本，填充到右侧
            - "match_down": 查找文本，填充到下方
        h_align: 水平对齐方式
            - "left": 左对齐
            - "center": 居中对齐
            - "right": 右对齐
            - "justify": 两端对齐
        v_align: 垂直对齐方式
            - "top": 顶部对齐
            - "center": 居中对齐
            - "bottom": 底部对齐
        width: 宽度（磅）
        height: 高度（磅）
        maintain_ratio: 是否保持宽高比
        match_mode: 匹配模式（仅在 match_right/match_down 模式下有效）
            - "all": 填充所有匹配位置（默认）
            - "first": 仅填充第一个匹配位置

    Raises:
        FillError: 图片文件不存在或格式不支持
        ValueError: 不支持的源类型
        PositionError: 位置无效

    Examples:
        >>> # 从文件路径填充
        >>> fill_image(doc, (1, 1, 2, 2), "logo.png")

        >>> # 从字节数据填充
        >>> with open("logo.png", "rb") as f:
        ...     data = f.read()
        >>> fill_image(doc, (1, 1, 2, 2), data)

        >>> # 查找文本并填充，指定尺寸
        >>> fill_image(doc, "印章：", "seal.png", mode="match_right", width=100, height=100)

        >>> # 通配符：所有表格的同一位置
        >>> fill_image(doc, (1, 0, 2, 2), "logo.png")

        >>> # 通配符：所有节的所有表格
        >>> fill_image(doc, (0, 0, 1, 1), "header.png")

        >>> # 匹配模式：仅填充第一个
        >>> fill_image(doc, "照片：", "photo.jpg", mode="match_right", match_mode="first")
    """
    import tempfile
    import os

    # 用于存储图片原始尺寸
    original_width_px = None
    original_height_px = None
    temp_file_path = None

    # 处理不同类型的输入（参考 load_docx 的实现模式）
    if isinstance(source, (str, Path)):
        # 文件路径
        file_path = Path(source)

        # 检查文件是否存在
        if not file_path.exists():
            raise FillError(f"图片文件不存在: {source}")

        image_path = str(file_path)

        # 使用 PIL 获取原始尺寸（可选）
        try:
            from PIL import Image as PILImage

            pil_image = PILImage.open(str(file_path))
            original_width_px, original_height_px = pil_image.size
        except ImportError:
            pass

    elif isinstance(source, bytes):
        # 字节数据 - 创建临时文件

        # 使用 PIL 获取原始尺寸（可选）
        try:
            from PIL import Image as PILImage
            from io import BytesIO

            pil_image = PILImage.open(BytesIO(source))
            original_width_px, original_height_px = pil_image.size
        except ImportError:
            pass

        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            tmp.write(source)
            temp_file_path = tmp.name

        image_path = temp_file_path

    else:
        raise ValueError(f"不支持的源类型: {type(source)}")

    try:
        # 确定目标单元格位置
        if mode == FillMode.POSITION:
            if isinstance(position, str):
                raise PositionError("position 模式需要位置元组")

            # 检查是否包含通配符
            if _has_wildcard(position):
                # 使用 get_cells 获取所有匹配的单元格
                cells_list = get_cells(doc, *position)
                if not cells_list:
                    raise PositionError(f"通配符位置 {position} 未匹配到任何单元格")

                # 批量填充
                for _, _, _, _, cell in cells_list:
                    _fill_single_cell_image(
                        cell,
                        image_path,
                        h_align,
                        v_align,
                        width,
                        height,
                        maintain_ratio,
                        original_width_px,
                        original_height_px,
                    )
                return
            else:
                # 单个单元格填充
                target_pos = position

        elif mode == FillMode.MATCH_RIGHT:
            if not isinstance(position, str):
                raise PositionError("match_right 模式需要查找文本字符串")
            positions = find_text(doc, position)
            if not positions:
                raise PositionError(f"未找到文本: {position}")

            # 根据 match_mode 决定填充所有还是仅第一个
            target_positions = (
                positions if match_mode == MatchMode.ALL else [positions[0]]
            )

            # 批量填充所有匹配位置
            for pos in target_positions:
                target_pos = (pos[0], pos[1], pos[2], pos[3] + 1)
                cell = get_cell(doc, *target_pos)
                _fill_single_cell_image(
                    cell,
                    image_path,
                    h_align,
                    v_align,
                    width,
                    height,
                    maintain_ratio,
                    original_width_px,
                    original_height_px,
                )
            return

        elif mode == FillMode.MATCH_DOWN:
            if not isinstance(position, str):
                raise PositionError("match_down 模式需要查找文本字符串")
            positions = find_text(doc, position)
            if not positions:
                raise PositionError(f"未找到文本: {position}")

            # 根据 match_mode 决定填充所有还是仅第一个
            target_positions = (
                positions if match_mode == MatchMode.ALL else [positions[0]]
            )

            # 批量填充所有匹配位置
            for pos in target_positions:
                target_pos = (pos[0], pos[1], pos[2] + 1, pos[3])
                cell = get_cell(doc, *target_pos)
                _fill_single_cell_image(
                    cell,
                    image_path,
                    h_align,
                    v_align,
                    width,
                    height,
                    maintain_ratio,
                    original_width_px,
                    original_height_px,
                )
            return
        else:
            raise FillError(f"不支持的填充模式: {mode}")

        # 单个单元格填充（position 模式且无通配符）
        cell = get_cell(doc, *target_pos)
        _fill_single_cell_image(
            cell,
            image_path,
            h_align,
            v_align,
            width,
            height,
            maintain_ratio,
            original_width_px,
            original_height_px,
        )

    except (PositionError, FillError, ValueError):
        raise
    except Exception as e:
        raise FillError(f"填充图片失败: {e}")

    finally:
        # 清理临时文件
        if temp_file_path and os.path.exists(temp_file_path):
            try:
                os.unlink(temp_file_path)
            except Exception:
                pass


def fill_date(
    doc: Document,
    position: Union[Position, str],
    date_str: str,
    font_name: str = DEFAULT_FONT,
    font_size: float = DEFAULT_FONT_SIZE,
    h_align: HorizontalAlignment = None,
    v_align: VerticalAlignment = None,
    match_mode: MatchMode = MatchMode.ALL,
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
        h_align: 水平对齐方式
            - "left": 左对齐
            - "center": 居中对齐
            - "right": 右对齐
            - "justify": 两端对齐
        v_align: 垂直对齐方式
            - "top": 顶部对齐
            - "center": 居中对齐
            - "bottom": 底部对齐
        match_mode: 匹配模式（仅在 position 为查找文本时有效）
            - "all": 填充所有匹配位置（默认）
            - "first": 仅填充第一个匹配位置

    Raises:
        PositionError: 位置无效
        ValidationError: 日期格式无效或日期不存在
        FillError: 填充失败

    Examples:
        >>> fill_date(doc, (1, 1, 4, 2), "2024年1月15日")
        >>> fill_date(doc, "日期：", "2024年1月15日")
        >>> fill_date(doc, (1, 1, 4, 2), "2024年1月15日", h_align="center")

        >>> # 通配符：所有表格的同一位置
        >>> fill_date(doc, (1, 0, 4, 2), "2024年1月15日")

        >>> # 通配符：所有节的所有表格
        >>> fill_date(doc, (0, 0, 4, 2), "2024年1月15日")

        >>> # 匹配模式：仅填充第一个
        >>> fill_date(doc, "日期：", "2024年1月15日", match_mode="first")
    """
    try:

        from .utils import parse_date_string, validate_date_string

        # 验证日期格式和有效性
        validate_date_string(date_str)

        numbers, separators = parse_date_string(date_str)

        if not numbers or not separators:
            raise FillError(
                f"无效的日期字符串: '{date_str}'，"
                f"期望格式如 '2024年1月15日' 或 '2024年01月15日'"
            )

        # 确定目标单元格位置
        if isinstance(position, str):
            # 字符串模式：查找文本并填充到右侧
            positions = find_text(doc, position)
            if not positions:
                raise PositionError(f"未找到文本: {position}")

            # 根据 match_mode 决定填充所有还是仅第一个
            target_positions = (
                positions if match_mode == MatchMode.ALL else [positions[0]]
            )

            # 批量填充所有匹配位置
            for pos in target_positions:
                target_pos = (pos[0], pos[1], pos[2], pos[3] + 1)
                cell = get_cell(doc, *target_pos)
                _fill_single_cell_date(
                    cell, numbers, separators, font_name, font_size, h_align, v_align
                )
            return
        else:
            # 位置元组模式
            # 检查是否包含通配符
            if _has_wildcard(position):
                # 使用 get_cells 获取所有匹配的单元格
                cells_list = get_cells(doc, *position)
                if not cells_list:
                    raise PositionError(f"通配符位置 {position} 未匹配到任何单元格")

                # 批量填充
                for _, _, _, _, cell in cells_list:
                    _fill_single_cell_date(
                        cell,
                        numbers,
                        separators,
                        font_name,
                        font_size,
                        h_align,
                        v_align,
                    )
                return
            else:
                # 单个单元格填充
                target_pos = position

        # 单个单元格填充（无通配符）
        cell = get_cell(doc, *target_pos)
        _fill_single_cell_date(
            cell, numbers, separators, font_name, font_size, h_align, v_align
        )

    except (PositionError, FillError, ValidationError):
        raise
    except Exception as e:
        raise FillError(f"填充日期失败: {e}")


def fill_grid(doc: Document, data: List[List[str]], position: Position) -> None:
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
                    cell = get_cell(doc, section_idx, table_idx, target_row, target_col)

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


def clear_cell(doc: Document, section: int, table: int, row: int, col: int) -> None:
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


def _find_variables(text: str, prefix: str, suffix: str) -> List[Tuple[str, str, str]]:
    """查找文本中的所有变量

    Args:
        text: 要查找的文本
        prefix: 变量前缀
        suffix: 变量后缀

    Returns:
        List[Tuple[完整变量, 变量名, 默认值]]
    """
    pattern = (
        re.escape(prefix)
        + r"([a-zA-Z_][a-zA-Z0-9_]*)(?:\|([^}]*))?"
        + re.escape(suffix)
    )
    matches = []
    for match in re.finditer(pattern, text):
        full_var = match.group(0)
        var_name = match.group(1)
        default_val = match.group(2) if match.group(2) is not None else ""
        matches.append((full_var, var_name, default_val))
    return matches


def fill_template(
    doc: Document,
    data: Dict[str, Any],
    *,
    missing_var_action: str = DEFAULT_MISSING_VAR_ACTION,
    placeholder_prefix: str = DEFAULT_VAR_PREFIX,
    placeholder_suffix: str = DEFAULT_VAR_SUFFIX,
) -> Dict[str, Any]:
    """批量替换模板变量

    Args:
        doc: Document 对象
        data: 变量数据字典
        missing_var_action: 缺失变量处理方式 ("error" | "ignore" | "empty")
        placeholder_prefix: 变量前缀
        placeholder_suffix: 变量后缀

    Returns:
        Dict: {"total": int, "replaced": int, "missing": list}

    Raises:
        VariableNotFoundError: 变量未找到时
        FillError: 填充失败时

    Examples:
        >>> fill_template(doc, {"name": "张三", "age": "25"})
        >>> fill_template(doc, data, missing_var_action="ignore")
    """
    try:
        stats = {"total": 0, "replaced": 0, "missing": []}
        replacements = {}

        # 遍历文档收集变量
        for section_idx in range(doc.Sections.Count):
            section = doc.Sections.get_Item(section_idx)

            # 段落
            for para_idx in range(section.Paragraphs.Count):
                paragraph = section.Paragraphs.get_Item(para_idx)
                matches = _find_variables(
                    paragraph.Text, placeholder_prefix, placeholder_suffix
                )
                stats["total"] += len(matches)
                for full_var, var_name, default_val in matches:
                    if full_var in replacements:
                        continue
                    if var_name in data:
                        replacements[full_var] = str(data[var_name])
                    elif default_val:
                        replacements[full_var] = default_val
                    elif missing_var_action == "error":
                        stats["missing"].append(var_name)
                        raise VariableNotFoundError(var_name, list(data.keys()))
                    elif missing_var_action == "empty":
                        replacements[full_var] = ""

            # 表格
            for table_idx in range(section.Tables.Count):
                table = section.Tables.get_Item(table_idx)
                for row_idx in range(table.Rows.Count):
                    row = table.Rows.get_Item(row_idx)
                    for cell_idx in range(row.Cells.Count):
                        cell = row.Cells.get_Item(cell_idx)
                        for para_idx in range(cell.Paragraphs.Count):
                            paragraph = cell.Paragraphs.get_Item(para_idx)
                            matches = _find_variables(
                                paragraph.Text, placeholder_prefix, placeholder_suffix
                            )
                            stats["total"] += len(matches)
                            for full_var, var_name, default_val in matches:
                                if full_var in replacements:
                                    continue
                                if var_name in data:
                                    replacements[full_var] = str(data[var_name])
                                elif default_val:
                                    replacements[full_var] = default_val
                                elif missing_var_action == "error":
                                    stats["missing"].append(var_name)
                                    raise VariableNotFoundError(
                                        var_name, list(data.keys())
                                    )
                                elif missing_var_action == "empty":
                                    replacements[full_var] = ""

        # 执行替换
        for full_var, value in replacements.items():
            replace_all(doc, full_var, value)
            stats["replaced"] += 1

        return stats

    except VariableNotFoundError:
        raise
    except Exception as e:
        raise FillError(f"填充模板失败: {e}")


def extract_template_vars(
    doc: Document,
    *,
    placeholder_prefix: str = DEFAULT_VAR_PREFIX,
    placeholder_suffix: str = DEFAULT_VAR_SUFFIX,
    unique: bool = True,
) -> List[str]:
    """提取模板中的所有变量

    Args:
        doc: Document 对象
        placeholder_prefix: 变量前缀
        placeholder_suffix: 变量后缀
        unique: 是否去重

    Returns:
        List[str]: 变量名列表

    Examples:
        >>> vars = extract_template_vars(doc)
        >>> vars = extract_template_vars(doc, unique=False)
    """
    try:
        all_vars = []

        for section_idx in range(doc.Sections.Count):
            section = doc.Sections.get_Item(section_idx)

            # 段落
            for para_idx in range(section.Paragraphs.Count):
                paragraph = section.Paragraphs.get_Item(para_idx)
                for _, var_name, _ in _find_variables(
                    paragraph.Text, placeholder_prefix, placeholder_suffix
                ):
                    all_vars.append(var_name)

            # 表格
            for table_idx in range(section.Tables.Count):
                table = section.Tables.get_Item(table_idx)
                for row_idx in range(table.Rows.Count):
                    row = table.Rows.get_Item(row_idx)
                    for cell_idx in range(row.Cells.Count):
                        cell = row.Cells.get_Item(cell_idx)
                        for para_idx in range(cell.Paragraphs.Count):
                            paragraph = cell.Paragraphs.get_Item(para_idx)
                            for _, var_name, _ in _find_variables(
                                paragraph.Text, placeholder_prefix, placeholder_suffix
                            ):
                                all_vars.append(var_name)

        if unique:
            seen = set()
            unique_vars = []
            for var in all_vars:
                if var not in seen:
                    seen.add(var)
                    unique_vars.append(var)
            return unique_vars

        return all_vars

    except Exception as e:
        raise FillError(f"提取模板变量失败: {e}")


def validate_template_data(
    doc: Document,
    data: Dict[str, Any],
    *,
    placeholder_prefix: str = DEFAULT_VAR_PREFIX,
    placeholder_suffix: str = DEFAULT_VAR_SUFFIX,
) -> Dict[str, Any]:
    """验证模板数据是否完整

    Args:
        doc: Document 对象
        data: 变量数据字典
        placeholder_prefix: 变量前缀
        placeholder_suffix: 变量后缀

    Returns:
        Dict: {"is_valid": bool, "missing_vars": list, "required_vars": list, "extra_vars": list}

    Examples:
        >>> result = validate_template_data(doc, {"name": "张三"})
        >>> result["is_valid"]
    """
    try:
        required_vars = set(
            extract_template_vars(
                doc,
                placeholder_prefix=placeholder_prefix,
                placeholder_suffix=placeholder_suffix,
            )
        )
        provided_vars = set(data.keys()) - {"__styles__"}

        return {
            "is_valid": required_vars.issubset(provided_vars),
            "required_vars": list(required_vars),
            "missing_vars": list(required_vars - provided_vars),
            "extra_vars": list(provided_vars - required_vars),
        }

    except Exception as e:
        raise FillError(f"验证模板数据失败: {e}")
