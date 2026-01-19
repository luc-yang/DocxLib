"""
DocxLib 读取模块

提供文档内容读取功能，与 fill.py 模块形成对称的 API 设计。
"""

from typing import Any, Dict, List, Tuple, Union

from spire.doc import *
from spire.doc.common import *

from .config import Options
from .constants import DEFAULT_VAR_PREFIX, DEFAULT_VAR_SUFFIX, Position
from .errors import DocumentError, FillError, PositionError
from .fill import _find_variables
from .style import get_cell_style
from .table import (
    find_text,
    get_cell,
    get_cell_text,
    get_cells,
    get_section_count,
    get_section_table_count,
    get_table_dimensions,
    get_table_text,
    iterate_cells,
)


def read_text(
    doc: Document,
    position: Union[Position, str],
    *,
    default: str = "",
    options: Options = None,
) -> str:
    """读取单元格文本

    Args:
        doc: Document 对象
        position: 位置元组 (section, table, row, col) 或查找文本
        default: 未找到时的默认值
        options: 填充模式配置（FillOptions）

    Returns:
        单元格文本内容，未找到时返回默认值

    Examples:
        >>> # 直接位置读取
        >>> text = read_text(doc, (1, 1, 2, 2))

        >>> # 查找右侧文本
        >>> name = read_text(doc, "姓名：", default="未知")

        >>> # 查找下方文本
        >>> value = read_text(doc, "项目", default="N/A")

        >>> # 使用匹配模式
        >>> text = read_text(
        ...     doc, "标签：",
        ...     options=FillOptions.match_right()
        ... )
    """
    if options is None:
        options = Options()

    try:
        mode = options.mode

        # 直接位置模式
        if mode == "position":
            if isinstance(position, str):
                raise PositionError("position 模式需要位置元组，不是字符串")
            try:
                return get_cell_text(doc, *position)
            except PositionError:
                return default

        # 查找右侧模式
        elif mode == "match_right":
            if not isinstance(position, str):
                raise PositionError("match_right 模式需要查找文本字符串")
            positions = find_text(doc, position)
            if not positions:
                return default
            # 取第一个匹配位置
            pos = positions[0]
            target_pos = (pos[0], pos[1], pos[2], pos[3] + 1)
            try:
                return get_cell_text(doc, *target_pos)
            except PositionError:
                return default

        # 查找下方模式
        elif mode == "match_down":
            if not isinstance(position, str):
                raise PositionError("match_down 模式需要查找文本字符串")
            positions = find_text(doc, position)
            if not positions:
                return default
            # 取第一个匹配位置
            pos = positions[0]
            target_pos = (pos[0], pos[1], pos[2] + 1, pos[3])
            try:
                return get_cell_text(doc, *target_pos)
            except PositionError:
                return default

        else:
            raise FillError(f"不支持的读取模式: {mode}")

    except (PositionError, FillError):
        raise
    except Exception as e:
        raise FillError(f"读取文本失败: {e}")


def read_grid(
    doc: Document,
    position: Position,
    *,
    include_empty: bool = True,
) -> List[List[str]]:
    """读取网格数据

    从指定位置开始读取表格数据，返回二维数组。

    Args:
        doc: Document 对象
        position: 起始位置 (section, table, row, col)
        include_empty: 是否包含空单元格

    Returns:
        二维文本数组

    Raises:
        PositionError: 表格不存在或位置无效

    Examples:
        >>> # 读取整个表格
        >>> grid = read_grid(doc, (1, 1, 1, 1))

        >>> # 读取特定区域（从第2行第2列开始）
        >>> partial = read_grid(doc, (1, 1, 2, 2))
    """
    try:
        section_idx, table_idx, start_row, start_col = position

        # 获取完整的表格数据
        full_table = get_table_text(doc, section_idx, table_idx)

        # 提取指定区域（从 start_row-1, start_col-1 开始）
        result = []
        for row_idx in range(start_row - 1, len(full_table)):
            row_data = full_table[row_idx]
            if start_col - 1 < len(row_data):
                result.append(row_data[start_col - 1 :])

        return result

    except PositionError:
        raise
    except Exception as e:
        raise FillError(f"读取网格数据失败: {e}")


def read_images(
    doc: Document,
    *,
    section: int = 0,
    table: int = 0,
    include_data: bool = False,
) -> List[Dict[str, Any]]:
    """提取文档中的所有图片

    Args:
        doc: Document 对象
        section: 节索引（0表示所有）
        table: 表格索引（0表示所有）
        include_data: 是否包含图片字节数据

    Returns:
        图片信息列表：
        [
            {
                "position": (1, 1, 2, 2),  # (section, table, row, col)
                "width": 100.5,              # 磅
                "height": 80.3,              # 磅
                "format": "png",             # 图片格式
                "index": 0,                  # 在单元格中的索引
                "data": bytes(...)           # 图片字节数据（仅当 include_data=True）
            },
            ...
        ]

    Raises:
        DocumentError: 提取失败时

    Examples:
        >>> # 提取所有图片（仅元数据）
        >>> images = read_images(doc)

        >>> # 提取图片并包含字节数据
        >>> images = read_images(doc, include_data=True)
        >>> for img in images:
        ...     with open(f"image_{img['index']}.{img['format']}", "wb") as f:
        ...         f.write(img["data"])

        >>> # 提取第1节的图片
        >>> images = read_images(doc, section=1)

        >>> # 提取特定表格的图片
        >>> images = read_images(doc, section=1, table=1)
    """
    try:
        from io import BytesIO

        images = []

        # 使用 get_cells 获取目标单元格
        cells_list = get_cells(doc, section=section, table=table)

        for sec_idx, tbl_idx, row_idx, col_idx, cell in cells_list:
            # 遍历单元格的段落
            for para_idx in range(cell.Paragraphs.Count):
                paragraph = cell.Paragraphs.get_Item(para_idx)

                # 遍历段落中的子对象
                for obj_idx in range(paragraph.ChildObjects.Count):
                    child = paragraph.ChildObjects.get_Item(obj_idx)

                    # 检查是否是图片
                    if isinstance(child, DocPicture):
                        # 获取图片格式
                        image_type = child.ImageType
                        format_map = {
                            "Bitmap": "bmp",
                            "JPEG": "jpg",
                            "GIF": "gif",
                            "PNG": "png",
                            "EMF": "emf",
                            "WMF": "wmf",
                        }
                        format_str = format_map.get(str(image_type), "unknown")

                        image_info = {
                            "position": (
                                sec_idx + 1,
                                tbl_idx + 1,
                                row_idx + 1,
                                col_idx + 1,
                            ),
                            "width": child.Width,
                            "height": child.Height,
                            "format": format_str,
                            "index": obj_idx,
                        }

                        # 如果需要包含图片数据
                        if include_data:
                            try:
                                # 使用 BytesIO 捕获图片数据
                                stream = BytesIO()
                                child.Image.Save(stream, child.ImageType)
                                image_info["data"] = stream.getvalue()
                            except Exception:
                                # 如果提取数据失败，添加空字节
                                image_info["data"] = b""

                        images.append(image_info)

        return images

    except Exception as e:
        raise DocumentError(f"提取图片失败: {e}")


def read_document_structure(
    doc: Document,
    *,
    include_paragraphs: bool = False,
    include_tables: bool = True,
) -> Dict[str, Any]:
    """读取文档结构

    分析文档的节、表格、段落结构，返回结构化信息。

    Args:
        doc: Document 对象
        include_paragraphs: 是否包含段落信息
        include_tables: 是否包含表格信息

    Returns:
        文档结构字典：
        {
            "section_count": 1,
            "sections": [
                {
                    "index": 1,
                    "table_count": 3,
                    "paragraph_count": 10,
                    "tables": [...]
                },
                ...
            ]
        }

    Raises:
        DocumentError: 分析失败时

    Examples:
        >>> # 基本结构
        >>> structure = read_document_structure(doc)

        >>> # 包含段落信息
        >>> structure = read_document_structure(doc, include_paragraphs=True)

        >>> # 仅表格信息
        >>> structure = read_document_structure(
        ...     doc, include_tables=True, include_paragraphs=False
        ... )
    """
    try:
        section_count = get_section_count(doc)
        sections = []

        for sec_idx in range(1, section_count + 1):
            section_info = {"index": sec_idx}
            section_obj = doc.Sections.get_Item(sec_idx - 1)

            # 统计段落数量
            if include_paragraphs:
                para_count = section_obj.Paragraphs.Count
                section_info["paragraph_count"] = para_count

            # 统计表格信息
            if include_tables:
                table_count = get_section_table_count(doc, sec_idx)
                section_info["table_count"] = table_count

                # 获取每个表格的尺寸
                tables = []
                for tbl_idx in range(1, table_count + 1):
                    try:
                        dimensions = get_table_dimensions(doc, sec_idx, tbl_idx)
                        tables.append(
                            {"index": tbl_idx, "rows": dimensions[0], "columns": dimensions[1]}
                        )
                    except PositionError:
                        # 跳过无法访问的表格
                        pass

                section_info["tables"] = tables

            sections.append(section_info)

        return {"section_count": section_count, "sections": sections}

    except Exception as e:
        raise DocumentError(f"读取文档结构失败: {e}")


def read_all_text(
    doc: Document,
    *,
    include_tables: bool = True,
    include_paragraphs: bool = True,
    separator: str = "\n",
) -> str:
    """提取文档所有文本

    Args:
        doc: Document 对象
        include_tables: 是否包含表格文本
        include_paragraphs: 是否包含段落文本
        separator: 文本分隔符

    Returns:
        完整的文本字符串

    Raises:
        DocumentError: 提取失败时

    Examples:
        >>> # 提取所有文本
        >>> text = read_all_text(doc)

        >>> # 仅表格文本
        >>> text = read_all_text(doc, include_paragraphs=False)

        >>> # 自定义分隔符
        >>> text = read_all_text(doc, separator="\\n\\n")
    """
    try:
        text_parts = []

        # 遍历所有节
        for sec_idx in range(doc.Sections.Count):
            section = doc.Sections.get_Item(sec_idx)

            # 提取段落文本
            if include_paragraphs:
                for para_idx in range(section.Paragraphs.Count):
                    paragraph = section.Paragraphs.get_Item(para_idx)
                    text = paragraph.Text.strip()
                    if text:  # 只添加非空段落
                        text_parts.append(text)

            # 提取表格文本
            if include_tables:
                for table_idx in range(section.Tables.Count):
                    try:
                        table_data = get_table_text(doc, sec_idx + 1, table_idx + 1)
                        # 将表格格式化为文本
                        for row in table_data:
                            row_text = " | ".join(row)
                            text_parts.append(row_text)
                    except PositionError:
                        # 跳过无法访问的表格
                        pass

        return separator.join(text_parts)

    except Exception as e:
        raise DocumentError(f"提取所有文本失败: {e}")


def read_table(
    doc: Document,
    section: int,
    table: int,
    *,
    include_style: bool = False,
    include_merged: bool = False,
) -> Union[List[List[str]], Dict[str, Any]]:
    """读取表格

    Args:
        doc: Document 对象
        section: 节索引
        table: 表格索引
        include_style: 是否包含样式信息
        include_merged: 是否标记合并单元格

    Returns:
        二维数组或完整信息字典

    Raises:
        PositionError: 表格不存在

    Examples:
        >>> # 基础读取
        >>> table_data = read_table(doc, 1, 1)

        >>> # 包含样式
        >>> table_info = read_table(doc, 1, 1, include_style=True)

        >>> # 完整信息
        >>> table_full = read_table(
        ...     doc, 1, 1,
        ...     include_style=True,
        ...     include_merged=True
        ... )
    """
    try:
        # 获取基础数据
        data = get_table_text(doc, section, table)
        dimensions = get_table_dimensions(doc, section, table)

        # 基础模式：仅返回数据
        if not include_style and not include_merged:
            return data

        # 完整模式：返回详细信息
        result = {"data": data, "dimensions": {"rows": dimensions[0], "columns": dimensions[1]}}

        # 包含样式信息
        if include_style:
            styles = []
            for row_idx, row_data in enumerate(data):
                row_styles = []
                for col_idx in range(len(row_data)):
                    try:
                        cell = get_cell(doc, section, table, row_idx + 1, col_idx + 1)
                        style = get_cell_style(cell)
                        row_styles.append(style)
                    except PositionError:
                        row_styles.append({})
                styles.append(row_styles)
            result["styles"] = styles

        # 包含合并单元格信息
        if include_merged:
            merged_cells = []
            try:
                table_obj = doc.Sections.get_Item(section - 1).Tables.get_Item(table - 1)
                for row_idx in range(table_obj.Rows.Count):
                    row_obj = table_obj.Rows.get_Item(row_idx)
                    for col_idx in range(row_obj.Cells.Count):
                        cell = row_obj.Cells.get_Item(col_idx)
                        cell_format = cell.CellFormat

                        # 检查垂直合并
                        v_merge = cell_format.VerticalMerge
                        h_merge = cell_format.HorizontalMerge

                        # 如果有合并
                        if v_merge != 0 or h_merge != 0:
                            # 记录合并信息（从1-based索引开始）
                            merged_cells.append(
                                {
                                    "start": (section, table, row_idx + 1, col_idx + 1),
                                    "vertical_merge": str(v_merge) if v_merge != 0 else None,
                                    "horizontal_merge": str(h_merge)
                                    if h_merge != 0
                                    else None,
                                }
                            )
            except Exception:
                # 如果获取合并信息失败，忽略
                pass

            result["merged_cells"] = merged_cells

        return result

    except PositionError:
        raise
    except Exception as e:
        raise FillError(f"读取表格失败: {e}")


def read_cells(
    doc: Document,
    positions: List[Position],
    *,
    default: str = "",
) -> List[str]:
    """批量读取多个单元格

    Args:
        doc: Document 对象
        positions: 位置列表
        default: 默认值

    Returns:
        文本列表

    Examples:
        >>> # 批量读取
        >>> positions = [(1, 1, 2, 2), (1, 1, 2, 3), (1, 1, 2, 4)]
        >>> values = read_cells(doc, positions, default="N/A")
    """
    result = []

    for position in positions:
        try:
            text = get_cell_text(doc, *position)
            result.append(text)
        except PositionError:
            result.append(default)

    return result


def extract_template_data(
    doc: Document,
    *,
    placeholder_prefix: str = DEFAULT_VAR_PREFIX,
    placeholder_suffix: str = DEFAULT_VAR_SUFFIX,
    unique: bool = True,
) -> Dict[str, Any]:
    """提取模板数据

    增强版 extract_template_vars，提供更详细的变量信息。

    Args:
        doc: Document 对象
        placeholder_prefix: 变量前缀
        placeholder_suffix: 变量后缀
        unique: 是否去重

    Returns:
        {
            "variables": ["name", "age"],
            "variable_details": [
                {
                    "name": "name",
                    "default": "",
                    "positions": [(1, 1, 2, 2), (1, 2, 1, 1)]
                },
                ...
            ],
            "total_count": 5
        }

    Raises:
        FillError: 提取失败时

    Examples:
        >>> # 基础提取
        >>> data = extract_template_data(doc)

        >>> # 自定义占位符
        >>> data = extract_template_data(
        ...     doc,
        ...     placeholder_prefix="{{",
        ...     placeholder_suffix="}}"
        ... )
    """
    try:
        variable_map = {}  # {var_name: {"default": str, "positions": list}}
        total_count = 0

        # 遍历所有单元格
        for sec_idx, tbl_idx, row_idx, col_idx, cell in iterate_cells(doc):
            # 遍历单元格的段落
            for para_idx in range(cell.Paragraphs.Count):
                paragraph = cell.Paragraphs.get_Item(para_idx)
                matches = _find_variables(
                    paragraph.Text, placeholder_prefix, placeholder_suffix
                )

                for full_var, var_name, default_val in matches:
                    total_count += 1

                    if var_name not in variable_map:
                        variable_map[var_name] = {"default": default_val, "positions": []}

                    # 添加位置信息
                    position = (sec_idx + 1, tbl_idx + 1, row_idx + 1, col_idx + 1)
                    variable_map[var_name]["positions"].append(position)

        # 构建变量详情列表
        variable_details = []
        for var_name, info in variable_map.items():
            variable_details.append(
                {"name": var_name, "default": info["default"], "positions": info["positions"]}
            )

        # 构建结果
        result = {
            "variables": list(variable_map.keys()),
            "variable_details": variable_details,
            "total_count": total_count,
        }

        return result

    except Exception as e:
        raise FillError(f"提取模板数据失败: {e}")
