"""
DocxLib 表格操作模块

提供表格遍历、单元格定位、文本查找等功能。
"""

from typing import Generator, List, Tuple, Union

from spire.doc import *
from spire.doc.common import *

from .errors import PositionError
from .constants import Position


def get_cell(doc: Document, section: int, table: int, row: int, col: int):
    """获取指定位置的单元格

    所有索引从 1 开始（不是 0）。

    Args:
        doc: Document 对象
        section: 节索引（从1开始）
        table: 表格索引（从1开始）
        row: 行索引（从1开始）
        col: 列索引（从1开始）

    Returns:
        Cell: 单元格对象

    Raises:
        PositionError: 位置越界

    Examples:
        >>> cell = get_cell(doc, 1, 1, 2, 2)
        >>> text = cell.Range.Text.strip()
        >>> print(text)
        '单元格内容'
    """
    try:
        # 获取节（索引从0开始）
        section_obj = doc.Sections.get_Item(section - 1)

        # 获取表格（索引从0开始）
        table_obj = section_obj.Tables.get_Item(table - 1)

        # 获取行（索引从0开始）
        row_obj = table_obj.Rows.get_Item(row - 1)

        # 获取单元格（索引从0开始）
        cell = row_obj.Cells.get_Item(col - 1)

        return cell

    except Exception as e:
        raise PositionError(
            f"无法获取位置 ({section}, {table}, {row}, {col}) 的单元格: {e}"
        )


def get_cells(
    doc: Document, section: int = 0, table: int = 0, row: int = 0, col: int = 0
) -> List[Tuple]:
    """通配符获取单元格

    0 表示所有。用于批量获取符合条件的单元格。

    Args:
        doc: Document 对象
        section: 节索引（0表示所有）
        table: 表格索引（0表示所有）
        row: 行索引（0表示所有）
        col: 列索引（0表示所有）

    Returns:
        List[Tuple]: [(section, table, row, col, cell), ...]

    Examples:
        >>> # 获取第1节、第1个表格的所有单元格
        >>> cells = get_cells(doc, section=1, table=1)

        >>> # 获取所有节的所有表格的所有单元格
        >>> all_cells = get_cells(doc)

        >>> # 获取第1节所有表格的第2行第2列
        >>> cells = get_cells(doc, section=1, row=2, col=2)
    """
    result = []

    # 遍历节
    section_count = doc.Sections.Count
    section_indices = range(section_count) if section == 0 else [section - 1]

    for sec_idx in section_indices:
        if sec_idx >= section_count:
            continue

        section_obj = doc.Sections.get_Item(sec_idx)

        # 遍历表格
        table_count = section_obj.Tables.Count
        table_indices = range(table_count) if table == 0 else [table - 1]

        for tbl_idx in table_indices:
            if tbl_idx >= table_count:
                continue

            table_obj = section_obj.Tables.get_Item(tbl_idx)

            # 遍历行
            row_count = table_obj.Rows.Count
            row_indices = range(row_count) if row == 0 else [row - 1]

            for r_idx in row_indices:
                if r_idx >= row_count:
                    continue

                row_obj = table_obj.Rows.get_Item(r_idx)

                # 遍历列
                col_count = row_obj.Cells.Count
                col_indices = range(col_count) if col == 0 else [col - 1]

                for c_idx in col_indices:
                    if c_idx >= col_count:
                        continue

                    cell = row_obj.Cells.get_Item(c_idx)
                    # 索引从1开始返回
                    result.append(
                        (sec_idx + 1, tbl_idx + 1, r_idx + 1, c_idx + 1, cell)
                    )

    return result


def find_text(doc: Document, text: str) -> List[Position]:
    """查找文档中包含指定文本的所有单元格位置

    Args:
        doc: Document 对象
        text: 要查找的文本

    Returns:
        List[Position]: 位置列表 [(section, table, row, col), ...]

    Examples:
        >>> positions = find_text(doc, "姓名")
        >>> print(positions)
        [(1, 1, 2, 1)]
    """
    positions = []

    for section_idx, table_idx, row_idx, col_idx, cell in iterate_cells(doc):
        # 获取单元格文本
        cell_text = ""
        for m in range(cell.Paragraphs.Count):
            paragraph = cell.Paragraphs.get_Item(m)
            paragraph_text = paragraph.Text.strip()
            cell_text += paragraph_text
        if text == cell_text:
            positions.append((section_idx, table_idx, row_idx, col_idx))

    return positions


def iterate_cells(doc: Document) -> Generator:
    """遍历文档中所有单元格

    Args:
        doc: Document 对象

    Yields:
        tuple: (section, table, row, col, cell)
        索引从1开始

    Examples:
        >>> for sec, tbl, row, col, cell in iterate_cells(doc):
        ...     text = cell.Range.Text.strip()
        ...     if text:
        ...         print(f"({sec}, {tbl}, {row}, {col}): {text}")
    """
    section_count = doc.Sections.Count

    for sec_idx in range(section_count):
        section_obj = doc.Sections.get_Item(sec_idx)
        table_count = section_obj.Tables.Count

        for tbl_idx in range(table_count):
            table_obj = section_obj.Tables.get_Item(tbl_idx)
            row_count = table_obj.Rows.Count

            for row_idx in range(row_count):
                row_obj = table_obj.Rows.get_Item(row_idx)
                col_count = row_obj.Cells.Count

                for col_idx in range(col_count):
                    cell = row_obj.Cells.get_Item(col_idx)
                    # 索引从1开始返回
                    yield (sec_idx + 1, tbl_idx + 1, row_idx + 1, col_idx + 1, cell)


def get_cell_text(doc: Document, section: int, table: int, row: int, col: int) -> str:
    """获取单元格文本内容

    Args:
        doc: Document 对象
        section: 节索引（从1开始）
        table: 表格索引（从1开始）
        row: 行索引（从1开始）
        col: 列索引（从1开始）

    Returns:
        str: 单元格文本内容（去除首尾空白）

    Examples:
        >>> text = get_cell_text(doc, 1, 1, 2, 2)
        >>> print(text)
        '单元格内容'
    """
    cell = get_cell(doc, section, table, row, col)
    return cell.Range.Text.strip()


def get_table_dimensions(doc: Document, section: int, table: int) -> Tuple[int, int]:
    """获取表格的行数和列数

    Args:
        doc: Document 对象
        section: 节索引（从1开始）
        table: 表格索引（从1开始）

    Returns:
        Tuple[int, int]: (行数, 列数)

    Raises:
        PositionError: 表格不存在

    Examples:
        >>> rows, cols = get_table_dimensions(doc, 1, 1)
        >>> print(f"表格大小: {rows}行 x {cols}列")
        表格大小: 10行 x 5列
    """
    try:
        section_obj = doc.Sections.get_Item(section - 1)
        table_obj = section_obj.Tables.get_Item(table - 1)

        rows = table_obj.Rows.Count
        # 获取第一行的列数（假设所有行列数相同）
        if rows > 0:
            cols = table_obj.Rows.get_Item(0).Cells.Count
        else:
            cols = 0

        return rows, cols

    except Exception as e:
        raise PositionError(f"无法获取表格 ({section}, {table}) 的尺寸: {e}")


def get_section_table_count(doc: Document, section: int) -> int:
    """获取指定节中的表格数量

    Args:
        doc: Document 对象
        section: 节索引（从1开始）

    Returns:
        int: 表格数量

    Examples:
        >>> count = get_section_table_count(doc, 1)
        >>> print(f"第1节有 {count} 个表格")
        第1节有 3 个表格
    """
    section_obj = doc.Sections.get_Item(section - 1)
    return section_obj.Tables.Count


def get_section_count(doc: Document) -> int:
    """获取文档中的节数量

    Args:
        doc: Document 对象

    Returns:
        int: 节数量

    Examples:
        >>> count = get_section_count(doc)
        >>> print(f"文档有 {count} 个节")
        文档有 1 个节
    """
    return doc.Sections.Count
