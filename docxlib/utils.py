"""
DocxLib 工具函数

提供通用的辅助函数，如文件格式验证、数据解析等。
"""

import zipfile
from pathlib import Path
from typing import Union, List, Dict, Any

from .errors import ValidationError


def is_valid_docx(source: Union[str, bytes, Path]) -> bool:
    """验证文件是否为有效的 DOCX 格式

    DOCX 文件实际上是一个 ZIP 压缩包，包含特定的 XML 文件结构。
    我们通过检查是否可以解压以及是否包含必要文件来验证。

    Args:
        source: 文件路径（str/Path）或字节数据（bytes）

    Returns:
        bool: 如果是有效的 DOCX 文件返回 True，否则返回 False

    Examples:
        >>> is_valid_docx("document.docx")
        True
        >>> is_valid_docx("document.txt")
        False
    """
    if isinstance(source, (str, Path)):
        file_path = Path(source)
        if not file_path.exists():
            return False

        # 检查文件扩展名
        if file_path.suffix.lower() not in ['.docx', '.dotx']:
            return False

        # 尝试打开为 ZIP 文件
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_file:
                # 检查是否包含必要的 DOCX 文件
                required_files = ['[Content_Types].xml', 'word/document.xml']
                zip_file_names = zip_file.namelist()
                return any(req in zip_file_names for req in required_files)
        except (zipfile.BadZipFile, zipfile.LargeZipFile):
            return False

    elif isinstance(source, bytes):
        # 检查字节数据是否为有效的 ZIP 格式
        try:
            import io
            with zipfile.ZipFile(io.BytesIO(source), 'r') as zip_file:
                required_files = ['[Content_Types].xml', 'word/document.xml']
                zip_file_names = zip_file.namelist()
                return any(req in zip_file_names for req in required_files)
        except (zipfile.BadZipFile, zipfile.LargeZipFile):
            return False

    return False


def validate_docx(source: Union[str, bytes, Path]) -> None:
    """验证 DOCX 文件格式，无效时抛出异常

    Args:
        source: 文件路径或字节数据

    Raises:
        ValidationError: 文件格式不是有效的 DOCX
    """
    if not is_valid_docx(source):
        file_desc = source if isinstance(source, (str, Path)) else "字节数据"
        raise ValidationError(f"'{file_desc}' 不是有效的 DOCX 文件格式")


def parse_csv(file_path: Union[str, Path]) -> List[List[str]]:
    """解析 CSV 文件

    Args:
        file_path: CSV 文件路径

    Returns:
        List[List[str]]: 二维数组，每个元素代表一个单元格的值

    Raises:
        FileNotFoundError: 文件不存在
        ValidationError: CSV 格式错误

    Examples:
        >>> data = parse_csv("data.csv")
        >>> print(data[0])  # 第一行数据
        ['列1', '列2', '列3']
    """
    import csv

    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")

    try:
        with open(path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            return [list(row) for row in reader]
    except csv.Error as e:
        raise ValidationError(f"CSV 格式错误: {e}")


def parse_json(file_path: Union[str, Path]) -> Dict[str, Any]:
    """解析 JSON 文件

    Args:
        file_path: JSON 文件路径

    Returns:
        Dict[str, Any]: 解析后的 JSON 数据

    Raises:
        FileNotFoundError: 文件不存在
        ValidationError: JSON 格式错误

    Examples:
        >>> data = parse_json("config.json")
        >>> print(data['title'])
        '文档标题'
    """
    import json

    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")

    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        raise ValidationError(f"JSON 格式错误: {e}")


def ensure_directory(file_path: Union[str, Path]) -> None:
    """确保目录存在，不存在则创建

    Args:
        file_path: 文件路径（会自动创建父目录）

    Examples:
        >>> ensure_directory("output/docs/report.docx")
        # 会自动创建 output/docs/ 目录
    """
    path = Path(file_path)
    if path.suffix:  # 如果是文件路径
        path.parent.mkdir(parents=True, exist_ok=True)
    else:  # 如果是目录路径
        path.mkdir(parents=True, exist_ok=True)


def parse_date_string(date_str: str) -> tuple:
    """解析日期字符串，分离数字和年月日

    用于填充日期时，将数字部分和年月日部分分开，
    以便应用不同的字体。

    Args:
        date_str: 日期字符串，如 "2024年1月15日"

    Returns:
        tuple: (数字部分列表, 年月日部分列表)
        例如：(['2024', '01', '15'], ['年', '月', '日'])

    Examples:
        >>> parse_date_string("2024年1月15日")
        (['2024', '01', '15'], ['年', '月', '日'])
    """
    import re

    # 匹配数字和中文
    pattern = r'(\d+)([年月日])'
    matches = re.findall(pattern, date_str)

    numbers = []
    separators = []

    for num, sep in matches:
        numbers.append(num.zfill(2))  # 不足两位补零
        separators.append(sep)

    return numbers, separators
