"""
DocxLib - Word 文档处理库

一个简单易用的 Word 文档处理库，基于 Spire.Doc 引擎。

主要功能：
    - 文档加载、保存、合并、格式转换
    - 表格单元格定位和遍历
    - 文本、图片、日期、网格数据填充
    - 样式应用（字体、颜色、格式等）

示例:
    >>> from docxlib import load_docx, fill_text, save_docx
    >>>
    >>> # 加载模板
    >>> doc = load_docx("template.docx")
    >>>
    >>> # 填充内容
    >>> fill_text(doc, (1, 1, 2, 2), "测试文本")
    >>> fill_text(doc, "姓名：", "张三", mode="match_right")
    >>>
    >>> # 保存文档
    >>> save_docx(doc, "output.docx")
"""

__version__ = "0.1.0"
__author__ = "DocxLib Contributors"

# ==================== 文档操作 ====================
from .document import (
    load_docx,
    save_docx,
    merge_docs,
    to_pdf,
    to_images,
    to_pdf_file,
    copy_doc,
)

# ==================== 表格操作 ====================
from .table import (
    get_cell,
    get_cells,
    find_text,
    iterate_cells,
    get_cell_text,
    get_table_dimensions,
    get_section_table_count,
    get_section_count,
)

# ==================== 字段填充 ====================
from .fill import (
    fill_text,
    fill_image,
    fill_date,
    fill_grid,
    replace_all,
    clear_cell,
)

# ==================== 样式管理 ====================
from .style import (
    parse_color,
    apply_font_style,
    set_cell_border,
)

# ==================== 异常类 ====================
from .errors import (
    DocxLibError,
    DocumentError,
    PositionError,
    FillError,
    ValidationError,
)

# ==================== 常量 ====================
from .constants import (
    DEFAULT_FONT,
    DEFAULT_FONT_SIZE,
    DEFAULT_COLOR,
    SUPPORTED_IMAGE_FORMATS,
    FileFormat,
    FillMode,
    Position,
)

# ==================== 工具函数 ====================
from .utils import (
    is_valid_docx,
    validate_docx,
    parse_csv,
    parse_json,
    ensure_directory,
    parse_date_string,
)

# ==================== 导出列表 ====================
__all__ = [
    # 版本信息
    "__version__",
    "__author__",
    # 文档操作
    "load_docx",
    "save_docx",
    "merge_docs",
    "to_pdf",
    "to_images",
    "to_pdf_file",
    "copy_doc",
    # 表格操作
    "get_cell",
    "get_cells",
    "find_text",
    "iterate_cells",
    "get_cell_text",
    "get_table_dimensions",
    "get_section_table_count",
    "get_section_count",
    # 字段填充
    "fill_text",
    "fill_image",
    "fill_date",
    "fill_grid",
    "replace_all",
    "clear_cell",
    # 样式管理
    "parse_color",
    "apply_font_style",
    "set_cell_border",
    # 异常类
    "DocxLibError",
    "DocumentError",
    "PositionError",
    "FillError",
    "ValidationError",
    # 常量
    "DEFAULT_FONT",
    "DEFAULT_FONT_SIZE",
    "DEFAULT_COLOR",
    "SUPPORTED_IMAGE_FORMATS",
    "FileFormat",
    "FillMode",
    "Position",
    # 工具函数
    "is_valid_docx",
    "validate_docx",
    "parse_csv",
    "parse_json",
    "ensure_directory",
    "parse_date_string",
]
