"""
DocxLib 常量定义

定义了库中使用的所有常量，包括默认值、文件格式、填充模式等。
"""

from typing import Dict


# ==================== 默认值 ====================

DEFAULT_FONT: str = "仿宋_GB2312"
DEFAULT_FONT_SIZE: float = 10.5
DEFAULT_COLOR: str = "black"


# ==================== 支持的图片格式 ====================

SUPPORTED_IMAGE_FORMATS: tuple = ('.png', '.jpg', '.jpeg', '.gif', '.bmp')


# ==================== 颜色映射表 ====================

COLOR_MAP: Dict[str, str] = {
    'black': '#000000',
    'red': '#FF0000',
    'blue': '#0000FF',
    'green': '#008000',
    'yellow': '#FFFF00',
    'white': '#FFFFFF',
    'gray': '#808080',
    'silver': '#C0C0C0',
    'maroon': '#800000',
    'purple': '#800080',
    'orange': '#FFA500',
    'pink': '#FFC0CB',
}


# ==================== 文件格式 ====================

class FileFormat:
    """文件格式常量"""

    DOC = ".doc"
    DOCX = ".docx"
    PDF = ".pdf"
    PNG = ".png"
    JPEG = ".jpeg"


# ==================== 填充模式 ====================

class FillMode:
    """填充模式常量"""

    POSITION = "position"
    MATCH_RIGHT = "match_right"
    MATCH_DOWN = "match_down"


# ==================== 对齐方式 ====================

class Alignment:
    """对齐方式常量"""

    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    JUSTIFY = "justify"


# ==================== 类型定义 ====================

# 位置元组类型：(section, table, row, col)
# 所有索引从 1 开始
Position = tuple


# ==================== Spire.Doc 相关 ====================

# 尝试导入 Spire.Doc，如果失败则设置为 None
try:
    from spire.doc import *
    SPIRE_AVAILABLE = True
except ImportError:
    SPIRE_AVAILABLE = False
