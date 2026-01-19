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

SUPPORTED_IMAGE_FORMATS: tuple = (".png", ".jpg", ".jpeg", ".bmp")


# ==================== 颜色映射表 ====================

COLOR_MAP: Dict[str, str] = {
    "black": "#000000",
    "red": "#FF0000",
    "blue": "#0000FF",
    "green": "#008000",
    "yellow": "#FFFF00",
    "white": "#FFFFFF",
    "gray": "#808080",
    "silver": "#C0C0C0",
    "maroon": "#800000",
    "purple": "#800080",
    "orange": "#FFA500",
    "pink": "#FFC0CB",
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


class MatchMode:
    """匹配模式常量（控制批量填充行为）"""

    ALL = "all"  # 填充所有匹配位置（默认）
    FIRST = "first"  # 仅填充第一个匹配位置


# ==================== 模板变量 ====================

# 模板变量默认值
DEFAULT_VAR_PREFIX: str = "${"  # 变量前缀
DEFAULT_VAR_SUFFIX: str = "}"  # 变量后缀
DEFAULT_MISSING_VAR_ACTION: str = "error"  # 缺失变量处理方式: "error", "ignore", "empty"


# ==================== 对齐方式 ====================
# 注意：对齐方式使用字符串字面量，定义在 docxlib.config 的 Literal 类型中
# 由于 from spire.doc import * 会导入 Spire.Doc 的 HorizontalAlignment/VerticalAlignment 类
# 因此不在本模块定义同名常量，避免命名冲突和混淆

# 对齐方式的有效值（仅供文档参考，实际使用时直接使用字符串字面量）：
# 水平对齐: "left", "center", "right", "justify"
# 垂直对齐: "top", "middle", "bottom"


# ==================== 类型定义 ====================

# 位置元组类型：(section, table, row, col)
# 所有索引从 1 开始计数
Position = tuple[int, int, int, int]


# ==================== Spire.Doc 相关 ====================

# 尝试导入 Spire.Doc，如果失败则设置为 None
try:
    from spire.doc import *

    SPIRE_AVAILABLE = True
except ImportError:
    SPIRE_AVAILABLE = False
