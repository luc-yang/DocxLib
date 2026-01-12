# DocxLib Word 文档处理库 - 开发需求文档

| 文档版本 | v1.0 |
|---------|------|
| 创建日期 | 2024-01-15 |
| 文档状态 | 待开发 |
| 项目类型 | Python 库开发 |

---

## 文档修订历史

| 版本 | 日期 | 修订人 | 修订内容 |
|------|------|--------|----------|
| v1.0 | 2024-01-15 | - | 初始版本 |

---

## 目录

1. [项目概述](#1-项目概述)
2. [需求分析](#2-需求分析)
3. [功能需求](#3-功能需求)
4. [非功能需求](#4-非功能需求)
5. [系统架构设计](#5-系统架构设计)
6. [API 设计规范](#6-api-设计规范)
7. [数据结构设计](#7-数据结构设计)
8. [模块详细设计](#8-模块详细设计)
9. [测试需求](#9-测试需求)
10. [部署与发布](#10-部署与发布)
11. [附录](#11-附录)

---

## 1. 项目概述

### 1.1 项目背景

开发一个独立的、易用的个人 Word 文档处理库，用于自动化处理 Word 文档的字段填充、样式应用、格式转换等常见操作。该项目基于 Spire.Doc 引擎，提供简洁的 Python API。

### 1.2 项目目标

- **简单直接**：提供直观的函数式 API，易于使用
- **功能完整**：覆盖文档处理、字段填充、样式应用全流程
- **灵活强大**：支持通配符、多种填充模式、批量处理
- **便于维护**：代码结构清晰，便于后续扩展

### 1.3 项目范围

**包含功能**：
- Word 文档加载、保存、合并
- 表格单元格定位和遍历
- 文本、图片、日期、网格数据填充
- 样式应用（字体、颜色、格式等）
- 文档格式转换（PDF、图片）

**不包含功能**：
- VBA 宏处理
- 复杂文档结构编辑（如插入节、页眉页脚等）
- 文档版本控制
- 协作编辑功能

### 1.4 技术栈

| 技术项 | 选型 | 说明 |
|--------|------|------|
| 开发语言 | Python 3.8+ | 核心开发语言 |
| 库名称 | DocxLib | Python 库名称 |
| 底层引擎 | Spire.Doc | Word 文档处理引擎 |
| 图片处理 | Pillow | 图片缩放和格式转换 |
| 测试框架 | pytest | 单元测试（可选） |
| 数据处理 | pandas | Excel 数据读取（可选） |

### 1.5 许可说明

**Spire.Doc 免费版限制**：
- 最多 500 段落
- 最多 25 个表格
- 转换 PDF 时有水印
- 仅限非商业用途

**注意**：如需商业使用或超出免费版限制，需要购买 Spire.Doc 商业版许可证。

---

## 2. 需求分析

### 2.1 用户角色

| 角色 | 描述 | 主要需求 |
|------|------|----------|
| 个人开发者 | 需要自动化处理 Word 文档的个人用户 | 简单易用的 API，清晰的使用文档 |
| 脚本编写者 | 编写自动化办公脚本的用户 | 批量处理能力，稳定的性能 |
| 项目集成者 | 将库集成到项目中的开发者 | 完整的功能，良好的错误处理 |

### 2.2 使用场景

**场景1：合同自动生成**
- 从 Excel 数据批量生成合同文档
- 填充合同编号、甲方乙方、金额、日期等字段
- 自动插入公司印章图片

**场景2：报表批量生成**
- 从数据库或 Excel 读取数据
- 批量生成月度/季度报表
- 支持多种报表模板

**场景3：文档格式转换**
- 将 Word 文档转换为 PDF
- 将 Word 文档转换为图片
- 批量转换多个文档

### 2.3 核心用例

```
用例1：基础字段填充
  - 用户加载 Word 模板
  - 用户指定填充位置和值
  - 系统填充字段到文档
  - 用户保存文档

用例2：批量文档生成
  - 用户加载模板和 Excel 数据
  - 系统遍历数据行
  - 对每行数据：复制模板 → 填充字段 → 保存文档
  - 系统生成多个文档

用例3：网格数据填充
  - 用户准备二维数组数据
  - 用户指定起始位置
  - 系统按行列填充数据到表格
```

---

## 3. 功能需求

### 3.1 文档操作模块 (document.py)

#### 3.1.1 加载文档

**功能描述**：从文件路径或字节数据加载 Word 文档

**函数签名**：
```python
def load_docx(source: Union[str, bytes]) -> Document:
    """加载文档

    Args:
        source: 文件路径（str）或字节数据（bytes）

    Returns:
        Document: Spire.Doc Document 对象

    Raises:
        DocumentError: 文件不存在或格式错误
        ValidationError: 文件格式不是 .docx
    """
```

**验收标准**：
- 支持从文件路径加载
- 支持从字节数据加载
- 文件不存在时抛出 DocumentError
- 非 .docx 格式抛出 ValidationError

#### 3.1.2 保存文档

**功能描述**：将文档保存到指定路径

**函数签名**：
```python
def save_docx(doc: Document, target: str) -> None:
    """保存文档

    Args:
        doc: Document 对象
        target: 保存路径

    Raises:
        DocumentError: 保存失败
    """
```

**验收标准**：
- 自动创建不存在的目录
- 成功保存为 .docx 格式
- 保存失败时抛出 DocumentError

#### 3.1.3 合并文档

**功能描述**：将多个文档合并为一个

**函数签名**：
```python
def merge_docs(doc_list: List[Document]) -> Document:
    """合并多个文档

    Args:
        doc_list: Document 对象列表

    Returns:
        Document: 合并后的 Document 对象
    """
```

#### 3.1.4 格式转换

**功能描述**：将文档转换为 PDF 或图片

**函数签名**：
```python
def to_pdf(doc: Document) -> bytes:
    """转换为 PDF

    Returns:
        bytes: PDF 文件字节数据
    """

def to_images(doc: Document) -> List[bytes]:
    """转换为图片列表

    Returns:
        List[bytes]: 图片字节数据列表
    """
```

### 3.2 表格操作模块 (table.py)

#### 3.2.1 获取单元格

**功能描述**：根据位置获取指定单元格

**函数签名**：
```python
def get_cell(doc: Document, section: int, table: int,
             row: int, col: int) -> Cell:
    """获取指定位置的单元格

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
    """
```

**位置定义**：
- 索引从 1 开始（不是 0）
- 位置格式：(section, table, row, col)
- 示例：(1, 1, 2, 2) = 第1节、第1个表格、第2行、第2列

#### 3.2.2 通配符获取单元格

**功能描述**：使用通配符获取多个单元格（0 表示所有）

**函数签名**：
```python
def get_cells(doc: Document, section: int = 0, table: int = 0,
              row: int = 0, col: int = 0) -> List[Tuple]:
    """通配符获取单元格

    Args:
        doc: Document 对象
        section: 节索引（0表示所有）
        table: 表格索引（0表示所有）
        row: 行索引（0表示所有）
        col: 列索引（0表示所有）

    Returns:
        List[Tuple]: [(section, table, row, col, cell), ...]
    """
```

#### 3.2.3 查找文本

**功能描述**：查找文档中包含指定文本的所有单元格位置

**函数签名**：
```python
def find_text(doc: Document, text: str) -> List[Position]:
    """查找文本位置

    Args:
        doc: Document 对象
        text: 要查找的文本

    Returns:
        List[Position]: 位置列表 [(section, table, row, col), ...]
    """
```

#### 3.2.4 遍历单元格

**功能描述**：遍历文档中所有单元格

**函数签名**：
```python
def iterate_cells(doc: Document) -> Generator:
    """遍历所有单元格

    Yields:
        (section, table, row, col, cell)
    """
```

### 3.3 字段填充模块 (fill.py)

#### 3.3.1 文本填充

**功能描述**：填充文本到指定单元格，支持三种模式

**函数签名**：
```python
def fill_text(
    doc: Document,
    position: PositionOrText,
    value: str,
    mode: str = "position",
    font_name: str = "仿宋_GB2312",
    font_size: float = 10.5,
    color: str = "black",
    bold: bool = False,
    italic: bool = False,
    underline: bool = False
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
    """
```

**类型定义**：
```python
Position = Tuple[int, int, int, int]  # (section, table, row, col)
PositionOrText = Union[Position, str]
```

**使用示例**：
```python
# 直接定位
fill_text(doc, (1, 1, 2, 2), "测试文本")

# 右侧填充
fill_text(doc, "姓名：", "张三", mode="match_right")

# 下方填充
fill_text(doc, "项目1", "智慧城市", mode="match_down")

# 带样式
fill_text(doc, "标题", "内容", font_name="黑体", font_size=16, bold=True)
```

#### 3.3.2 图片填充

**功能描述**：填充图片到指定单元格

**函数签名**：
```python
def fill_image(
    doc: Document,
    position: PositionOrText,
    image_path: str,
    mode: str = "position",
    width: Optional[float] = None,
    height: Optional[float] = None,
    maintain_ratio: bool = True
) -> None:
    """填充图片到文档

    Args:
        doc: Document 对象
        position: 位置元组或查找文本
        image_path: 图片文件路径
        mode: 填充模式（同 fill_text）
        width: 宽度（磅）
        height: 高度（磅）
        maintain_ratio: 是否保持宽高比

    Raises:
        FillError: 图片文件不存在或格式不支持
    """
```

**支持的图片格式**：PNG, JPG, JPEG, GIF, BMP

#### 3.3.3 日期填充

**功能描述**：填充日期，数字和年月日使用不同字体

**函数签名**：
```python
def fill_date(
    doc: Document,
    position: PositionOrText,
    date_str: str,
    font_name: str = "仿宋_GB2312",
    font_size: float = 10.5
) -> None:
    """填充日期

    Args:
        doc: Document 对象
        position: 位置元组或查找文本
        date_str: 日期字符串，如 "2024年1月15日"
        font_name: 数字字体（年月日使用宋体）
        font_size: 字体大小

    Note:
        数字部分使用 font_name，年月日部分使用宋体
    """
```

#### 3.3.4 网格数据填充

**功能描述**：从二维数组填充数据到表格

**函数签名**：
```python
def fill_grid(
    doc: Document,
    data: List[List[str]],
    position: Position
) -> None:
    """填充网格数据

    Args:
        doc: Document 对象
        data: 二维数组，每个元素代表一个单元格的值
        position: 起始位置 (section, table, row, col)

    Raises:
        PositionError: 数据超出表格边界

    Example:
        data = [
            ["序号", "项目", "金额"],
            ["1", "设备费", "50000"],
            ["2", "人工费", "30000"],
        ]
        fill_grid(doc, data, position=(1, 1, 7, 1))
    """
```

#### 3.3.5 全局替换

**功能描述**：全局替换文档中的文本

**函数签名**：
```python
def replace_all(doc: Document, old_text: str, new_text: str) -> None:
    """全局替换文本"""
```

### 3.4 样式管理模块 (style.py)

#### 3.4.1 颜色解析

**功能描述**：解析颜色字符串为 Color 对象

**函数签名**：
```python
def parse_color(color_str: str) -> Color:
    """解析颜色字符串

    Args:
        color_str: 颜色字符串
            - 颜色名称: 'black', 'red', 'blue' 等
            - 十六进制: '#FF0000' 或 'FF0000'

    Returns:
        Color: Spire.Doc Color 对象
        解析失败返回黑色
    """
```

**支持的颜色名称**：
```
black, red, blue, green, yellow, white, gray,
silver, maroon, purple, orange, pink
```

#### 3.4.2 应用字体样式

**函数签名**：
```python
def apply_font_style(
    run: TextRange,
    font_name: str,
    font_size: float,
    color: str,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False
) -> None:
    """应用字体样式到文本范围"""
```

### 3.5 错误处理模块 (errors.py)

**异常类层次结构**：

```python
class DocxLibError(Exception):
    """基础异常类"""

class DocumentError(DocxLibError):
    """文档操作错误"""

class PositionError(DocxLibError):
    """位置定位错误"""

class FillError(DocxLibError):
    """字段填充错误"""

class ValidationError(DocxLibError):
    """数据验证错误"""
```

**异常使用原则**：
- 文件不存在或无法读取 → `DocumentError`
- 位置越界或无效 → `PositionError`
- 数据格式错误 → `ValidationError`
- 填充操作失败 → `FillError`
- 颜色解析失败 → 返回黑色（静默）

### 3.6 常量定义模块 (constants.py)

```python
# 默认值
DEFAULT_FONT = "仿宋_GB2312"
DEFAULT_FONT_SIZE = 10.5
DEFAULT_COLOR = "black"

# 支持的图片格式
SUPPORTED_IMAGE_FORMATS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp']

# 文件格式
class FileFormat:
    DOC = ".doc"
    DOCX = ".docx"
    PDF = ".pdf"
    PNG = ".png"
    JPEG = ".jpeg"

# 填充模式
class FillMode:
    POSITION = "position"
    MATCH_RIGHT = "match_right"
    MATCH_DOWN = "match_down"
```

---

## 4. 非功能需求

### 4.1 性能要求

| 操作 | 性能指标 | 测试条件 |
|------|---------|---------|
| 加载文档 | < 1s | 5MB 以内文档 |
| 单字段填充 | < 100ms | 直接定位模式 |
| 100字段批量填充 | < 5s | 模板复用 |
| 1000文档批量生成 | < 5min | 多进程 + 分批 |
| 文档转PDF | < 3s/页 | 单页文档 |
| 内存占用 | < 200MB | 单文档处理 |

### 4.2 兼容性要求

#### 4.2.1 Python 版本
- **最低版本**：Python 3.8
- **推荐版本**：Python 3.9+
- **测试覆盖**：Python 3.8, 3.9, 3.10, 3.11

#### 4.2.2 操作系统
- **完全支持**：Windows 10/11, Windows 8/8.1
- **国产 Linux**：中标麒麟、中科方德（官方支持）
- **社区支持**：Ubuntu、Debian、CentOS（需自行测试）
- **不支持**：macOS（官方未明确支持）

#### 4.2.3 Word 版本
| Word 版本 | 支持状态 | 说明 |
|-----------|---------|------|
| Word 2007+ | ✅ 完全支持 | .docx 格式 |
| Word 2003- | ⚠️ 有限支持 | 需转换为 .docx |

#### 4.2.4 字体兼容性
**默认中文字体**（Windows 系统自带）：
- 仿宋_GB2312（默认）
- 黑体
- 宋体
- 楷体
- 微软雅黑

### 4.3 可靠性要求

- **错误处理**：所有公共 API 必须有异常处理
- **输入验证**：文件路径、位置参数、图片格式必须验证
- **边界检查**：数组访问前必须检查边界
- **资源释放**：及时释放文档对象，避免内存泄漏

### 4.4 可维护性要求

- **代码规范**：遵循 PEP 8
- **类型注解**：核心 API 必须有类型注解
- **文档字符串**：所有公共函数必须有 docstring
- **模块化**：功能分离，模块职责清晰

### 4.5 易用性要求

- **简单安装**：`pip install docxlib`
- **清晰文档**：提供完整的使用示例
- **友好错误**：错误信息清晰易懂
- **示例代码**：提供常见场景的示例

---

## 5. 系统架构设计

### 5.1 项目结构

```
docxlib/
├── docxlib/
│   ├── __init__.py            # 导出主要的类和函数
│   ├── document.py            # 文档加载、保存、合并、转换
│   ├── table.py               # 表格遍历、单元格操作
│   ├── fill.py                # 字段填充（文本、图片、日期、网格）
│   ├── style.py               # 样式管理
│   ├── errors.py              # 异常类定义
│   ├── constants.py           # 常量定义
│   └── utils.py               # 工具函数
├── examples/                  # 示例代码
│   ├── basic_usage.py
│   ├── batch_processing.py
│   └── advanced_features.py
├── tests/                     # 单元测试
│   ├── __init__.py
│   ├── test_document.py
│   ├── test_fill.py
│   └── test_style.py
├── fixtures/                  # 测试数据
│   ├── templates/
│   │   └── simple.docx
│   └── images/
│       └── logo.png
├── README.md
├── setup.py
└── requirements.txt
```

### 5.2 模块依赖关系

```
┌─────────────┐
│   document  │◄──────┐
└──────┬──────┘       │
       │              │
       ▼              │
┌─────────────┐       │
│    table    │───────┘
└──────┬──────┘
       │
       ▼
┌─────────────┐
│    fill     │
└──────┬──────┘
       │
   ┌───┴────┬─────────┐
   ▼        ▼         ▼
┌──────┐ ┌──────┐ ┌──────┐
│ style│ │errors│ │const │
└──────┘ └──────┘ └──────┘
   │        │         │
   └────────┴─────────┘
            ▼
       ┌──────────┐
       │  utils   │
       └──────────┘
```

### 5.3 设计原则

- **简单直接**：不追求设计模式，代码直观易懂
- **功能优先**：实现核心功能，架构可以简单
- **易于修改**：代码结构清晰，方便后续调整
- **最小依赖**：只依赖必要的第三方库

---

## 6. API 设计规范

### 6.1 命名规范

- **函数名**：小写，下划线分隔（如 `load_docx`）
- **类名**：大驼峰（如 `DocumentError`）
- **常量**：大写，下划线分隔（如 `DEFAULT_FONT`）
- **参数名**：小写，下划线分隔（如 `font_name`）

### 6.2 函数设计原则

1. **参数顺序**：
   ```python
   def func(doc, position, value, mode="position", **options):
       # 1. 必需参数：doc, position, value
       # 2. 可选参数：mode
       # 3. 样式参数：**options
   ```

2. **返回值**：
   - 有返回值时返回结果
   - 无返回值时返回 None
   - 失败时抛出异常（不返回错误码）

3. **默认参数**：
   - 所有样式参数提供合理默认值
   - 默认字体：仿宋_GB2312
   - 默认字号：10.5
   - 默认颜色：black

### 6.3 类型注解规范

```python
from typing import Union, Tuple, Optional, List

# 类型别名
Position = Tuple[int, int, int, int]  # (section, table, row, col)
PositionOrText = Union[Position, str]

# 函数注解
def fill_text(
    doc: Document,
    position: PositionOrText,
    value: str,
    mode: str = "position",
    **styles
) -> None:
    """填充文本到文档"""
```

---

## 7. 数据结构设计

### 7.1 位置元组

```python
Position = Tuple[int, int, int, int]
# 格式：(section, table, row, col)
# 索引：从 1 开始
# 示例：(1, 1, 2, 2) = 第1节、第1个表格、第2行、第2列
```

### 7.2 网格数据

```python
GridData = List[List[str]]
# 格式：二维数组
# 示例：
# [
#     ["序号", "项目", "金额"],
#     ["1", "设备费", "50000"],
#     ["2", "人工费", "30000"],
# ]
```

### 7.3 样式参数

```python
StyleParams = {
    'font_name': str,      # 字体名称
    'font_size': float,    # 字体大小
    'color': str,          # 颜色
    'bold': bool,          # 粗体
    'italic': bool,        # 斜体
    'underline': bool      # 下划线
}
```

---

## 8. 模块详细设计

### 8.1 document.py

**职责**：文档的加载、保存、合并、转换

**主要函数**：
- `load_docx(source: Union[str, bytes]) -> Document`
- `save_docx(doc: Document, target: str) -> None`
- `merge_docs(doc_list: List[Document]) -> Document`
- `to_pdf(doc: Document) -> bytes`
- `to_images(doc: Document) -> List[bytes]`

**依赖**：
- `spire.doc`
- `docxlib.errors`

### 8.2 table.py

**职责**：表格遍历、单元格操作

**主要函数**：
- `get_cell(doc, section, table, row, col) -> Cell`
- `get_cells(doc, section=0, table=0, row=0, col=0) -> List`
- `find_text(doc, text) -> List[Position]`
- `iterate_cells(doc) -> Generator`

**依赖**：
- `spire.doc`
- `docxlib.errors`

### 8.3 fill.py

**职责**：字段填充（文本、图片、日期、网格）

**主要函数**：
- `fill_text(doc, position, value, mode="position", **styles) -> None`
- `fill_image(doc, position, image_path, mode="position", ...) -> None`
- `fill_date(doc, position, date_str, font_name, font_size) -> None`
- `fill_grid(doc, data, position) -> None`
- `replace_all(doc, old_text, new_text) -> None`

**依赖**：
- `spire.doc`
- `docxlib.table`
- `docxlib.style`
- `docxlib.errors`
- `PIL` (图片处理)

### 8.4 style.py

**职责**：样式管理

**主要函数**：
- `parse_color(color_str) -> Color`
- `apply_font_style(run, font_name, font_size, color, ...) -> None`

**常量**：
- `COLOR_MAP`：颜色名称映射表

**依赖**：
- `spire.doc`
- `docxlib.constants`

### 8.5 errors.py

**职责**：异常类定义

**异常类**：
- `DocxLibError`：基础异常类
- `DocumentError`：文档操作错误
- `PositionError`：位置定位错误
- `FillError`：字段填充错误
- `ValidationError`：数据验证错误

**依赖**：无

### 8.6 constants.py

**职责**：常量定义

**常量**：
- `DEFAULT_FONT`：默认字体
- `DEFAULT_FONT_SIZE`：默认字号
- `DEFAULT_COLOR`：默认颜色
- `COLOR_MAP`：颜色映射表
- `SUPPORTED_IMAGE_FORMATS`：支持的图片格式
- `FileFormat`：文件格式类
- `FillMode`：填充模式类

**依赖**：
- `spire.doc`

### 8.7 utils.py

**职责**：工具函数

**主要函数**：
- `is_valid_docx(data) -> bool`：验证 DOCX 格式
- `parse_csv(file_path) -> List`：解析 CSV 文件
- `parse_json(file_path) -> Dict`：解析 JSON 文件

**依赖**：无

---

## 9. 测试需求

### 9.1 测试策略

- **单元测试**：测试单个函数的功能
- **集成测试**：测试模块间的协作
- **示例测试**：通过运行示例代码验证功能

### 9.2 测试覆盖率目标

| 模块 | 目标覆盖率 | 说明 |
|------|-----------|------|
| document.py | ≥ 90% | 核心功能 |
| table.py | ≥ 90% | 核心功能 |
| fill.py | ≥ 85% | 核心功能 |
| style.py | ≥ 80% | 支持功能 |
| utils.py | ≥ 75% | 工具函数 |
| **整体** | **≥ 80%** | **平均目标** |

### 9.3 测试用例示例

#### 9.3.1 文档加载测试

```python
def test_load_docx_success():
    """测试成功加载文档"""
    doc = load_docx("fixtures/templates/simple.docx")
    assert doc is not None
    assert doc.Sections.Count > 0

def test_load_docx_file_not_exists():
    """测试文件不存在时抛出异常"""
    with pytest.raises(DocumentError):
        load_docx("nonexistent.docx")

def test_load_docx_invalid_format():
    """测试无效格式时抛出异常"""
    with pytest.raises(ValidationError):
        load_docx("test.txt")
```

#### 9.3.2 字段填充测试

```python
def test_fill_text_by_position():
    """测试按位置填充文本"""
    doc = load_docx("fixtures/templates/simple.docx")
    fill_text(doc, (1, 1, 2, 2), "测试文本")

    cell = get_cell(doc, 1, 1, 2, 2)
    text = cell.Range.Text.strip()
    assert "测试文本" in text

def test_fill_text_by_match_right():
    """测试 match_right 模式填充"""
    doc = load_docx("fixtures/templates/simple.docx")
    fill_text(doc, "姓名", "张三", mode="match_right")

    positions = find_text(doc, "张三")
    assert len(positions) > 0

def test_fill_text_invalid_position():
    """测试无效位置时抛出异常"""
    doc = load_docx("fixtures/templates/simple.docx")
    with pytest.raises(PositionError):
        fill_text(doc, (99, 99, 99, 99), "测试")
```

#### 9.3.3 网格填充测试

```python
def test_fill_grid_success():
    """测试成功填充网格数据"""
    doc = load_docx("fixtures/templates/simple.docx")

    data = [
        ["序号", "项目", "金额"],
        ["1", "设备费", "50000"],
        ["2", "人工费", "30000"],
    ]
    fill_grid(doc, data, position=(1, 1, 7, 1))

    cell = get_cell(doc, 1, 1, 7, 1)
    assert cell.Range.Text.strip() == "序号"

def test_fill_grid_out_of_bounds():
    """测试数据超出边界时抛出异常"""
    doc = load_docx("fixtures/templates/simple.docx")
    data = [["测试"] * 100]

    with pytest.raises(PositionError):
        fill_grid(doc, data, position=(1, 1, 1, 1))
```

### 9.4 测试数据准备

```
fixtures/
├── templates/
│   ├── simple.docx         # 简单模板
│   ├── complex.docx        # 复杂模板
│   └── large.docx          # 大型文档
├── images/
│   ├── logo.png
│   └── seal.png
└── data/
    ├── test.json
    ├── test.csv
    └── test.xlsx
```

### 9.5 测试执行

```bash
# 安装测试依赖
pip install pytest pytest-cov

# 运行所有测试
pytest tests/ -v

# 生成覆盖率报告
pytest tests/ --cov=docxlib --cov-report=html

# 运行特定测试
pytest tests/test_document.py::test_load_docx_success -v

# 运行包含关键字的测试
pytest tests/ -k "fill" -v
```

---

## 10. 部署与发布

### 10.1 安装方式

#### 10.1.1 从源码安装

```bash
# 克隆仓库
git clone https://github.com/xxx/docxlib.git
cd docxlib

# 安装
pip install -e .
```

#### 10.1.2 从 PyPI 安装

```bash
pip install docxlib
```

### 10.2 依赖管理

**核心依赖**（requirements.txt）：
```
spire-doc>=10.0.0
Pillow>=9.0.0
```

**可选依赖**：
```
pandas>=1.3.0    # Excel 数据读取
openpyxl>=3.0.0  # Excel 文件支持
```

**开发依赖**：
```
pytest>=7.0.0
pytest-cov>=3.0.0
```

### 10.3 版本管理

**版本号格式**：`主版本.次版本.修订号`

- `0.1.0`：初始版本
- `0.2.0`：新增功能
- `0.2.1`：Bug 修复
- `1.0.0`：稳定版本

### 10.4 发布流程

1. 更新版本号（setup.py）
2. 运行测试套件
3. 生成文档
4. 构建 Python 包
5. 上传到 PyPI
6. 发布 GitHub Release

---

## 11. 附录

### 11.1 API 快速参考

```python
from docxlib import load_docx, fill_text, fill_image, fill_date, fill_grid, save_docx

# 加载文档
doc = load_docx("template.docx")

# 文本填充
fill_text(doc, (1, 1, 2, 2), "测试文本")
fill_text(doc, "姓名：", "张三", mode="match_right")

# 图片填充
fill_image(doc, "印章：", "seal.png", mode="match_right", width=100, height=100)

# 日期填充
fill_date(doc, (1, 1, 4, 2), "2024年1月15日")

# 网格填充
data = [["1", "设备费", "50000"]]
fill_grid(doc, data, position=(1, 1, 7, 1))

# 保存文档
save_docx(doc, "output.docx")
```

### 11.2 最佳实践

#### 11.2.1 批量处理优化

```python
# 好的做法：模板复用
template = load_docx("template.docx")
for idx, row in data.iterrows():
    doc = copy.deepcopy(template)
    fill_text(doc, "姓名", row["姓名"], mode="match_right")
    save_docx(doc, f"output_{idx}.docx")

# 不好的做法：重复加载
for idx, row in data.iterrows():
    doc = load_docx("template.docx")  # 慢
    fill_text(doc, "姓名", row["姓名"], mode="match_right")
    save_docx(doc, f"output_{idx}.docx")
```

#### 11.2.2 错误处理

```python
from docxlib.errors import DocumentError, PositionError, FillError

try:
    doc = load_docx("template.docx")
    fill_text(doc, (1, 1, 2, 2), "测试")
    save_docx(doc, "output.docx")
except DocumentError as e:
    print(f"文档错误: {e}")
except PositionError as e:
    print(f"位置错误: {e}")
except FillError as e:
    print(f"填充错误: {e}")
```

#### 11.2.3 样式管理

```python
# 统一样式配置
STYLES = {
    "title": {"font_name": "黑体", "font_size": 16, "bold": True},
    "body": {"font_name": "宋体", "font_size": 12},
    "note": {"font_name": "楷体", "font_size": 10}
}

fill_text(doc, "标题", "内容", **STYLES["title"])
```

### 11.3 常见问题

**Q1: 填充后内容位置不对？**

A: 使用 `iterate_cells` 查看实际位置
```python
for sec_idx, tbl_idx, row_idx, col_idx, cell in iterate_cells(doc):
    text = cell.Range.Text.strip()
    if "姓名" in text:
        print(f"找到: ({sec_idx}, {tbl_idx}, {row_idx}, {col_idx})")
```

**Q2: 图片填充失败？**

A: 检查图片路径和格式
```python
from pathlib import Path
image_path = Path("images/logo.png")
assert image_path.exists()
assert image_path.suffix.lower() in ['.png', '.jpg', '.jpeg']
```

**Q3: 批量处理内存溢出？**

A: 使用分批处理
```python
batch_size = 100
for i in range(0, len(data), batch_size):
    batch = data[i:i+batch_size]
    # 处理批次...
    gc.collect()
```

### 11.4 性能优化建议

1. **避免重复加载**：使用 `copy.deepcopy` 复制模板
2. **使用直接定位**：`position` 模式比 `match` 模式快
3. **分批处理**：大数据集使用 `chunksize=100-500`
4. **及时释放**：处理完调用 `del doc`
5. **定期回收**：每 100 次循环调用 `gc.collect()`

### 11.5 已知限制

1. **Spire.Doc 免费版**
   - 最多 500 段落
   - 最多 25 个表格
   - PDF 转换有水印

2. **平台支持**
   - 主要支持 Windows
   - macOS/Linux 有限支持

3. **格式支持**
   - 不支持 VBA 宏
   - 复杂样式可能不完全保留

### 11.6 参考资源

- **Spire.Doc 官方文档**：[https://www.e-iceblue.com/](https://www.e-iceblue.com/)
- **Python 文档**：[https://docs.python.org/3/](https://docs.python.org/3/)
- **PEP 8**：[https://peps.python.org/pep-0008/](https://peps.python.org/pep-0008/)

---

**文档结束**
