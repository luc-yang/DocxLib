# DocxLib

> 一个简单易用的 Word 文档处理库，基于 Spire.Doc 引擎

[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

## 特性

- **简单直接**：提供直观的函数式 API，易于使用
- **功能完整**：覆盖文档处理、字段填充、样式应用全流程
- **灵活强大**：支持通配符、多种填充模式、批量处理
- **便于维护**：代码结构清晰，便于后续扩展

## 主要功能

- 文档加载、保存、合并
- 表格单元格定位和遍历
- 文本、图片、日期、网格数据填充
- 样式应用（字体、颜色、格式等）
- 文档格式转换（PDF、图片）

## 安装

```bash
pip install docxlib
```

或从源码安装：

```bash
git clone https://github.com/luc-yang/DocxLib.git
cd docxlib
pip install -e .
```

## 快速开始

### 基础用法

```python
from docxlib import load_docx, fill_text, save_docx

# 加载模板
doc = load_docx("sample.docx")

# 填充内容
fill_text(doc, (1, 1, 2, 2), "测试文本")

# 保存文档
save_docx(doc, "output.docx")
```

### 多种填充模式

```python
from docxlib import load_docx, fill_text, save_docx

doc = load_docx("sample.docx")

# 直接定位填充
fill_text(doc, (1, 1, 2, 2), "测试文本")

# 右侧填充（查找"姓名："并在右侧填充）
fill_text(doc, "姓名：", "张三", mode="match_right")

# 下方填充（查找"项目1"并在下方填充）
fill_text(doc, "项目1", "智慧城市", mode="match_down")

# 带样式填充
fill_text(doc, "标题", "内容",
          font_name="黑体",
          font_size=16,
          bold=True,
          color="red")

save_docx(doc, "output.docx")
```

### 图片填充

```python
from docxlib import load_docx, fill_image, save_docx

doc = load_docx("sample.docx")

# 填充图片
fill_image(doc, (1, 1, 2, 2), "logo.png",
           width=100, height=100)

# 右侧填充图片
fill_image(doc, "印章：", "seal.png",
           mode="match_right",
           width=80, height=80)

save_docx(doc, "output.docx")
```

### 日期填充

```python
from docxlib import load_docx, fill_date, save_docx

doc = load_docx("sample.docx")

# 填充日期（数字和年月日使用不同字体）
fill_date(doc, (1, 1, 4, 2), "2024年1月15日")

save_docx(doc, "output.docx")
```

### 网格数据填充

```python
from docxlib import load_docx, fill_grid, save_docx

doc = load_docx("sample.docx")

# 填充网格数据
data = [
    ["序号", "项目", "金额"],
    ["1", "设备费", "50000"],
    ["2", "人工费", "30000"],
    ["3", "材料费", "20000"],
]
fill_grid(doc, data, position=(1, 1, 7, 1))

save_docx(doc, "output.docx")
```

### 批量文档生成

```python
import copy
from docxlib import load_docx, fill_text, save_docx

# 加载模板（只加载一次）
template = load_docx("sample.docx")

# 批量生成文档
data = [
    {"name": "张三", "amount": "50000"},
    {"name": "李四", "amount": "30000"},
    {"name": "王五", "amount": "20000"},
]

for i, item in enumerate(data):
    # 复制模板
    doc = copy_doc(template)

    # 填充数据
    fill_text(doc, "姓名：", item["name"], mode="match_right")
    fill_text(doc, "金额：", item["amount"], mode="match_right")

    # 保存文档
    save_docx(doc, f"output_{i+1}.docx")
```

### 文档格式转换

```python
from docxlib import load_docx, to_pdf, to_images

doc = load_docx("document.docx")

# 转换为 PDF
pdf_bytes = to_pdf(doc)
with open("output.pdf", "wb") as f:
    f.write(pdf_bytes)

# 转换为图片
images = to_images(doc)
for i, img_bytes in enumerate(images):
    with open(f"page_{i+1}.png", "wb") as f:
        f.write(img_bytes)
```

### 遍历单元格

```python
from docxlib import load_docx, iterate_cells

doc = load_docx("sample.docx")

# 遍历所有单元格
for sec, tbl, row, col, cell in iterate_cells(doc):
    text = cell.Range.Text.strip()
    if text:
        print(f"({sec}, {tbl}, {row}, {col}): {text}")
```

### 查找文本

```python
from docxlib import load_docx, find_text

doc = load_docx("sample.docx")

# 查找包含"姓名"的单元格
positions = find_text(doc, "姓名")
print(f"找到 {len(positions)} 个匹配项")
for pos in positions:
    print(f"位置: {pos}")
```

## 位置说明

所有索引从 **1** 开始（不是 0）：

- `section`: 节索引
- `table`: 表格索引
- `row`: 行索引
- `col`: 列索引

示例：`(1, 1, 2, 2)` = 第1节、第1个表格、第2行、第2列

## API 参考

### 文档操作

| 函数 | 说明 |
|------|------|
| `load_docx(source)` | 加载文档 |
| `save_docx(doc, target)` | 保存文档 |
| `merge_docs(doc_list)` | 合并文档 |
| `to_pdf(doc)` | 转换为 PDF |
| `to_images(doc)` | 转换为图片 |
| `copy_doc(doc)` | 复制文档 |

### 表格操作

| 函数 | 说明 |
|------|------|
| `get_cell(doc, s, t, r, c)` | 获取单元格 |
| `get_cells(doc, ...)` | 通配符获取单元格 |
| `find_text(doc, text)` | 查找文本 |
| `iterate_cells(doc)` | 遍历单元格 |

### 字段填充

| 函数 | 说明 |
|------|------|
| `fill_text(doc, pos, val, ...)` | 填充文本 |
| `fill_image(doc, pos, path, ...)` | 填充图片 |
| `fill_date(doc, pos, date)` | 填充日期 |
| `fill_grid(doc, data, pos)` | 填充网格数据 |
| `replace_all(doc, old, new)` | 全局替换 |

## 注意事项

### Spire.Doc 免费版限制

- 最多 500 段落
- 最多 25 个表格
- 转换 PDF 时有水印
- 仅限非商业用途

如需商业使用或超出免费版限制，需要购买 Spire.Doc 商业版许可证。

### 平台支持

- **完全支持**：Windows 10/11
- **国产 Linux**：中标麒麟、中科方德（官方支持）
- **社区支持**：Ubuntu、Debian、CentOS（需自行测试）

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！

## 链接

- [Spire.Doc 官方文档](https://www.e-iceblue.com/)
- [GitHub 仓库](https://github.com/luc-yang/DocxLib)
