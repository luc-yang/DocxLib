# DocxLib 安装和测试指南

## 安装

### 1. 克隆项目

```bash
git clone https://github.com/yourusername/docxlib.git
cd docxlib
```

### 2. 创建虚拟环境（推荐）

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux/Mac
source .venv/bin/activate
```

### 3. 安装依赖

```bash
pip install -e .
```

或使用 uv：

```bash
uv pip install -e .
```

### 4. 验证安装

```bash
python -c "from docxlib import __version__; print(__version__)"
```

应该输出：`0.1.0`

## 创建测试模板

由于无法直接创建 Word 文档，请按以下步骤手动创建测试模板：

### 创建基础测试模板 (fixtures/templates/sample.docx)

1. 打开 Microsoft Word
2. 创建一个新文档
3. 插入一个表格（5行 x 3列）
4. 填充以下内容：

| | 列1 | 列2 | 列3 |
|---|---|---|---|
| 行1 | 姓名 | | |
| 行2 | 年龄 | | |
| 行3 | 日期 | | |
| 行4 | 项目 | | |
| 行5 | | | |

5. 保存为 `fixtures/templates/sample.docx`

### 创建带图片的模板 (fixtures/templates/with_image.docx)

1. 打开 Microsoft Word
2. 创建一个新文档
3. 插入一个表格（3行 x 2列）
4. 填充以下内容：

| | 列1 | 列2 |
|---|---|---|
| 行1 | 照片： | |
| 行2 | 姓名： | |
| 行3 | | |

5. 保存为 `fixtures/templates/with_image.docx`

### 创建测试图片 (fixtures/images/logo.png)

找一个简单的 PNG 或 JPG 图片，复制到 `fixtures/images/` 目录，重命名为 `logo.png`

## 运行测试

### 基础功能测试

```bash
python tests/test_basic.py
```

### 运行示例

```bash
# 基础用法示例
python examples/basic_usage.py

# 批量处理示例
python examples/batch_processing.py

# 高级功能示例
python examples/advanced_features.py
```

### 运行单元测试

```bash
pytest tests/ -v
```

## 常见问题

### Q: 导入失败 "No module named 'spire.doc'"

A: 请先安装 spire-doc-free：

```bash
pip install spire-doc-free
```

### Q: 保存文档失败

A: 确保 output 目录存在，或程序有创建目录的权限。

### Q: 图片填充失败

A: 确保：
1. 图片文件存在
2. 图片格式为 PNG、JPG、JPEG、GIF 或 BMP
3. 图片路径正确（可以是相对路径或绝对路径）

### Q: 查找文本找不到

A: 使用 `iterate_cells` 函数查看实际的单元格内容：

```python
from docxlib import load_docx, iterate_cells

doc = load_docx("sample.docx")
for sec, tbl, row, col, cell in iterate_cells(doc):
    text = cell.Range.Text.strip()
    if text:
        print(f"({sec}, {tbl}, {row}, {col}): {text}")
```

## Spire.Doc 免费版限制

- 最多 500 段落
- 最多 25 个表格
- 转换 PDF 时有水印
- 仅限非商业用途

如需商业使用或超出免费版限制，需要购买 Spire.Doc 商业版许可证。

## 许可证

MIT License
