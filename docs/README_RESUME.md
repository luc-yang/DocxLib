# 简历模板示例

本示例演示如何使用 DocxLib 的读取模块（read.py）和填充模块（fill.py）来处理个人简历模板。

## 模板结构

个人简历模板是一个 12 行 x 7 列的表格，包含以下内容：

```
第1行: 姓名 | 性别 | 出生年月 | 民族 | 籍贯 | 政治面貌 | 身体状况
第2行: 入党（团）时间 | 参加工作时间
第3行: 专业特长
第4行: 专业职称
第5行: 获得职称时间
第6-9行: 学习经历（起止时间、在何院校学习专业、证明人）
第10行: 工作经历
第11行: 学习工作成果
第12行: 奖惩情况
```

## 功能展示

### 1. 读取简历内容

使用 `read_text()` 和 `read_cells()` 函数读取模板中的现有数据：

```python
from docxlib import load_docx, read_text, read_cells

doc = load_docx("fixtures/templates/sample.docx")

# 读取单个字段
name = read_text(doc, (1, 1, 1, 2), default="")  # 姓名
gender = read_text(doc, (1, 1, 1, 4), default="")  # 性别

# 批量读取学习经历
positions = [(1, 1, 6, 1), (1, 1, 6, 2), (1, 1, 6, 3)]  # 第1行的学习经历
education = read_cells(doc, positions, default="")
```

### 2. 填充简历数据

使用 `fill_text()` 和 `fill_grid()` 函数填充简历数据：

```python
from docxlib import fill_text, fill_grid, FontStyle

# 填充基本信息（带样式）
fill_text(
    doc, (1, 1, 1, 2), "张三",
    style=FontStyle(font_name="宋体", font_size=14, bold=True)
)

# 填充学习经历（网格填充）
education_data = [
    ["2008.09-2012.06", "北京大学 计算机科学与技术", "李教授"],
    ["2012.09-2015.06", "清华大学 软件工程", "王导师"],
    ["", "", ""],
    ["", "", ""]
]
fill_grid(doc, education_data, position=(1, 1, 6, 1))
```

### 3. 批量生成简历

使用 `copy_doc()` 复制模板，为每个人生成独立的简历：

```python
from docxlib import copy_doc

template = load_docx("fixtures/templates/sample.docx")

for person in people_list:
    # 创建独立副本
    doc = copy_doc(template)

    # 填充数据
    fill_text(doc, (1, 1, 1, 2), person["姓名"])

    # 保存
    save_docx(doc, f"output/个人简历_{person['姓名']}.docx")
```

## 完整示例

运行示例代码：

```bash
python examples/resume_example.py
```

该示例包含：
1. 读取现有模板内容
2. 填充单份简历
3. 批量生成多份简历

## 生成的文件

示例运行后会在 `output/` 目录生成以下文件：

- `个人简历_张三.docx` (18KB) - 完整的示例简历
- `个人简历_李四.docx` (18KB) - 批量生成示例1
- `个人简历_王五.docx` (18KB) - 批量生成示例2

## API 对称性

读取和填充 API 形成完美的对称关系：

```python
# 填充数据
fill_text(doc, (1, 1, 1, 2), "张三")

# 读取数据
name = read_text(doc, (1, 1, 1, 2))
# 返回: "张三"

# 批量填充
fill_grid(doc, data, position=(1, 1, 6, 1))

# 批量读取
data = read_grid(doc, position=(1, 1, 6, 1))
```

## 高级功能

### 提取图片信息

```python
from docxlib import read_images

images = read_images(doc, include_data=True)
for img in images:
    print(f"位置: {img['position']}")
    print(f"尺寸: {img['width']} x {img['height']} 磅")
    # 保存图片
    with open(f"photo.{img['format']}", "wb") as f:
        f.write(img["data"])
```

### 填充照片

```python
from docxlib import fill_image, ImageConfig

fill_image(
    doc, (1, 1, 1, 1),  # 照片位置
    "fixtures/images/photo.jpg",
    config=ImageConfig(width=80, height=100)
)
```

## 实际应用场景

1. **HR 系统**：批量生成员工简历
2. **招聘网站**：导出简历为 Word 格式
3. **人事档案**：统一格式管理简历信息
4. **简历导入**：从 Word 简历中提取结构化数据

## 注意事项

1. 所有索引从 1 开始（不是 0）
2. 使用 `copy_doc()` 避免修改模板对象
3. 确保输出目录存在：`os.makedirs("output", exist_ok=True)`
4. Windows 控制台可能有编码问题，建议在 IDE 中运行

## 相关文件

- [examples/resume_example.py](resume_example.py) - 完整示例代码
- [docxlib/read.py](../docxlib/read.py) - 读取模块
- [docxlib/fill.py](../docxlib/fill.py) - 填充模块
- [fixtures/templates/sample.docx](../fixtures/templates/sample.docx) - 简历模板
