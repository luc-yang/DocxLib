# 模板变量语法功能

## 概述

DocxLib 现在支持使用 `${variable}` 格式的模板变量，可以批量填充 Word 文档模板。

## 核心功能

### 1. 批量变量填充

```python
from docxlib import load_docx, fill_template, save_docx

# 加载模板
doc = load_docx("template.docx")

# 批量填充变量
fill_template(doc, {
    "name": "张三",
    "age": "25",
    "dept": "研发部",
    "date": "2024年1月15日"
})

# 保存文档
save_docx(doc, "output.docx")
```

### 2. 提取模板变量

```python
from docxlib import load_docx, extract_template_vars

doc = load_docx("template.docx")

# 提取所有变量
vars = extract_template_vars(doc)
print(vars)  # ['name', 'age', 'dept', 'date']
```

### 3. 验证数据完整性

```python
from docxlib import load_docx, validate_template_data

doc = load_docx("template.docx")

# 验证数据
result = validate_template_data(doc, {
    "name": "张三",
    "age": "25"
})

if not result["is_valid"]:
    print(f"缺失变量: {result['missing_vars']}")
```

### 4. 默认值支持

模板中使用 `${变量|默认值}` 语法：

```python
# 模板中：${dept|未知}
# 数据中不提供 dept 变量
fill_template(doc, {"name": "张三"})
# 结果：dept 会被填充为 "未知"
```

### 5. 灵活的错误处理

```python
# 抛出异常（默认）
fill_template(doc, data, missing_var_action="error")

# 忽略缺失变量
fill_template(doc, data, missing_var_action="ignore")

# 替换为空字符串
fill_template(doc, data, missing_var_action="empty")
```

## 变量语法规则

- **格式**: `${variable_name}`
- **命名规则**:
  - 必须以字母或下划线开头
  - 只能包含字母、数字和下划线
  - 示例：`${name}`, `${user_age}`, `${date}`
- **默认值**: `${variable|default_value}`

## 使用示例

查看 `examples/template_usage_example.py` 了解完整的使用示例，包括：

- 基础用法
- 变量提取
- 数据验证
- 默认值使用
- 批量文档生成
- 合同生成

## 运行示例

```bash
python examples/template_usage_example.py
```

## 测试

运行测试：

```bash
pytest tests/test_template.py -v
```

## 向后兼容性

此功能完全向后兼容，不影响现有代码：

```python
# 旧代码仍然有效
from docxlib import replace_all
replace_all(doc, "{name}", "张三")

# 新功能提供更好的体验
from docxlib import fill_template
fill_template(doc, {"name": "张三"})
```

## API 参考

### fill_template()

批量替换模板变量。

**参数**:
- `doc`: Document 对象
- `data`: 变量数据字典
- `missing_var_action`: 缺失变量处理方式 ("error" | "ignore" | "empty")
- `placeholder_prefix`: 变量前缀（默认 "${"）
- `placeholder_suffix`: 变量后缀（默认 "}"）
- 以及可选的样式参数（font_name, font_size, color, bold, italic, underline, h_align, v_align）

**返回**:
- 统计信息字典: `{"total": int, "replaced": int, "missing": list, "errors": list}`

### extract_template_vars()

提取模板中的所有变量。

**参数**:
- `doc`: Document 对象
- `unique`: 是否返回唯一值（默认 True）
- `placeholder_prefix`: 变量前缀（默认 "${"）
- `placeholder_suffix`: 变量后缀（默认 "}"）

**返回**:
- 变量名列表（不包含前后缀）

### validate_template_data()

验证数据完整性。

**参数**:
- `doc`: Document 对象
- `data`: 变量数据字典
- `placeholder_prefix`: 变量前缀（默认 "${"）
- `placeholder_suffix`: 变量后缀（默认 "}"）

**返回**:
- 验证结果字典:
  - `"is_valid"`: bool
  - `"required_vars"`: list
  - `"missing_vars"`: list
  - `"extra_vars"`: list

## 测试模板

项目包含两个测试模板：

1. `fixtures/templates/template_vars.docx` - 基础变量测试模板
2. `fixtures/templates/contract_template.docx` - 合同模板示例

可以使用以下脚本重新生成测试模板：

```bash
python scripts/create_test_template.py
```

## 注意事项

1. **批量生成文档时**：建议每次重新加载模板，而不是使用 `copy_doc()`
2. **样式保持**：当前实现使用简单替换，可能会丢失部分格式
3. **性能**：对于包含大量变量的文档，性能可能会受到影响
4. **Spire.Doc 限制**：免费版有段落和表格数量限制

## 未来改进

以下功能计划在未来版本中实现：

- 变量级别样式支持
- 嵌套变量访问 `${user.name}`
- 条件渲染
- 循环渲染（表格）
- 更好的样式保持

## 相关文件

- 实现: [docxlib/fill.py](docxlib/fill.py)
- 异常: [docxlib/errors.py](docxlib/errors.py)
- 常量: [docxlib/constants.py](docxlib/constants.py)
- 测试: [tests/test_template.py](tests/test_template.py)
- 示例: [examples/template_usage_example.py](examples/template_usage_example.py)
