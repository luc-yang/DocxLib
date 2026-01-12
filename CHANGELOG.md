# DocxLib 变更日志

## [0.1.0] - 2024-01-15

### 新增功能

#### 核心模块
- `errors.py` - 异常类定义
  - `DocxLibError` - 基础异常类
  - `DocumentError` - 文档操作错误
  - `PositionError` - 位置定位错误
  - `FillError` - 字段填充错误
  - `ValidationError` - 数据验证错误

- `constants.py` - 常量定义
  - 默认字体、字号、颜色
  - 支持的图片格式
  - 文件格式常量
  - 填充模式常量

- `utils.py` - 工具函数
  - `is_valid_docx()` - 验证 DOCX 格式
  - `parse_csv()` - 解析 CSV 文件
  - `parse_json()` - 解析 JSON 文件
  - `ensure_directory()` - 确保目录存在
  - `parse_date_string()` - 解析日期字符串

- `style.py` - 样式管理
  - `parse_color()` - 解析颜色（支持名称和十六进制）
  - `apply_font_style()` - 应用字体样式
  - `set_cell_border()` - 设置单元格边框

#### 文档操作 (document.py)
- `load_docx()` - 加载文档（支持文件路径或字节数据）
- `save_docx()` - 保存文档（自动创建目录）
- `merge_docs()` - 合并多个文档
- `to_pdf()` - 转换为 PDF（返回字节数据）
- `to_images()` - 转换为图片列表
- `to_pdf_file()` - 转换为 PDF 文件
- `copy_doc()` - 复制文档（用于批量处理）

#### 表格操作 (table.py)
- `get_cell()` - 获取指定位置的单元格（1-based 索引）
- `get_cells()` - 通配符获取单元格（0表示所有）
- `find_text()` - 查找包含指定文本的所有单元格
- `iterate_cells()` - 遍历所有单元格（生成器）
- `get_cell_text()` - 获取单元格文本内容
- `get_table_dimensions()` - 获取表格行数和列数
- `get_section_table_count()` - 获取节中的表格数量
- `get_section_count()` - 获取文档中的节数量

#### 字段填充 (fill.py)
- `fill_text()` - 填充文本
  - 支持三种模式：`position`、`match_right`、`match_down`
  - 支持自定义字体、字号、颜色、粗体、斜体、下划线

- `fill_image()` - 填充图片
  - 支持三种填充模式
  - 支持设置宽高
  - 自动保持宽高比

- `fill_date()` - 填充日期
  - 数字和年月日使用不同字体
  - 支持三种填充模式

- `fill_grid()` - 填充网格数据
  - 从二维数组填充数据到表格

- `replace_all()` - 全局替换文本

- `clear_cell()` - 清空单元格内容

### 示例代码

- `examples/basic_usage.py` - 基础用法示例
- `examples/batch_processing.py` - 批量处理示例
- `examples/advanced_features.py` - 高级功能示例

### 测试

- `tests/test_basic.py` - 基础功能测试
- `tests/test_document.py` - 文档操作测试
- `tests/test_fill.py` - 字段填充测试
- `tests/test_table.py` - 表格操作测试

### 文档

- `README.md` - 项目说明和使用指南
- `INSTALL.md` - 安装和测试指南
- `CHANGELOG.md` - 变更日志

### 已修复问题

- 修复 `fill_image()` 中图片加载的问题
- 修复 `parse_color()` 中 Color 对象创建问题（使用 `Color.FromArgb(255, r, g, b)`）
- 修复测试脚本中的编码问题（使用 ASCII 字符替代 Unicode 符号）

## 未来计划

### [0.2.0] - 计划中

- 添加更多文档格式支持（RTF、HTML 等）
- 添加段落和文本操作功能
- 添加表格创建和编辑功能
- 改进错误处理和提示信息
- 添加更多单元测试

### [0.3.0] - 计划中

- 添加批量处理优化（多进程支持）
- 添加模板变量替换功能
- 添加条件格式支持
- 添加数据验证功能

### [1.0.0] - 计划中

- 完整的 API 文档
- 性能优化
- 更多示例代码
- 生产环境测试

## 已知限制

1. **Spire.Doc 免费版限制**
   - 最多 500 段落
   - 最多 25 个表格
   - PDF 转换有水印

2. **平台支持**
   - 主要支持 Windows
   - macOS/Linux 有限支持

3. **功能限制**
   - 不支持 VBA 宏处理
   - 不支持复杂文档结构编辑（如插入节、页眉页脚等）
