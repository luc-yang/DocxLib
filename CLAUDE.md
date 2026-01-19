# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

DocxLib is a Python library for Word document manipulation based on the Spire.Doc engine. It provides a functional API (no class instantiation needed) for document operations, table navigation, field filling, and styling.

**Key Design**: All functions accept a Spire.Doc `Document` object as the first parameter and operate on it directly. This allows chaining operations without managing complex object hierarchies.

## Development Commands

```bash
# Installation
make install              # Install package
make install-dev          # Install with dev dependencies

# Testing
make test                 # Run basic tests (tests/test_basic.py)
pytest tests/ -v          # Run all tests with verbose output
pytest tests/test_fill.py # Run specific test file
pytest tests/ --cov=docxlib --cov-report=html  # Run with coverage

# Code Quality
make format               # Format code with black
make lint                 # Check code with flake8

# Build
make build                # Build package for distribution

# CLI Tools
python -m docxlib.cli validate fixtures/templates/sample.docx
python -m docxlib.cli inspect fixtures/templates/sample.docx
python -m docxlib.cli extract-vars template.docx -o vars.json
python -m docxlib.cli fill template.docx data.json -o output.docx
python -m docxlib.cli convert input.docx -f pdf -o output.pdf
python -m docxlib.cli info          # Show library info
python -m docxlib.cli version       # Show version
```

## Architecture

### Module Structure

```
docxlib/
├── __init__.py      # Public API exports (functional interface)
├── cli.py           # Command-line interface (validate, inspect, fill, convert)
├── config.py        # Configuration classes (Style, Alignment, Options, ImageConfig)
├── document.py      # Document I/O, merge, format conversion
├── table.py         # Cell navigation, lookup, iteration
├── fill.py          # Field filling (text, image, date, grid, template vars)
├── read.py          # Read operations (text, grid, table, structure) - mirrors fill.py
├── style.py         # Font, color, alignment, borders
├── utils.py         # Validation, parsing, utilities
├── constants.py     # Default values, enums, type aliases
└── errors.py        # Exception hierarchy
```

### Positioning System (Critical)

All positional indices are **1-based**, not 0-based:
- Section index: which section in the document
- Table index: which table within the section
- Row index: which row in the table
- Column index: which column in the row

Position tuple format: `(section, table, row, col)`

Example: `(1, 1, 2, 2)` = Section 1, Table 1, Row 2, Column 2

Use `0` as a wildcard in `get_cells()` to select all:
- `get_cells(doc, 0, 0, 0, 0)` - all cells in all sections/tables/rows/cols
- `get_cells(doc, 1, 0, 2, 0)` - all cells in section 1, all tables, row 2, all columns

**Merged Cells**: The library handles merged cells correctly:
- `get_cell_text()` returns the combined text from merged cells
- `get_table_dimensions()` returns actual table dimensions (may differ from visual appearance due to merging)
- Position-based operations work with merged cells using the top-left cell position

### Fill Modes

The `fill_text()` and `fill_image()` functions support three modes via `Options`:

1. **`"position"`** (default): Direct position tuple
   ```python
   fill_text(doc, (1, 1, 2, 2), "content")
   ```

2. **`"match_right"`**: Find text pattern, fill in the cell to the right
   ```python
   fill_text(doc, "姓名：", "张三", options=Options.match_right())
   ```

3. **`"match_down"`**: Find text pattern, fill in the cell below
   ```python
   fill_text(doc, "项目1", "智慧城市", options=Options.match_down())
   ```

**Image Filling** uses similar modes with `ImageConfig`:

```python
from docxlib import ImageConfig

# Direct position
fill_image(doc, (1, 1, 2, 2), "photo.jpg")

# Match right with image config
fill_image(doc, "照片：", "photo.jpg",
           options=Options.match_right(),
           config=ImageConfig(width=80, height=80))

# Centered image with preset
fill_image(doc, "印章", "seal.png",
           options=Options.match_right(),
           config=ImageConfig.centered(width=100, height=100))

# Fixed size (no aspect ratio)
fill_image(doc, (1, 1, 1, 1), "logo.png",
           config=ImageConfig.fixed_size(200, 100))
```

### Match Mode Parameter

When using `match_right` or `match_down` modes, the `match_mode` parameter controls behavior when multiple matches are found:

- **`"all"`** (default): Fill all matching positions
- **`"first"`**: Only fill the first match

```python
# Fill all occurrences
fill_text(doc, "姓名：", "张三", options=Options.match_right(match_mode="all"))

# Fill only the first occurrence
fill_text(doc, "姓名：", "张三", options=Options.match_right(match_mode="first"))
```

### Text Normalization for Matching

**Important**: By default, `fill_text()`, `fill_image()`, and `fill_date()` use **text normalization** when matching text to handle template formatting issues. This means:

- `fill_text(doc, "姓名", "张三", options=Options.match_right())` will match:
  - `姓名` (exact)
  - `姓    名` (spaces between characters)
  - `姓\n名` (newline between characters)
  - `姓\t名` (tab between characters)
  - `姓　名` (full-width space)

**User Control**: You can control this behavior via the `normalize` parameter:

```python
# Using fill_text with Options
fill_text(doc, "姓名", "张三",
          options=Options.match_right())  # Default: normalize=True
fill_text(doc, "姓名", "张三",
          options=Options.match_right(normalize=False))  # Exact match

# Using fill_image
fill_image(doc, "照片", "photo.jpg",
           options=Options.match_right())  # Default: normalize=True
fill_image(doc, "照片", "photo.jpg",
           options=Options.match_right(normalize=False))  # Exact match

# Using fill_date
fill_date(doc, "日期", "2024年1月15日")  # Default: normalize=True
fill_date(doc, "日期", "2024年1月15日", normalize=False)  # Exact match

# Using find_text directly
positions = find_text(doc, "姓名")  # Default: normalize=True
positions = find_text(doc, "姓名", normalize=False)  # Exact match
```

**Recommendation**: Keep normalization enabled (`normalize=True`) for most use cases. Only disable it if you need exact text matching. Template designers often add spaces or newlines for visual formatting, and normalization ensures `fill_text()` works reliably.

### Date Filling

`fill_date()` handles dates with special formatting - the numeric part and "年月日" text use different fonts (common in Chinese documents):

```python
from docxlib import fill_date

# Direct position - numbers use style.font_name, "年月日" uses 宋体
fill_date(doc, (1, 1, 4, 2), "2024年1月15日")

# Match mode (uses text normalization by default)
fill_date(doc, "日期：", "2024年3月20日")

# With custom style and alignment
fill_date(doc, "签订日期", "2024年6月30日",
          style=Style(font_family="Arial", font_size=11),
          alignment=Alignment.center())

# Exact match (disable normalization)
fill_date(doc, "日期", "2024年1月15日", normalize=False)
```

**Note**: Unlike `fill_text()`, `fill_date()` accepts `normalize` directly, not via `options`.

### Template Variable System

DocxLib supports a template variable system for declarative document filling:

**Variable syntax**: `${variable_name}` or `${variable_name|default_value}`

```python
from docxlib import load_docx, fill_template, extract_template_vars

# Extract variables from a template
doc = load_docx("template.docx")
variables = extract_template_vars(doc, unique=True)
# Returns: ["name", "age", "department"]

# Fill template with data
data = {
    "name": "张三",
    "age": "25",
    "department": "研发部"
}
result = fill_template(doc, data, missing_var_action="ignore")
# Returns: {"total": 3, "replaced": 3, "missing": []}

# Validate template data before filling
from docxlib import validate_template_data
validation = validate_template_data(doc, data)
# Returns: {"is_valid": true, "required_vars": [...], "missing_vars": []}
```

**Missing variable actions**:
- `"error"` (default): Raise `VariableNotFoundError`
- `"ignore"`: Skip missing variables
- `"empty"`: Replace with empty string

### Styling System

Styles use configuration classes (`Style`, `Alignment`, `ImageConfig`) for a cleaner API:

```python
from docxlib import Style, Alignment, ImageConfig

# Via Style and Alignment objects
fill_text(doc, (1,1,2,2), "text",
          style=Style(font_family="黑体", font_size=16, bold=True, color="red"),
          alignment=Alignment.center())

# Using preset styles
fill_text(doc, "标题", "内容", style=Style.title())
fill_text(doc, "章节", "第一章", style=Style.heading(level=2))
fill_text(doc, "正文", "内容", style=Style.body())
fill_text(doc, "注意", "重要", style=Style.emphasis())

# Via style functions (for direct cell manipulation)
from docxlib import apply_font_style, apply_cell_alignment
cell = get_cell(doc, 1, 1, 2, 2)
apply_font_style(cell, font_name="黑体", font_size=16, bold=True, color="red")
apply_cell_alignment(cell, "center")
```

Supported color formats in `parse_color()`:
- Named colors: black, red, blue, green, yellow, white, gray, silver, maroon, purple, orange, pink
- Hex: `#RRGGBB`

### Document Copying for Batch Processing

When generating multiple documents from a template, always copy the loaded template:

```python
template = load_docx("template.docx")
for item in data:
    doc = copy_doc(template)  # Creates independent copy
    fill_text(doc, "name:", item["name"], options=Options.match_right())
    save_docx(doc, f"output_{item['id']}.docx")
```

### Read Functions (Symmetric to Fill)

The `read` module provides read functionality that mirrors the `fill` API:

```python
from docxlib import read_text, read_grid, read_table, read_cells

# Read text from position or by match
name = read_text(doc, (1, 1, 2, 2))
name = read_text(doc, "姓名：", default="未知")  # Read right of matched text
value = read_text(doc, "项目", default="N/A", options=Options.match_down())

# Read grid data as 2D array
data = read_grid(doc, position=(1, 1, 7, 1), rows=4, cols=3)

# Read entire table
table_data = read_table(doc, section=1, table=1)

# Read multiple cells
texts = read_cells(doc, (1, 0, 2, 0))  # All cells in row 2, section 1

# Read document structure
structure = read_document_structure(doc)
# Returns sections, tables, dimensions info

# Extract all template variables
vars_dict = extract_template_data(doc)
# Returns {"var_name": "default_value"} dict
```

## Platform Considerations

- **Primary support**: Windows 10/11
- **Limited support**: Domestic Linux distributions (NeoKylin, NFS China)
- **Community support**: Ubuntu, Debian, CentOS (may require testing)

Spire.Doc is a .NET-based library with COM interop on Windows. Linux support varies.

## Spire.Doc Free Version Limitations

Be aware of these constraints when designing features:
- Maximum 500 paragraphs per document
- Maximum 25 tables per document
- PDF conversion includes watermark
- Non-commercial use only

## Import Patterns

The package exposes all public APIs via `docxlib/__init__.py`. Users should import from the top level:

```python
from docxlib import load_docx, fill_text, save_docx
```

Internal modules (e.g., `docxlib.document`) are not part of the public API.

## Testing Notes

- Test fixtures located in `fixtures/templates/`
- Test images in `fixtures/images/`
- Output files go to `output/` directory (auto-created by `save_docx()`)
- Use `copy_doc()` in tests to avoid modifying shared template objects
- Run specific test with: `pytest tests/test_fill.py -v -k test_fill_text`
- Run tests with coverage: `pytest tests/ --cov=docxlib --cov-report=html`
- Test files: test_basic.py, test_fill.py, test_template.py, test_document.py, test_table.py, test_read.py, test_read_edge_cases.py, test_read_module.py, test_utils.py, test_cli.py