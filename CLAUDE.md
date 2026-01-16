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
├── document.py      # Document I/O, merge, format conversion
├── table.py         # Cell navigation, lookup, iteration
├── fill.py          # Field filling (text, image, date, grid, template vars)
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

### Fill Modes

The `fill_text()` and `fill_image()` functions support three modes via the `mode` parameter:

1. **`"position"`** (default): Direct position tuple
   ```python
   fill_text(doc, (1, 1, 2, 2), "content")
   ```

2. **`"match_right"`**: Find text pattern, fill in the cell to the right
   ```python
   fill_text(doc, "姓名：", "张三", mode="match_right")
   ```

3. **`"match_down"`**: Find text pattern, fill in the cell below
   ```python
   fill_text(doc, "项目1", "智慧城市", mode="match_down")
   ```

### Match Mode Parameter

When using `match_right` or `match_down` modes, the `match_mode` parameter controls behavior when multiple matches are found:

- **`"all"`** (default): Fill all matching positions
- **`"first"`**: Only fill the first match

```python
# Fill all occurrences
fill_text(doc, "姓名：", "张三", mode="match_right", match_mode="all")

# Fill only the first occurrence
fill_text(doc, "姓名：", "张三", mode="match_right", match_mode="first")
```

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

Styles can be applied directly in fill functions or separately:

```python
# Via fill parameters
fill_text(doc, (1,1,2,2), "text", font_name="黑体", font_size=16, bold=True, color="red")

# Via style functions
from docxlib import apply_font_style, apply_cell_alignment
cell = get_cell(doc, 1, 1, 2, 2)
apply_font_style(cell, font_name="黑体", font_size=16, bold=True, color="red")
apply_cell_alignment(cell, "center")
```

Supported color formats in `parse_color()`:
- Named colors: black, red, blue, green, yellow, white, gray, cyan, magenta, orange, purple, brown
- Hex: `#RRGGBB`

### Document Copying for Batch Processing

When generating multiple documents from a template, always copy the loaded template:

```python
template = load_docx("template.docx")
for item in data:
    doc = copy_doc(template)  # Creates independent copy
    fill_text(doc, "name:", item["name"], mode="match_right")
    save_docx(doc, f"output_{item['id']}.docx")
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
- Test files: test_basic.py, test_fill.py, test_template.py, test_document.py, etc.