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
```

## Architecture

### Module Structure

```
docxlib/
├── __init__.py      # Public API exports (functional interface)
├── document.py      # Document I/O, merge, format conversion
├── table.py         # Cell navigation, lookup, iteration
├── fill.py          # Field filling (text, image, date, grid)
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