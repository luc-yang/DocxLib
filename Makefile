.PHONY: help install install-dev test test-cov clean build upload docs

# 默认目标
help:
	@echo "DocxLib - Makefile 命令"
	@echo ""
	@echo "可用命令:"
	@echo "  make install       - 安装包"
	@echo "  make install-dev   - 安装包（含开发依赖）"
	@echo "  make test          - 运行测试"
	@echo "  make test-cov      - 运行测试并生成覆盖率报告"
	@echo "  make clean         - 清理临时文件"
	@echo "  make build         - 构建包"
	@echo "  make upload        - 上传到 PyPI"
	@echo "  make docs          - 生成文档"
	@echo "  make format        - 格式化代码"
	@echo "  make lint          - 代码检查"
	@echo "  make validate      - 验证 DOCX 文件"
	@echo "  make inspect       - 检查文档结构"

# 安装
install:
	pip install -e .

install-dev:
	pip install -e ".[dev]"

# 测试
test:
	python tests/test_basic.py

test-cov:
	pytest tests/ --cov=docxlib --cov-report=html --cov-report=term

test-all:
	pytest tests/ -v

# 清理
clean:
	rm -rf build/
	rm -rf dist/
	rm -rf *.egg-info
	rm -rf .pytest_cache/
	rm -rf htmlcov/
	rm -rf .coverage
	find . -type d -name __pycache__ -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete

# 构建
build: clean
	python -m build

# 上传到 PyPI
upload: build
	python -m twine upload dist/*

# 文档
docs:
	@echo "Generating documentation..."
	@echo "See README.md for documentation"

# 代码格式化
format:
	black docxlib/ examples/ tests/

# 代码检查
lint:
	flake8 docxlib/ examples/ tests/

# 验证文档
validate:
	python -m docxlib.cli validate fixtures/templates/sample.docx

# 检查文档
inspect:
	python -m docxlib.cli inspect fixtures/templates/sample.docx

# 显示信息
info:
	python -m docxlib.cli info

# 版本
version:
	python -m docxlib.cli version
