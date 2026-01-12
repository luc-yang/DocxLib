"""
DocxLib - Word 文档处理库
"""

from pathlib import Path
from setuptools import setup, find_packages

# 读取 README
readme_file = Path(__file__).parent / "README.md"
long_description = readme_file.read_text(encoding="utf-8") if readme_file.exists() else ""

setup(
    name="docxlib",
    use_scm_version=False,
    version="0.1.0",
    description="A simple and easy-to-use Word document processing library based on Spire.Doc",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="DocxLib Contributors",
    python_requires=">=3.8",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        "spire-doc-free>=12.12.0",
    ],
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "pytest-cov>=3.0.0",
        ],
        "excel": [
            "pandas>=1.3.0",
            "openpyxl>=3.0.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "docxlib=docxlib.cli:main",
        ],
    },
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Office/Business",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
    keywords="word docx document spire office",
    project_urls={
        "Homepage": "https://github.com/yourusername/docxlib",
        "Repository": "https://github.com/yourusername/docxlib",
        "Documentation": "https://github.com/yourusername/docxlib/blob/main/README.md",
        "Bug Tracker": "https://github.com/yourusername/docxlib/issues",
    },
)
