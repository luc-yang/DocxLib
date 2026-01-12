"""
DocxLib 文档操作模块

提供文档的加载、保存、合并、格式转换等功能。
"""

import io
from pathlib import Path
from typing import Union, List

from spire.doc import *
from spire.doc.common import *
from spire.doc import FileFormat as SpireFileFormat

from .errors import DocumentError, ValidationError
from .utils import is_valid_docx, ensure_directory


def load_docx(source: Union[str, bytes, Path]) -> Document:
    """加载文档

    从文件路径或字节数据加载 Word 文档。

    Args:
        source: 文件路径（str/Path）或字节数据（bytes）

    Returns:
        Document: Spire.Doc Document 对象

    Raises:
        DocumentError: 文件不存在或格式错误
        ValidationError: 文件格式不是 .docx

    Examples:
        >>> doc = load_docx("template.docx")

        >>> # 从字节数据加载
        >>> with open("template.docx", "rb") as f:
        ...     data = f.read()
        >>> doc = load_docx(data)
    """
    doc = Document()

    if isinstance(source, (str, Path)):
        file_path = Path(source)

        # 检查文件是否存在
        if not file_path.exists():
            raise DocumentError(f"文件不存在: {source}")

        # 检查文件扩展名
        if file_path.suffix.lower() not in ['.docx', '.dotx']:
            raise ValidationError(f"文件格式不是 .docx: {source}")

        # 验证 DOCX 格式
        if not is_valid_docx(file_path):
            raise ValidationError(f"文件不是有效的 DOCX 格式: {source}")

        try:
            doc.LoadFromFile(str(file_path))
        except Exception as e:
            raise DocumentError(f"加载文档失败: {e}")

    elif isinstance(source, bytes):
        # 验证字节数据
        if not is_valid_docx(source):
            raise ValidationError("字节数据不是有效的 DOCX 格式")

        try:
            stream = io.BytesIO(source)
            doc.LoadFromStream(stream)
        except Exception as e:
            raise DocumentError(f"从字节数据加载文档失败: {e}")
    else:
        raise ValidationError(f"不支持的源类型: {type(source)}")

    return doc


def save_docx(doc: Document, target: Union[str, Path]) -> None:
    """保存文档

    将文档保存到指定路径。自动创建不存在的目录。

    Args:
        doc: Document 对象
        target: 保存路径

    Raises:
        DocumentError: 保存失败

    Examples:
        >>> save_docx(doc, "output.docx")

        >>> # 自动创建目录
        >>> save_docx(doc, "output/reports/report.docx")
    """
    target_path = Path(target)

    # 确保目录存在
    try:
        ensure_directory(target_path)
    except Exception as e:
        raise DocumentError(f"创建目录失败: {e}")

    # 保存文档
    try:
        doc.SaveToFile(str(target_path), SpireFileFormat.Docx)
    except Exception as e:
        raise DocumentError(f"保存文档失败: {e}")


def merge_docs(doc_list: List[Document]) -> Document:
    """合并多个文档

    将多个文档按顺序合并为一个新文档。

    Args:
        doc_list: Document 对象列表

    Returns:
        Document: 合并后的 Document 对象

    Raises:
        DocumentError: 合并失败

    Examples:
        >>> doc1 = load_docx("part1.docx")
        >>> doc2 = load_docx("part2.docx")
        >>> merged = merge_docs([doc1, doc2])
        >>> save_docx(merged, "combined.docx")
    """
    if not doc_list:
        raise DocumentError("文档列表不能为空")

    try:
        # 创建新文档
        merged_doc = Document()

        # 复制第一个文档的内容
        for i, doc in enumerate(doc_list):
            for section_idx in range(doc.Sections.Count):
                section = doc.Sections.get_Item(section_idx)

                # 为每个文档的节创建新节
                for sec_idx in range(section.Sections.Count):
                    src_section = section.Sections.get_Item(sec_idx)

                    # 克隆节
                    new_section = src_section.Clone()

                    # 添加到合并文档
                    merged_doc.Sections.Add(new_section)

        return merged_doc

    except Exception as e:
        raise DocumentError(f"合并文档失败: {e}")


def to_pdf(doc: Document) -> bytes:
    """将文档转换为 PDF

    Args:
        doc: Document 对象

    Returns:
        bytes: PDF 文件字节数据

    Raises:
        DocumentError: 转换失败

    Examples:
        >>> doc = load_docx("document.docx")
        >>> pdf_bytes = to_pdf(doc)
        >>> with open("output.pdf", "wb") as f:
        ...     f.write(pdf_bytes)

    Note:
        Spire.Doc 免费版转换的 PDF 会有水印
    """
    try:
        stream = Stream()
        doc.SaveToStream(stream, SpireFileFormat.PDF)
        return stream.ToArray()
    except Exception as e:
        raise DocumentError(f"转换为 PDF 失败: {e}")


def to_images(doc: Document) -> List[bytes]:
    """将文档转换为图片列表

    每一页转换为一张图片。

    Args:
        doc: Document 对象

    Returns:
        List[bytes]: 图片字节数据列表

    Raises:
        DocumentError: 转换失败

    Examples:
        >>> doc = load_docx("document.docx")
        >>> images = to_images(doc)
        >>> for i, img_bytes in enumerate(images):
        ...     with open(f"page_{i+1}.png", "wb") as f:
        ...         f.write(img_bytes)
    """
    try:
        images = []
        for page_index in range(doc.PageCount):
            image_stream = doc.SaveImageToStreams(
                page_index, ImageType.Bitmap
                )
            images.append(image_stream.ToArray())
        return images
    except Exception as e:
        raise DocumentError(f"转换为图片失败: {e}")


def to_pdf_file(doc: Document, file_path: Union[str, Path]) -> None:
    """将文档转换为 PDF 并保存到文件

    Args:
        doc: Document 对象
        file_path: 保存路径

    Raises:
        DocumentError: 转换失败

    Examples:
        >>> doc = load_docx("document.docx")
        >>> to_pdf_file(doc, "output.pdf")
    """
    target_path = Path(file_path)

    # 确保目录存在
    try:
        ensure_directory(target_path)
    except Exception as e:
        raise DocumentError(f"创建目录失败: {e}")

    # 转换并保存
    try:
        doc.SaveToFile(str(target_path), SpireFileFormat.PDF)
    except Exception as e:
        raise DocumentError(f"转换为 PDF 失败: {e}")


def copy_doc(doc: Document) -> Document:
    """复制文档

    创建文档的深拷贝，用于批量生成时复用模板。

    Args:
        doc: Document 对象

    Returns:
        Document: 文档副本

    Examples:
        >>> template = load_docx("template.docx")
        >>> for i in range(10):
        ...     doc = copy_doc(template)
        ...     # 修改文档...
        ...     save_docx(doc, f"output_{i}.docx")
    """
    import copy

    try:
        return copy.deepcopy(doc)
    except Exception as e:
        raise DocumentError(f"复制文档失败: {e}")
