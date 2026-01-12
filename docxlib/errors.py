"""
DocxLib 异常类定义

定义了库中使用的所有异常类，继承自基础异常类 DocxLibError。
"""


class DocxLibError(Exception):
    """DocxLib 基础异常类

    所有 DocxLib 异常的基类，用于捕获库相关的所有错误。
    """

    pass


class DocumentError(DocxLibError):
    """文档操作错误

    当文档加载、保存、转换等操作失败时抛出。

    使用场景：
        - 文件不存在或无法读取
        - 文件格式错误
        - 保存操作失败
    """

    pass


class PositionError(DocxLibError):
    """位置定位错误

    当表格位置、单元格位置越界或无效时抛出。

    使用场景：
        - 索引越界
        - 位置参数无效
        - 数据超出表格边界
    """

    pass


class FillError(DocxLibError):
    """字段填充错误

    当字段填充操作失败时抛出。

    使用场景：
        - 图片文件不存在或格式不支持
        - 填充操作失败
        - 无效的填充参数
    """

    pass


class ValidationError(DocxLibError):
    """数据验证错误

    当输入数据格式不正确或无法通过验证时抛出。

    使用场景：
        - 文件格式不是 .docx
        - 数据格式错误
        - 参数验证失败
    """

    pass
