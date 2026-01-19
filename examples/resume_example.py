"""
简历模板读取和填充示例

演示如何使用 read.py 和 fill.py 模块来读取和填充个人简历模板。

模板结构（12行 x 7列）：
- 第1行: 姓名、性别、出生年月、民族、籍贯、政治面貌、身体状况
- 第2行: 入党（团）时间、参加工作时间
- 第3行: 专业特长
- 第4行: 专业职称
- 第5行: 获得职称时间
- 第6-9行: 学习经历（起止时间、在何院校学习专业、证明人）
- 第10行: 工作经历
- 第11行: 学习工作成果
- 第12行: 奖惩情况
"""

from docxlib import (
    load_docx,
    save_docx,
    copy_doc,
    # 读取函数
    read_text,
    read_grid,
    read_cells,
    read_images,
    # 填充函数
    fill_text,
    fill_grid,
    fill_image,
    FillOptions,
    FontStyle,
    Alignment,
)
from docxlib.config import FillOptions
import os


def read_resume(template_path):
    """读取简历模板内容

    Args:
        template_path: 模板文件路径

    Returns:
        dict: 简历信息字典
    """
    print("=" * 60)
    print("读取简历模板")
    print("=" * 60)

    doc = load_docx(template_path)

    # 读取基本信息（使用位置元组）
    print("\n【基本信息】")
    info = {}

    # 读取第1行的基本信息
    # 姓名
    info["姓名"] = read_text(doc, (1, 1, 1, 2), default="")
    # 性别
    info["性别"] = read_text(doc, (1, 1, 1, 4), default="")
    # 出生年月
    info["出生年月"] = read_text(doc, (1, 1, 1, 6), default="")

    print(f"姓名: {info['姓名']}")
    print(f"性别: {info['性别']}")
    print(f"出生年月: {info['出生年月']}")

    # 读取第2行的时间信息
    print("\n【时间信息】")
    # 入党（团）时间
    info["入党团时间"] = read_text(doc, (1, 1, 2, 2), default="")
    # 参加工作时间
    info["工作时间"] = read_text(doc, (1, 1, 2, 4), default="")

    print(f"入党（团）时间: {info['入党团时间']}")
    print(f"参加工作时间: {info['工作时间']}")

    # 读取专业信息
    print("\n【专业信息】")
    # 专业特长
    info["专业特长"] = read_text(doc, (1, 1, 3, 2), default="")
    # 专业职称
    info["专业职称"] = read_text(doc, (1, 1, 4, 2), default="")
    # 获得职称时间
    info["职称时间"] = read_text(doc, (1, 1, 5, 2), default="")

    print(f"专业特长: {info['专业特长']}")
    print(f"专业职称: {info['专业职称']}")
    print(f"获得职称时间: {info['职称时间']}")

    # 批量读取学习经历（第6-9行，共4行，每行3列）
    print("\n【学习经历】")
    education_positions = []
    for row in range(6, 10):  # 第6-9行
        for col in range(1, 4):  # 第1-3列
            education_positions.append((1, 1, row, col))

    education_values = read_cells(doc, education_positions, default="")

    # 将学习经历格式化为列表
    info["学习经历"] = []
    for i in range(0, len(education_values), 3):
        if i + 2 < len(education_values):
            start_end = education_values[i]
            school = education_values[i + 1]
            witness = education_values[i + 2]

            if start_end or school or witness:
                info["学习经历"].append({
                    "起止时间": start_end,
                    "在何院校": school,
                    "证明人": witness
                })

    for idx, edu in enumerate(info["学习经历"], 1):
        print(f"  {idx}. {edu['起止时间']} - {edu['在何院校']} (证明人: {edu['证明人']})")

    # 检查图片（照片）
    print("\n【图片信息】")
    images = read_images(doc)
    print(f"照片数量: {len(images)}")
    for img in images:
        print(f"  位置: 节{img['position'][0]}, 表格{img['position'][1]}, 行{img['position'][2]}, 列{img['position'][3]}")

    return info


def fill_resume(template_path, output_path, resume_data):
    """填充简历数据

    Args:
        template_path: 模板文件路径
        output_path: 输出文件路径
        resume_data: 简历数据字典
    """
    print("\n" + "=" * 60)
    print("填充简历数据")
    print("=" * 60)

    # 加载模板
    doc = load_docx(template_path)

    # 填充第1行：基本信息
    print("\n填充基本信息...")
    fill_text(
        doc, (1, 1, 1, 2),
        resume_data.get("姓名", ""),
        style=FontStyle(font_family="宋体", font_size=14, bold=True)
    )
    fill_text(doc, (1, 1, 1, 4), resume_data.get("性别", ""))
    fill_text(doc, (1, 1, 1, 6), resume_data.get("出生年月", ""))

    # 填充第2行：时间信息
    print("填充时间信息...")
    fill_text(doc, (1, 1, 2, 2), resume_data.get("入党团时间", ""))
    fill_text(doc, (1, 1, 2, 4), resume_data.get("工作时间", ""))

    # 填充专业信息
    print("填充专业信息...")
    fill_text(doc, (1, 1, 3, 2), resume_data.get("专业特长", ""))
    fill_text(doc, (1, 1, 4, 2), resume_data.get("专业职称", ""))
    fill_text(doc, (1, 1, 5, 2), resume_data.get("职称时间", ""))

    # 填充学习经历（使用网格填充）
    print("填充学习经历...")
    education = resume_data.get("学习经历", [])
    education_grid = []

    for edu in education[:4]:  # 最多4行
        education_grid.append([
            edu.get("起止时间", ""),
            edu.get("在何院校", ""),
            edu.get("证明人", "")
        ])

    # 如果不足4行，添加空行
    while len(education_grid) < 4:
        education_grid.append(["", "", ""])

    # 填充学习经历网格
    if education_grid:
        fill_grid(doc, education_grid, position=(1, 1, 7, 2))

    # 填充其他信息（使用查找文本模式）
    print("填充其他信息...")

    # 工作经历（第10行，第2列开始）
    work_exp = resume_data.get("工作经历", "")
    options = FillOptions(mode="match_right")
    style = FontStyle(font_family="黑体", font_size=10.5)
    align = Alignment.left_top()
    fill_text(doc, "工作经历", work_exp, options=options, style=style, alignment=align)

    # 学习工作成果（第11行，第2列开始）
    achievements = resume_data.get("学习工作成果", "")
    fill_text(doc, "学习与工作成果", achievements, options=options, style=style, alignment=align) 

    # 奖惩情况（第12行，第2列开始）
    awards = resume_data.get("奖惩情况", "")
    fill_text(doc, "奖惩情况", awards, options=options, style=style, alignment=align)

    # 填充照片（如果有）
    photo_path = resume_data.get("照片路径")
    if photo_path and os.path.exists(photo_path):
        print("填充照片...")
        from docxlib import ImageConfig
        fill_image(
            doc, (1, 1, 1, 7),
            photo_path,
            config=ImageConfig(width=80, height=100)
        )

    # 保存文档
    print(f"\n保存简历到: {output_path}")
    save_docx(doc, output_path)
    print("[OK] 简历生成成功！")


def main():
    """主函数：演示读取和填充流程"""

    # 模板路径
    template_path = "fixtures/templates/sample.docx"

    # 示例1：读取现有模板
    print("\n" + "=" * 60)
    print("示例1：读取简历模板")
    print("=" * 60)
    resume_info = read_resume(template_path)
    print(resume_info)

    # 示例2：创建新简历
    print("\n" + "=" * 60)
    print("示例2：填充新简历数据")
    print("=" * 60)

    # # 准备简历数据
    new_resume_data = {
        "姓名": "张三",
        "性别": "男",
        "出生年月": "1990年5月",
        "入党团时间": "2012年6月",
        "工作时间": "2015年7月",
        "专业特长": "Java开发、Python开发、系统架构设计",
        "专业职称": "软件工程师",
        "职称时间": "2020年10月",
        "学习经历": [
            {
                "起止时间": "2008.09-2012.06",
                "在何院校": "北京大学 计算机科学与技术 本科",
                "证明人": "李教授"
            },
            {
                "起止时间": "2012.09-2015.06",
                "在何院校": "清华大学 软件工程 硕士",
                "证明人": "王导师"
            },
            
        ],
        "工作经历": "2015.07-2018.06  某某科技公司 软件工程师\n2018.07-至今  某某互联网公司 高级软件工程师",
        "学习工作成果": "主持开发了多个大型项目，发表了3篇技术论文，获得2项专利",
        "奖惩情况": "2021年度优秀员工",
        "照片路径": "fixtures/images/photo.png"
    }

    # # 填充简历
    output_path = "output/个人简历_张三.docx"
    fill_resume(template_path, output_path, new_resume_data)

if __name__ == "__main__":
    # 确保输出目录存在
    os.makedirs("output", exist_ok=True)

    # 运行示例
    main()
