#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
学术论文自动化排版工具 - 核心处理脚本
======================================

功能：读取未经排版的 .docx 文档，按照学术论文排版规范对其进行格式重构。

排版规则：
  1. 正文：宋体 + Times New Roman（英文/数字），小四号字，首行缩进2字符，1.5倍行距
  2. 摘要/关键词：识别并加粗标签
  3. 一级标题（如 "1 引言"）：黑体，三号，加粗，居中，段前段后1行
  4. 二级标题（如 "1.1 研究背景"）：黑体，四号，加粗，左对齐，段前段后0.5行
  5. 图表标题（如 "表 1 变量定义"）：黑体，五号，居中，无缩进

依赖安装：
  pip install python-docx

使用方法：
  python format_paper.py input.docx output.docx
"""

import re
import sys
import logging
from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ============================================================
# 日志配置
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

# ============================================================
# 字号映射表（中国标准字号 → 磅值）
# ============================================================
FONT_SIZE_MAP = {
    "初号": 42,
    "小初": 36,
    "一号": 26,
    "小一": 24,
    "二号": 22,
    "小二": 18,
    "三号": 16,
    "小三": 15,
    "四号": 14,
    "小四": 12,
    "五号": 10.5,
    "小五": 9,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5,
}

# ============================================================
# 正则表达式定义（核心匹配逻辑）
# ============================================================

# --- 一级标题匹配 ---
# 匹配规则：以 1-9 开头的数字 + 一个或多个空格 + 至少一个中文字符
# 示例匹配："1 引言"、"2 研究设计"、"3 模型的估计与检验"
# 要求该段落仅包含这一行内容（独占一行），因此使用 ^ 和 $ 锚定
RE_HEADING_L1 = re.compile(
    r"^\d+\s+[\u4e00-\u9fff][\u4e00-\u9fff\w\s\-—、（）()]*$"
)

# --- 二级标题匹配 ---
# 匹配规则：数字.数字 + 可选空格 + 至少一个中文字符
# 示例匹配："1.1研究背景"、"2.1 模型构建"、"3.2 数据来源与描述"
RE_HEADING_L2 = re.compile(
    r"^\d+\.\d+\s*[\u4e00-\u9fff][\u4e00-\u9fff\w\s\-—、（）()]*$"
)

# --- 图表标题匹配 ---
# 匹配规则：以 "图" 或 "表" 开头 + 可选空格 + 数字 + 后续内容（较短段落）
# 示例匹配："表 1 变量定义"、"图 1 散点图"、"表1 回归结果"
# 限制总长度不超过 40 个字符，以避免误匹配正文段落
RE_FIGURE_TABLE = re.compile(
    r"^[图表]\s*\d+[\s\.\-—:：]*.{0,35}$"
)

# --- 摘要标识匹配 ---
# 匹配规则：段落起始处包含 "摘要" + 可选的标点符号（如 ":"、"："）
RE_ABSTRACT = re.compile(r"^摘\s*要\s*[:：]?\s*")

# --- 关键词标识匹配 ---
# 匹配规则：段落起始处包含 "关键词" + 可选的标点符号
RE_KEYWORDS = re.compile(r"^关\s*键\s*词\s*[:：]?\s*")


# ============================================================
# 段落分类枚举
# ============================================================
class ParagraphType:
    """段落类型常量"""
    HEADING_L1 = "heading_l1"        # 一级标题
    HEADING_L2 = "heading_l2"        # 二级标题
    FIGURE_TABLE = "figure_table"    # 图表标题
    ABSTRACT = "abstract"            # 摘要段落
    KEYWORDS = "keywords"            # 关键词段落
    BODY = "body"                    # 正文段落


# ============================================================
# 段落分类函数
# ============================================================
def classify_paragraph(text: str) -> str:
    """
    根据段落纯文本内容判断其类型。

    分类优先级（从高到低）：
      1. 摘要 → 包含"摘要："开头
      2. 关键词 → 包含"关键词："开头
      3. 一级标题 → "数字 空格 中文" 格式
      4. 二级标题 → "数字.数字 中文" 格式
      5. 图表标题 → "图/表 数字" 开头的短段落
      6. 正文 → 以上都不匹配时的默认类型

    Args:
        text: 段落的纯文本内容（已 strip）

    Returns:
        ParagraphType 常量字符串
    """
    stripped = text.strip()

    if not stripped:
        return ParagraphType.BODY  # 空段落当作正文处理

    # 优先匹配摘要和关键词
    if RE_ABSTRACT.match(stripped):
        return ParagraphType.ABSTRACT

    if RE_KEYWORDS.match(stripped):
        return ParagraphType.KEYWORDS

    # 匹配一级标题（注意：先匹配一级，再匹配二级，避免误判）
    if RE_HEADING_L1.match(stripped):
        return ParagraphType.HEADING_L1

    # 匹配二级标题
    if RE_HEADING_L2.match(stripped):
        return ParagraphType.HEADING_L2

    # 匹配图表标题
    if RE_FIGURE_TABLE.match(stripped):
        return ParagraphType.FIGURE_TABLE

    # 默认为正文
    return ParagraphType.BODY


# ============================================================
# 底层格式设置工具函数
# ============================================================
def _set_run_font(run, cn_font: str, en_font: str, size_pt: float, bold: bool = False, color: RGBColor = None):
    """
    设置 run 级别的字体属性。

    通过直接操作底层 XML 确保中文字体（eastAsia）和西文字体分别正确设置。
    python-docx 的高级 API 无法单独设置 eastAsia 字体，因此需要手动操作 XML。

    Args:
        run:      docx Run 对象
        cn_font:  中文字体名称（如 "宋体"、"黑体"）
        en_font:  西文字体名称（如 "Times New Roman"）
        size_pt:  字号磅值
        bold:     是否加粗
        color:    字体颜色（可选）
    """
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.name = en_font  # 设置西文（Latin）字体

    if color:
        run.font.color.rgb = color

    # 通过底层 XML 设置中文字体（eastAsia）
    # python-docx 不直接支持 eastAsia 字体设置，需要操作 XML
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = parse_xml(f'<w:rFonts {nsdecls("w")} />')
        r_pr.insert(0, r_fonts)

    r_fonts.set(qn("w:eastAsia"), cn_font)
    r_fonts.set(qn("w:ascii"), en_font)
    r_fonts.set(qn("w:hAnsi"), en_font)
    r_fonts.set(qn("w:cs"), en_font)  # 复杂脚本字体也设为西文字体


def _set_paragraph_format(
    paragraph,
    alignment=None,
    first_line_indent=None,
    space_before=None,
    space_after=None,
    line_spacing=None,
    line_spacing_rule=None,
):
    """
    设置段落格式属性。

    Args:
        paragraph:          docx Paragraph 对象
        alignment:          对齐方式（WD_ALIGN_PARAGRAPH 枚举值）
        first_line_indent:  首行缩进（Cm/Pt 等 docx.shared 对象）
        space_before:       段前间距
        space_after:        段后间距
        line_spacing:       行距数值
        line_spacing_rule:  行距规则（WD_LINE_SPACING 枚举值）
    """
    pf = paragraph.paragraph_format

    if alignment is not None:
        pf.alignment = alignment

    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent
    else:
        # 显式清除首行缩进（避免继承模板样式）
        pf.first_line_indent = None

    if space_before is not None:
        pf.space_before = space_before

    if space_after is not None:
        pf.space_after = space_after

    if line_spacing is not None:
        pf.line_spacing = line_spacing

    if line_spacing_rule is not None:
        pf.line_spacing_rule = line_spacing_rule


def _clear_paragraph_style(paragraph):
    """
    清除段落的已有样式设置，防止模板样式干扰排版。
    将段落样式重置为 Normal。
    """
    paragraph.style = "Normal"


# ============================================================
# 各类型段落的格式化函数
# ============================================================
def format_body(paragraph):
    """
    正文格式：
      - 中文字体：宋体
      - 西文字体：Times New Roman
      - 字号：小四（12pt）
      - 首行缩进：2 个中文字符（约 0.74cm × 2 ≈ 对于小四号字约 24pt）
      - 行距：1.5 倍行距
    """
    _clear_paragraph_style(paragraph)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,  # 两端对齐（学术论文常用）
        first_line_indent=Pt(24),               # 2个中文字符（小四12pt × 2 = 24pt）
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )

    for run in paragraph.runs:
        _set_run_font(run, cn_font="宋体", en_font="Times New Roman", size_pt=12)


def format_heading_l1(paragraph):
    """
    一级标题格式：
      - 字体：黑体
      - 字号：三号（16pt）
      - 加粗
      - 居中对齐
      - 段前段后：各 1 行（对于三号字 16pt，1 行间距 ≈ 16pt）
    """
    _clear_paragraph_style(paragraph)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),  # 标题无缩进
        space_before=Pt(16),      # 段前 1 行（三号字高度 16pt）
        space_after=Pt(16),       # 段后 1 行
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )

    for run in paragraph.runs:
        _set_run_font(run, cn_font="黑体", en_font="Times New Roman", size_pt=16, bold=True)


def format_heading_l2(paragraph):
    """
    二级标题格式：
      - 字体：黑体
      - 字号：四号（14pt）
      - 加粗
      - 左对齐
      - 段前段后：各 0.5 行（约 7pt）
    """
    _clear_paragraph_style(paragraph)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Pt(0),  # 标题无缩进
        space_before=Pt(7),       # 段前 0.5 行（四号字 14pt × 0.5 = 7pt）
        space_after=Pt(7),        # 段后 0.5 行
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )

    for run in paragraph.runs:
        _set_run_font(run, cn_font="黑体", en_font="Times New Roman", size_pt=14, bold=True)


def format_figure_table(paragraph):
    """
    图表标题格式：
      - 字体：黑体
      - 字号：五号（10.5pt）
      - 居中对齐
      - 取消首行缩进
    """
    _clear_paragraph_style(paragraph)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),  # 取消首行缩进
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )

    for run in paragraph.runs:
        _set_run_font(run, cn_font="黑体", en_font="Times New Roman", size_pt=10.5)


def format_abstract_or_keywords(paragraph, label_pattern: re.Pattern):
    """
    摘要 / 关键词段落格式：
      - 整体按正文格式（宋体，小四，首行缩进，1.5倍行距）
      - 将标签部分（如 "摘要："、"关键词："）加粗以起强调作用

    实现思路：
      1. 将段落所有 run 的文本拼接
      2. 用正则找到标签的结束位置
      3. 将段落拆分为 "标签 run" 和 "正文 run" 两部分
      4. 标签 run 设为加粗，正文 run 不加粗

    Args:
        paragraph:     docx Paragraph 对象
        label_pattern: 用于匹配标签的正则表达式
    """
    _clear_paragraph_style(paragraph)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_indent=Pt(24),
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )

    # 拼接所有 run 的文本内容
    full_text = paragraph.text
    match = label_pattern.match(full_text)

    if match:
        label_end = match.end()  # 标签部分的字符偏移量

        # ---------- 策略：清空所有旧 run，重新创建两个 run ----------
        # 保存文本
        label_text = full_text[:label_end]
        body_text = full_text[label_end:]

        # 清除段落中所有已有的 run
        for run in paragraph.runs:
            run._element.getparent().remove(run._element)

        # 创建标签 run（加粗）
        run_label = paragraph.add_run(label_text)
        _set_run_font(run_label, cn_font="宋体", en_font="Times New Roman", size_pt=12, bold=True)

        # 创建正文 run（不加粗）
        if body_text:
            run_body = paragraph.add_run(body_text)
            _set_run_font(run_body, cn_font="宋体", en_font="Times New Roman", size_pt=12, bold=False)
    else:
        # 无法匹配标签时，整体按正文格式处理
        for run in paragraph.runs:
            _set_run_font(run, cn_font="宋体", en_font="Times New Roman", size_pt=12)


# ============================================================
# 主函数：学术论文排版
# ============================================================
def format_academic_paper(input_path: str, output_path: str) -> bool:
    """
    读取未排版的 .docx 文档，根据学术论文排版规则进行格式重构，保存为新文档。

    Args:
        input_path:  输入文档路径（.docx）
        output_path: 输出文档路径（.docx）

    Returns:
        True 表示排版成功，False 表示处理失败
    """
    # ---------- 1. 文件校验与读取 ----------
    input_file = Path(input_path)

    if not input_file.exists():
        logger.error(f"输入文件不存在：{input_path}")
        return False

    if not input_file.suffix.lower() == ".docx":
        logger.error(f"不支持的文件格式（仅支持 .docx）：{input_file.suffix}")
        return False

    try:
        doc = Document(input_path)
        logger.info(f"成功读取文档：{input_path}（共 {len(doc.paragraphs)} 个段落）")
    except Exception as e:
        logger.error(f"无法读取文档 {input_path}：{e}")
        return False

    return _process_document(doc, output_path)

def format_academic_paper_from_text(text: str, output_path: str) -> bool:
    """
    将纯文本内容转换为符合学术论文排版规则的 .docx 文档。

    Args:
        text:        输入的纯文本内容（多行）
        output_path: 输出文档路径（.docx）

    Returns:
        True 表示排版成功，False 表示处理失败
    """
    try:
        doc = Document()
        lines = text.split('\n')
        for line in lines:
            # 忽略完全空白的行，或者你也可以保留空行作为段落
            if line.strip() or True: 
                doc.add_paragraph(line)
        logger.info(f"成功从文本创建文档（共 {len(doc.paragraphs)} 个段落）")
    except Exception as e:
        logger.error(f"无法从文本创建文档：{e}")
        return False

    return _process_document(doc, output_path)

def _process_document(doc, output_path: str) -> bool:
    """内部处理逻辑，将 Document 对象排版并保存。"""
    # ---------- 2. 设置默认文档级字体 ----------
    try:
        style = doc.styles["Normal"]
        style.font.name = "Times New Roman"
        style.font.size = Pt(12)

        # 设置 Normal 样式的中文字体为宋体
        r_pr = style.element.get_or_add_rPr()
        r_fonts = r_pr.find(qn("w:rFonts"))
        if r_fonts is None:
            r_fonts = parse_xml(f'<w:rFonts {nsdecls("w")} />')
            r_pr.insert(0, r_fonts)
        r_fonts.set(qn("w:eastAsia"), "宋体")
    except Exception as e:
        logger.warning(f"设置默认样式时出现警告：{e}")

    # ---------- 3. 统计信息 ----------
    stats = {
        ParagraphType.HEADING_L1: 0,
        ParagraphType.HEADING_L2: 0,
        ParagraphType.FIGURE_TABLE: 0,
        ParagraphType.ABSTRACT: 0,
        ParagraphType.KEYWORDS: 0,
        ParagraphType.BODY: 0,
    }

    # ---------- 4. 遍历并格式化每个段落 ----------
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        para_type = classify_paragraph(text)
        stats[para_type] += 1

        if para_type == ParagraphType.HEADING_L1:
            logger.info(f"  [一级标题] 第{i+1}段: \"{text}\"")
            format_heading_l1(paragraph)

        elif para_type == ParagraphType.HEADING_L2:
            logger.info(f"  [二级标题] 第{i+1}段: \"{text}\"")
            format_heading_l2(paragraph)

        elif para_type == ParagraphType.FIGURE_TABLE:
            logger.info(f"  [图表标题] 第{i+1}段: \"{text}\"")
            format_figure_table(paragraph)

        elif para_type == ParagraphType.ABSTRACT:
            logger.info(f"  [摘要段落] 第{i+1}段: \"{text[:30]}...\"")
            format_abstract_or_keywords(paragraph, RE_ABSTRACT)

        elif para_type == ParagraphType.KEYWORDS:
            logger.info(f"  [关键词段] 第{i+1}段: \"{text[:30]}...\"")
            format_abstract_or_keywords(paragraph, RE_KEYWORDS)

        else:
            # 正文段落（含空段落）
            format_body(paragraph)

    # ---------- 5. 保存输出文档 ----------
    try:
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        doc.save(output_path)
        logger.info(f"排版完成！已保存至：{output_path}")
    except Exception as e:
        logger.error(f"无法保存文档 {output_path}：{e}")
        return False

    # ---------- 6. 输出统计摘要 ----------
    logger.info("=" * 50)
    logger.info("排版统计：")
    logger.info(f"  一级标题：{stats[ParagraphType.HEADING_L1]} 个")
    logger.info(f"  二级标题：{stats[ParagraphType.HEADING_L2]} 个")
    logger.info(f"  图表标题：{stats[ParagraphType.FIGURE_TABLE]} 个")
    logger.info(f"  摘要段落：{stats[ParagraphType.ABSTRACT]} 个")
    logger.info(f"  关键词段：{stats[ParagraphType.KEYWORDS]} 个")
    logger.info(f"  正文段落：{stats[ParagraphType.BODY]} 个")
    logger.info("=" * 50)

    return True


# ============================================================
# 命令行入口
# ============================================================
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python format_paper.py <输入文件.docx> [输出文件.docx]")
        print("示例：python format_paper.py 论文初稿.docx 论文_排版后.docx")
        sys.exit(1)

    input_doc = sys.argv[1]

    if len(sys.argv) >= 3:
        output_doc = sys.argv[2]
    else:
        # 默认输出文件名：在原文件名后添加 "_formatted" 后缀
        p = Path(input_doc)
        output_doc = str(p.parent / f"{p.stem}_formatted{p.suffix}")

    success = format_academic_paper(input_doc, output_doc)
    sys.exit(0 if success else 1)
