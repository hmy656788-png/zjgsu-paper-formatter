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
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml

# ============================================================
# 日志配置
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

# ============================================================
# 正则表达式定义（核心匹配逻辑）
# ============================================================

# --- 一级标题匹配 ---
# 匹配规则：以 1-9 开头的数字 + 一个或多个空格 + 至少一个中英文字符
# 示例匹配："1 引言"、"2 研究设计"、"3 模型的估计与检验"
# 要求该段落仅包含这一行内容（独占一行），因此使用 ^ 和 $ 锚定
RE_HEADING_L1 = re.compile(
    r"^\d+\s+[A-Za-z\u4e00-\u9fff][\u4e00-\u9fffA-Za-z0-9\s\-—、（）()/&.,:：]*$"
)

# --- 二级标题匹配 ---
# 匹配规则：数字.数字 + 可选空格 + 至少一个中文字符
# 示例匹配："1.1研究背景"、"2.1 模型构建"、"3.2 数据来源与描述"
RE_HEADING_L2 = re.compile(
    r"^\d+\.\d+\s*[A-Za-z\u4e00-\u9fff][\u4e00-\u9fffA-Za-z0-9\s\-—、（）()/&.,:：]*$"
)

# --- 三级标题匹配 ---
# 示例匹配："1.1.1 研究假设"、"2.3.4 稳健性检验"
RE_HEADING_L3 = re.compile(
    r"^\d+\.\d+\.\d+\s*[A-Za-z\u4e00-\u9fff][\u4e00-\u9fffA-Za-z0-9\s\-—、（）()/&.,:：]*$"
)

# --- 图表标题匹配 ---
# 支持 "图 1 xxx"、"【图9】xxx"、"表8 xxx"、"图n" 等草稿写法
RE_FIGURE_CAPTION = re.compile(
    r"^(?:【\s*)?图\s*(?P<index>\d+|[A-Za-z]+|[一二三四五六七八九十百千万]+)(?:\s*[】\]\)])?\s*[:：.\-—、]?\s*(?P<caption>.*)$"
)
RE_TABLE_CAPTION = re.compile(
    r"^(?:【\s*)?表\s*(?P<index>\d+|[A-Za-z]+|[一二三四五六七八九十百千万]+)(?:\s*[】\]\)])?\s*[:：.\-—、]?\s*(?P<caption>.*)$"
)

# --- 摘要标识匹配 ---
# 匹配规则：段落起始处包含 "摘要" + 可选的标点符号（如 ":"、"："）
RE_ABSTRACT = re.compile(r"^摘\s*要\s*[:：]?\s*")

# --- 关键词标识匹配 ---
# 匹配规则：段落起始处包含 "关键词" + 可选的标点符号
RE_KEYWORDS = re.compile(r"^关\s*键\s*词\s*[:：]?\s*")
RE_REFERENCES_HEADING = re.compile(r"^参\s*考\s*文\s*献\s*$")
RE_SECTION_HEADING = re.compile(
    r"^(致谢|附录|作者简介|基金项目|英文摘要|abstract|acknowledg(?:e)?ments?)\s*$",
    re.IGNORECASE,
)
RE_REFERENCE_ENTRY_TEXT = re.compile(
    r"^(?:\[\d+\]|\(\d+\)|（\d+）|\d+\.\s*|\d+、\s*).+"
)
RE_TITLE_METADATA_PREFIX = re.compile(
    r"^(作者|姓名|学院|学校|专业|指导教师|导师|学号|班级|单位|联系方式|电话|邮箱|电子邮箱|email|e-mail)\s*[:：]",
    re.IGNORECASE,
)

PAGE_LAYOUT = {
    "page_size": "A4",
    "page_width_cm": 21.0,
    "page_height_cm": 29.7,
    "margins_cm": {
        "top": 2.54,
        "bottom": 2.54,
        "left": 3.18,
        "right": 3.18,
    },
    "header_distance_cm": 1.5,
    "footer_distance_cm": 1.5,
}
DEFAULT_HEADER_TEXT = "浙江工商大学学术论文"
RUNNING_HEADER_MAX_LENGTH = 28
CAPTION_MAX_LENGTH = 60


# ============================================================
# 段落分类枚举
# ============================================================
class ParagraphType:
    """段落类型常量"""
    TITLE = "title"                  # 论文标题
    HEADING_L1 = "heading_l1"        # 一级标题
    HEADING_L2 = "heading_l2"        # 二级标题
    HEADING_L3 = "heading_l3"        # 三级标题
    FIGURE_CAPTION = "figure_caption"  # 图标题
    TABLE_CAPTION = "table_caption"    # 表标题
    SECTION_HEADING = "section_heading"  # 非编号章节标题（如致谢/附录）
    REFERENCES_HEADING = "references_heading"  # 参考文献标题
    REFERENCE_ENTRY = "reference_entry"        # 参考文献条目
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
      3. 参考文献标题 → "参考文献"
      4. 非编号章节标题 → "致谢"、"附录" 等独占标题
      5. 一级标题 → "数字 空格 中英文" 格式
      6. 二级标题 → "数字.数字 中英文" 格式
      7. 三级标题 → "数字.数字.数字 中英文" 格式
      8. 图标题 / 表标题 → "图 1"、"【图9】"、"表8" 等短标题
      9. 正文 → 以上都不匹配时的默认类型

    Args:
        text: 段落的纯文本内容（已 strip）

    Returns:
        ParagraphType 常量字符串
    """
    stripped = normalize_text_for_matching(text)

    if not stripped:
        return ParagraphType.BODY  # 空段落当作正文处理

    # 优先匹配摘要和关键词
    if RE_ABSTRACT.match(stripped):
        return ParagraphType.ABSTRACT

    if RE_KEYWORDS.match(stripped):
        return ParagraphType.KEYWORDS

    if RE_REFERENCES_HEADING.match(stripped):
        return ParagraphType.REFERENCES_HEADING

    if RE_SECTION_HEADING.match(stripped):
        return ParagraphType.SECTION_HEADING

    # 匹配一级标题（注意：先匹配一级，再匹配二级，避免误判）
    if RE_HEADING_L1.match(stripped):
        return ParagraphType.HEADING_L1

    # 匹配二级标题
    if RE_HEADING_L2.match(stripped):
        return ParagraphType.HEADING_L2

    # 匹配三级标题
    if RE_HEADING_L3.match(stripped):
        return ParagraphType.HEADING_L3

    caption_match = match_caption(stripped)
    if caption_match:
        return caption_match[0]

    # 默认为正文
    return ParagraphType.BODY


def split_text_to_paragraphs(text: str) -> list[str]:
    """
    规范化纯文本输入并拆分为段落列表。

    - 统一处理 Windows/macOS/Linux 换行符
    - 保留中间空行，便于在导出的文档中保留段落间距
    - 避免把换行控制字符残留到段落正文里
    """
    normalized = (text or "").replace("\r\n", "\n").replace("\r", "\n").lstrip("\ufeff")
    lines = normalized.split("\n")

    while lines and lines[-1] == "":
        lines.pop()

    return lines or [""]


def normalize_text_for_matching(text: str) -> str:
    """
    规范化段落文本，便于处理软回车、全角空格和多个连续空白。

    对标题、图表标题等“独占一行”的识别尤其重要。
    """
    normalized = (text or "").replace("\r", "\n").replace("\v", "\n").replace("\u00a0", " ").replace("\u3000", " ")
    normalized = re.sub(r"\n+", " ", normalized)
    normalized = re.sub(r"[ \t]+", " ", normalized)
    return normalized.strip()


def match_caption(text: str):
    """
    识别图/表标题，并返回类型、匹配结果和规范化后的文本。

    通过长度限制降低误把正文识别为图表标题的风险。
    """
    normalized = normalize_text_for_matching(text)
    if not normalized or len(normalized) > CAPTION_MAX_LENGTH:
        return None

    figure_match = RE_FIGURE_CAPTION.match(normalized)
    if figure_match:
        return ParagraphType.FIGURE_CAPTION, figure_match, normalized

    table_match = RE_TABLE_CAPTION.match(normalized)
    if table_match:
        return ParagraphType.TABLE_CAPTION, table_match, normalized

    return None


def rebuild_caption_text(kind: str, number: int, match: re.Match) -> str:
    """按照统一编号规则重建图/表标题文本。"""
    label = "图" if kind == ParagraphType.FIGURE_CAPTION else "表"
    caption = (match.group("caption") or "").strip()
    return f"{label} {number}" if not caption else f"{label} {number} {caption}"


def is_title_candidate(text: str) -> bool:
    """判断一个段落是否像论文标题。"""
    stripped = normalize_text_for_matching(text)

    if not stripped:
        return False

    if not 6 <= len(stripped) <= 40:
        return False

    if stripped.endswith(("。", "！", "？", "!", "?", "；", ";")):
        return False

    if "@" in stripped or RE_TITLE_METADATA_PREFIX.match(stripped):
        return False

    return classify_paragraph(stripped) == ParagraphType.BODY


def is_reference_entry_text(text: str) -> bool:
    """判断参考文献段落是否具备常见的编号前缀。"""
    return bool(RE_REFERENCE_ENTRY_TEXT.match(normalize_text_for_matching(text)))


def find_title_paragraph_index(paragraphs) -> int | None:
    """
    尝试识别论文主标题。

    规则保持保守：
    - 只考虑前 3 个非空段落中的候选项
    - 段落本身必须像标题
    - 后续 3 个非空段落内需要出现“摘要”或“关键词”
    """
    non_empty = [
        (index, normalize_text_for_matching(paragraph.text))
        for index, paragraph in enumerate(paragraphs)
        if normalize_text_for_matching(paragraph.text)
    ]

    if not non_empty:
        return None

    for candidate_position, (candidate_index, candidate_text) in enumerate(non_empty[:3]):
        if not is_title_candidate(candidate_text):
            continue

        for _, text in non_empty[candidate_position + 1:candidate_position + 4]:
            para_type = classify_paragraph(text)
            if para_type in {ParagraphType.ABSTRACT, ParagraphType.KEYWORDS}:
                return candidate_index

            if para_type in {
                ParagraphType.HEADING_L1,
                ParagraphType.HEADING_L2,
                ParagraphType.FIGURE_CAPTION,
                ParagraphType.TABLE_CAPTION,
            }:
                break

    return None


# ============================================================
# 底层格式设置工具函数
# ============================================================
def _set_run_font(run, cn_font: str, en_font: str, size_pt: float, bold: bool | None = None, color: RGBColor = None):
    """
    设置 run 级别的字体属性。

    通过直接操作底层 XML 确保中文字体（eastAsia）和西文字体分别正确设置。
    python-docx 的高级 API 无法单独设置 eastAsia 字体，因此需要手动操作 XML。

    Args:
        run:      docx Run 对象
        cn_font:  中文字体名称（如 "宋体"、"黑体"）
        en_font:  西文字体名称（如 "Times New Roman"）
        size_pt:  字号磅值
        bold:     是否加粗；传 None 时保留原有加粗状态
        color:    字体颜色（可选）
    """
    run.font.size = Pt(size_pt)
    if bold is not None:
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


def _remove_all_runs(paragraph):
    """删除段落中已有的 run，便于重建页眉页脚等结构。"""
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)


def _replace_paragraph_text(paragraph, text: str):
    """安全重建纯文本段落内容。仅用于标题、图表标题等纯文本段落。"""
    _remove_all_runs(paragraph)
    if text:
        paragraph.add_run(text)


def _apply_run_fonts(paragraph, cn_font: str, en_font: str, size_pt: float, bold: bool | None = None):
    """遍历段落中所有 run，统一设置中西文字体。"""
    for run in paragraph.runs:
        current_bold = run.font.bold if bold is None else bold
        _set_run_font(run, cn_font=cn_font, en_font=en_font, size_pt=size_pt, bold=current_bold)


def _is_list_paragraph(paragraph) -> bool:
    """判断段落是否为项目符号/编号列表，尽量保留其原有列表样式和缩进。"""
    style_name = ""
    if paragraph.style is not None:
        style_name = (paragraph.style.name or "").lower()

    if style_name.startswith("list"):
        return True

    p_pr = paragraph._element.pPr
    return p_pr is not None and p_pr.numPr is not None


def iter_table_paragraphs(tables):
    """递归遍历所有表格单元格内的段落。"""
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph
                yield from iter_table_paragraphs(cell.tables)


def _has_drawing(paragraph) -> bool:
    """判断段落中是否包含图片等 drawing 元素。"""
    return bool(paragraph._element.findall(".//" + qn("w:drawing")))


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


def _clear_paragraph_style(paragraph, preserve_list_style: bool = False):
    """
    清除段落的已有样式设置，防止模板样式干扰排版。
    将段落样式重置为 Normal。
    """
    if preserve_list_style and _is_list_paragraph(paragraph):
        return

    paragraph.style = "Normal"


def _get_primary_paragraph(container):
    """获取页眉/页脚中的主段落，并清理多余段落。"""
    paragraphs = list(container.paragraphs)
    paragraph = paragraphs[0] if paragraphs else container.add_paragraph()

    for extra in paragraphs[1:]:
        extra._element.getparent().remove(extra._element)

    _remove_all_runs(paragraph)
    return paragraph


def _append_page_number_field(paragraph):
    """在页脚插入 Word 自动页码字段。"""
    size_pt = 10.5

    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    run_begin = paragraph.add_run()
    _set_run_font(run_begin, cn_font="宋体", en_font="Times New Roman", size_pt=size_pt)
    run_begin._r.append(fld_char_begin)

    instr_text = OxmlElement("w:instrText")
    instr_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr_text.text = "PAGE"
    run_instr = paragraph.add_run()
    _set_run_font(run_instr, cn_font="宋体", en_font="Times New Roman", size_pt=size_pt)
    run_instr._r.append(instr_text)

    fld_char_separate = OxmlElement("w:fldChar")
    fld_char_separate.set(qn("w:fldCharType"), "separate")
    run_separate = paragraph.add_run()
    _set_run_font(run_separate, cn_font="宋体", en_font="Times New Roman", size_pt=size_pt)
    run_separate._r.append(fld_char_separate)

    run_text = paragraph.add_run("1")
    _set_run_font(run_text, cn_font="宋体", en_font="Times New Roman", size_pt=size_pt)

    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")
    run_end = paragraph.add_run()
    _set_run_font(run_end, cn_font="宋体", en_font="Times New Roman", size_pt=size_pt)
    run_end._r.append(fld_char_end)


def apply_document_layout(doc, header_text: str) -> dict:
    """统一设置纸张、页边距、页眉和页码。"""
    raw_header = (header_text or "").strip()
    if not raw_header:
        header_value = DEFAULT_HEADER_TEXT
    elif len(raw_header) <= RUNNING_HEADER_MAX_LENGTH:
        header_value = raw_header
    else:
        header_value = f"{raw_header[:RUNNING_HEADER_MAX_LENGTH - 3].rstrip()}..."

    for section in doc.sections:
        section.page_width = Cm(PAGE_LAYOUT["page_width_cm"])
        section.page_height = Cm(PAGE_LAYOUT["page_height_cm"])
        section.top_margin = Cm(PAGE_LAYOUT["margins_cm"]["top"])
        section.bottom_margin = Cm(PAGE_LAYOUT["margins_cm"]["bottom"])
        section.left_margin = Cm(PAGE_LAYOUT["margins_cm"]["left"])
        section.right_margin = Cm(PAGE_LAYOUT["margins_cm"]["right"])
        section.header_distance = Cm(PAGE_LAYOUT["header_distance_cm"])
        section.footer_distance = Cm(PAGE_LAYOUT["footer_distance_cm"])
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False

        header_paragraph = _get_primary_paragraph(section.header)
        _clear_paragraph_style(header_paragraph)
        _set_paragraph_format(
            header_paragraph,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            first_line_indent=Pt(0),
            line_spacing=1.0,
            line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
        )
        header_run = header_paragraph.add_run(header_value)
        _set_run_font(header_run, cn_font="宋体", en_font="Times New Roman", size_pt=10.5)

        footer_paragraph = _get_primary_paragraph(section.footer)
        _clear_paragraph_style(footer_paragraph)
        _set_paragraph_format(
            footer_paragraph,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            first_line_indent=Pt(0),
            line_spacing=1.0,
            line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
        )
        _append_page_number_field(footer_paragraph)

    return {
        "page_size": PAGE_LAYOUT["page_size"],
        "page_width_cm": PAGE_LAYOUT["page_width_cm"],
        "page_height_cm": PAGE_LAYOUT["page_height_cm"],
        "margins_cm": PAGE_LAYOUT["margins_cm"].copy(),
        "header_distance_cm": PAGE_LAYOUT["header_distance_cm"],
        "footer_distance_cm": PAGE_LAYOUT["footer_distance_cm"],
        "header_text": header_value,
        "page_number_position": "footer_center",
    }


# ============================================================
# 各类型段落的格式化函数
# ============================================================
def format_body(paragraph, in_table: bool = False):
    """
    正文格式：
      - 中文字体：宋体
      - 西文字体：Times New Roman
      - 字号：小四（12pt）
      - 首行缩进：2 个中文字符（约 0.74cm × 2 ≈ 对于小四号字约 24pt）
      - 行距：1.5 倍行距
    """
    normalized_text = normalize_text_for_matching(paragraph.text)

    if _has_drawing(paragraph) and not normalized_text:
        _set_paragraph_format(
            paragraph,
            alignment=paragraph.paragraph_format.alignment or WD_ALIGN_PARAGRAPH.CENTER,
            first_line_indent=Pt(0),
        )
        _apply_run_fonts(paragraph, cn_font="宋体", en_font="Times New Roman", size_pt=12)
        return

    if _is_list_paragraph(paragraph):
        # 列表段落保留原有项目符号/编号样式，仅统一行距和字体。
        pf = paragraph.paragraph_format
        pf.line_spacing = 1.5
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    else:
        _clear_paragraph_style(paragraph)
        _set_paragraph_format(
            paragraph,
            alignment=WD_ALIGN_PARAGRAPH.LEFT if in_table else WD_ALIGN_PARAGRAPH.JUSTIFY,
            first_line_indent=Pt(0) if in_table else Pt(24),
            line_spacing=1.5,
            line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
        )

    _apply_run_fonts(paragraph, cn_font="宋体", en_font="Times New Roman", size_pt=12)


def format_title(paragraph, text_override: str | None = None):
    """
    论文标题格式：
      - 字体：黑体
      - 字号：小二（18pt）
      - 加粗
      - 居中对齐
      - 段后适当留白，和摘要区隔开
    """
    _clear_paragraph_style(paragraph)
    if text_override is not None:
        _replace_paragraph_text(paragraph, text_override)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),
        space_after=Pt(18),
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )

    _apply_run_fonts(paragraph, cn_font="黑体", en_font="Times New Roman", size_pt=18, bold=True)


def format_heading_l1(paragraph, text_override: str | None = None):
    """
    一级标题格式：
      - 字体：黑体
      - 字号：三号（16pt）
      - 加粗
      - 居中对齐
      - 段前段后：各 1 行（对于三号字 16pt，1 行间距 ≈ 16pt）
    """
    _clear_paragraph_style(paragraph)
    if text_override is not None:
        _replace_paragraph_text(paragraph, text_override)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),  # 标题无缩进
        space_before=Pt(16),      # 段前 1 行（三号字高度 16pt）
        space_after=Pt(16),       # 段后 1 行
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )

    _apply_run_fonts(paragraph, cn_font="黑体", en_font="Times New Roman", size_pt=16, bold=True)


def format_heading_l2(paragraph, text_override: str | None = None):
    """
    二级标题格式：
      - 字体：黑体
      - 字号：四号（14pt）
      - 加粗
      - 左对齐
      - 段前段后：各 0.5 行（约 7pt）
    """
    _clear_paragraph_style(paragraph)
    if text_override is not None:
        _replace_paragraph_text(paragraph, text_override)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Pt(0),  # 标题无缩进
        space_before=Pt(7),       # 段前 0.5 行（四号字 14pt × 0.5 = 7pt）
        space_after=Pt(7),        # 段后 0.5 行
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )

    _apply_run_fonts(paragraph, cn_font="黑体", en_font="Times New Roman", size_pt=14, bold=True)


def format_heading_l3(paragraph, text_override: str | None = None):
    """
    三级标题格式：
      - 字体：宋体
      - 字号：小四（12pt）
      - 加粗
      - 左对齐
      - 段前段后：各 0.5 行
    """
    _clear_paragraph_style(paragraph)
    if text_override is not None:
        _replace_paragraph_text(paragraph, text_override)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Pt(0),
        space_before=Pt(6),
        space_after=Pt(6),
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )

    _apply_run_fonts(paragraph, cn_font="宋体", en_font="Times New Roman", size_pt=12, bold=True)


def format_figure_table(paragraph, text_override: str | None = None):
    """
    图表标题格式：
      - 字体：黑体
      - 字号：五号（10.5pt）
      - 居中对齐
      - 取消首行缩进
    """
    _clear_paragraph_style(paragraph)
    if text_override is not None:
        _replace_paragraph_text(paragraph, text_override)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),  # 取消首行缩进
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )

    _apply_run_fonts(paragraph, cn_font="黑体", en_font="Times New Roman", size_pt=10.5)


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
        _apply_run_fonts(paragraph, cn_font="宋体", en_font="Times New Roman", size_pt=12)


def format_reference_entry(paragraph):
    """
    参考文献条目格式：
      - 宋体 / Times New Roman，小四号
      - 悬挂缩进 2 个字符
      - 1.5 倍行距
    """
    _clear_paragraph_style(paragraph)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_indent=Pt(-24),
        space_after=Pt(3),
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )
    paragraph.paragraph_format.left_indent = Pt(24)

    _apply_run_fonts(paragraph, cn_font="宋体", en_font="Times New Roman", size_pt=12)


def _append_outline_entry(outline: list[dict], para_type: str, text: str):
    """记录识别到的结构化大纲，供前端预览展示。"""
    if not text:
        return

    level = {
        ParagraphType.TITLE: "title",
        ParagraphType.HEADING_L1: "h1",
        ParagraphType.HEADING_L2: "h2",
        ParagraphType.HEADING_L3: "h3",
        ParagraphType.SECTION_HEADING: "section",
        ParagraphType.REFERENCES_HEADING: "references",
    }.get(para_type)

    if level is None:
        return

    outline.append({"level": level, "text": text})


# ============================================================
# 主函数：学术论文排版
# ============================================================
def format_academic_paper(input_path: str, output_path: str) -> dict | bool:
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

def format_academic_paper_from_text(text: str, output_path: str) -> dict | bool:
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
        for line in split_text_to_paragraphs(text):
            doc.add_paragraph(line)
        logger.info(f"成功从文本创建文档（共 {len(doc.paragraphs)} 个段落）")
    except Exception as e:
        logger.error(f"无法从文本创建文档：{e}")
        return False

    return _process_document(doc, output_path)

def _process_document(doc, output_path: str) -> dict | bool:
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
        ParagraphType.TITLE: 0,
        ParagraphType.HEADING_L1: 0,
        ParagraphType.HEADING_L2: 0,
        ParagraphType.HEADING_L3: 0,
        ParagraphType.FIGURE_CAPTION: 0,
        ParagraphType.TABLE_CAPTION: 0,
        ParagraphType.SECTION_HEADING: 0,
        ParagraphType.REFERENCES_HEADING: 0,
        ParagraphType.REFERENCE_ENTRY: 0,
        ParagraphType.ABSTRACT: 0,
        ParagraphType.KEYWORDS: 0,
        ParagraphType.BODY: 0,
    }

    title_index = find_title_paragraph_index(doc.paragraphs)
    title_text = ""
    if title_index is not None:
        title_text = normalize_text_for_matching(doc.paragraphs[title_index].text)

    page_setup = apply_document_layout(doc, title_text)
    outline = []
    in_references = False
    figure_counter = 0
    table_counter = 0

    # ---------- 4. 遍历并格式化每个段落 ----------
    for i, paragraph in enumerate(doc.paragraphs):
        raw_text = paragraph.text
        text = normalize_text_for_matching(raw_text)
        para_type = ParagraphType.TITLE if i == title_index else classify_paragraph(raw_text)

        if para_type == ParagraphType.REFERENCES_HEADING:
            in_references = True
        elif in_references and text:
            if para_type in {
                ParagraphType.HEADING_L1,
                ParagraphType.HEADING_L2,
                ParagraphType.HEADING_L3,
                ParagraphType.SECTION_HEADING,
            }:
                in_references = False
            elif is_reference_entry_text(text):
                para_type = ParagraphType.REFERENCE_ENTRY
            else:
                in_references = False

        stats[para_type] += 1
        _append_outline_entry(outline, para_type, text)

        if para_type == ParagraphType.TITLE:
            logger.info(f"  [论文标题] 第{i+1}段: \"{text}\"")
            format_title(paragraph, text_override=text)

        elif para_type == ParagraphType.HEADING_L1:
            logger.info(f"  [一级标题] 第{i+1}段: \"{text}\"")
            format_heading_l1(paragraph, text_override=text)

        elif para_type == ParagraphType.HEADING_L2:
            logger.info(f"  [二级标题] 第{i+1}段: \"{text}\"")
            format_heading_l2(paragraph, text_override=text)

        elif para_type == ParagraphType.HEADING_L3:
            logger.info(f"  [三级标题] 第{i+1}段: \"{text}\"")
            format_heading_l3(paragraph, text_override=text)

        elif para_type == ParagraphType.FIGURE_CAPTION:
            figure_counter += 1
            caption_match = match_caption(raw_text)
            caption_text = rebuild_caption_text(ParagraphType.FIGURE_CAPTION, figure_counter, caption_match[1])
            logger.info(f"  [图标题] 第{i+1}段: \"{caption_text}\"")
            format_figure_table(paragraph, text_override=caption_text)

        elif para_type == ParagraphType.TABLE_CAPTION:
            table_counter += 1
            caption_match = match_caption(raw_text)
            caption_text = rebuild_caption_text(ParagraphType.TABLE_CAPTION, table_counter, caption_match[1])
            logger.info(f"  [表标题] 第{i+1}段: \"{caption_text}\"")
            format_figure_table(paragraph, text_override=caption_text)

        elif para_type == ParagraphType.SECTION_HEADING:
            logger.info(f"  [非编号章节标题] 第{i+1}段: \"{text}\"")
            format_heading_l1(paragraph, text_override=text)

        elif para_type == ParagraphType.REFERENCES_HEADING:
            logger.info(f"  [参考文献标题] 第{i+1}段: \"{text}\"")
            format_heading_l1(paragraph, text_override=text)

        elif para_type == ParagraphType.REFERENCE_ENTRY:
            logger.info(f"  [参考文献条目] 第{i+1}段: \"{text[:30]}...\"")
            format_reference_entry(paragraph)

        elif para_type == ParagraphType.ABSTRACT:
            logger.info(f"  [摘要段落] 第{i+1}段: \"{text[:30]}...\"")
            format_abstract_or_keywords(paragraph, RE_ABSTRACT)

        elif para_type == ParagraphType.KEYWORDS:
            logger.info(f"  [关键词段] 第{i+1}段: \"{text[:30]}...\"")
            format_abstract_or_keywords(paragraph, RE_KEYWORDS)

        else:
            # 正文段落（含空段落）
            format_body(paragraph)

    # ---------- 5. 处理表格内段落 ----------
    table_paragraph_count = 0
    for paragraph in iter_table_paragraphs(doc.tables):
        table_paragraph_count += 1
        format_body(paragraph, in_table=True)

    # ---------- 6. 保存输出文档 ----------
    try:
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        doc.save(output_path)
        logger.info(f"排版完成！已保存至：{output_path}")
    except Exception as e:
        logger.error(f"无法保存文档 {output_path}：{e}")
        return False

    # ---------- 7. 输出统计摘要 ----------
    logger.info("=" * 50)
    logger.info("排版统计：")
    logger.info(f"  论文标题：{stats[ParagraphType.TITLE]} 个")
    logger.info(f"  一级标题：{stats[ParagraphType.HEADING_L1]} 个")
    logger.info(f"  二级标题：{stats[ParagraphType.HEADING_L2]} 个")
    logger.info(f"  三级标题：{stats[ParagraphType.HEADING_L3]} 个")
    logger.info(f"  图标题：{stats[ParagraphType.FIGURE_CAPTION]} 个")
    logger.info(f"  表标题：{stats[ParagraphType.TABLE_CAPTION]} 个")
    logger.info(f"  非编号章节标题：{stats[ParagraphType.SECTION_HEADING]} 个")
    logger.info(f"  参考文献标题：{stats[ParagraphType.REFERENCES_HEADING]} 个")
    logger.info(f"  参考文献条目：{stats[ParagraphType.REFERENCE_ENTRY]} 条")
    logger.info(f"  摘要段落：{stats[ParagraphType.ABSTRACT]} 个")
    logger.info(f"  关键词段：{stats[ParagraphType.KEYWORDS]} 个")
    logger.info(f"  正文段落：{stats[ParagraphType.BODY]} 个")
    logger.info(f"  表格内段落：{table_paragraph_count} 个")
    logger.info("=" * 50)

    return {
        "stats": stats,
        "title_text": title_text,
        "page_setup": page_setup,
        "table_paragraphs": table_paragraph_count,
        "outline": outline,
    }


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
