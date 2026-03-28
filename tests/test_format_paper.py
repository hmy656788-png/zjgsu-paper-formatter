import base64
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches

from format_paper import (
    find_title_paragraph_index,
    format_academic_paper,
    format_academic_paper_from_text,
    split_text_to_paragraphs,
)


class FormatPaperFromTextTestCase(unittest.TestCase):
    @staticmethod
    def assert_run_uses_mixed_font_pair(test_case, run):
        r_fonts = run._element.rPr.rFonts
        test_case.assertEqual(r_fonts.get(qn("w:eastAsia")), "宋体")
        test_case.assertEqual(r_fonts.get(qn("w:ascii")), "Times New Roman")
        test_case.assertEqual(r_fonts.get(qn("w:hAnsi")), "Times New Roman")

    def test_split_text_to_paragraphs_normalizes_line_endings(self):
        text = "第一段\r\n\r\n第二段\r第三段\n"

        self.assertEqual(
            split_text_to_paragraphs(text),
            ["第一段", "", "第二段", "第三段"],
        )

    def test_format_academic_paper_from_text_preserves_blank_lines_without_newline_chars(self):
        text = "第一段\r\n\r\n第二段\r\n"

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)

        try:
            self.assertTrue(format_academic_paper_from_text(text, str(output_path)))

            doc = Document(str(output_path))
            self.assertEqual([paragraph.text for paragraph in doc.paragraphs], ["第一段", "", "第二段"])
        finally:
            output_path.unlink(missing_ok=True)

    def test_find_title_paragraph_index_only_when_abstract_follows(self):
        doc = Document()
        doc.add_paragraph("基于多元回归模型的城市化研究")
        doc.add_paragraph("作者：张三")
        doc.add_paragraph("摘要：这是摘要内容")

        self.assertEqual(find_title_paragraph_index(doc.paragraphs), 0)

    def test_find_title_paragraph_index_skips_metadata_prefix(self):
        doc = Document()
        doc.add_paragraph("作者：张三")
        doc.add_paragraph("基于多元回归模型的城市化研究")
        doc.add_paragraph("摘要：这是摘要内容")

        self.assertEqual(find_title_paragraph_index(doc.paragraphs), 1)

    def test_find_title_paragraph_index_rejects_metadata_without_real_title(self):
        doc = Document()
        doc.add_paragraph("作者：张三")
        doc.add_paragraph("摘要：这是摘要内容")
        doc.add_paragraph("关键词：城市化 回归")

        self.assertIsNone(find_title_paragraph_index(doc.paragraphs))

    def test_format_academic_paper_from_text_formats_detected_title(self):
        text = "基于多元回归模型的城市化研究\n摘要：这是摘要内容\n关键词：城市化 回归\n1 引言\n正文内容"

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)

        try:
            self.assertTrue(format_academic_paper_from_text(text, str(output_path)))

            doc = Document(str(output_path))
            title = doc.paragraphs[0]
            self.assertEqual(title.text, "基于多元回归模型的城市化研究")
            self.assertEqual(title.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertTrue(title.runs[0].font.bold)
            self.assertEqual(title.runs[0].font.size.pt, 18.0)
        finally:
            output_path.unlink(missing_ok=True)

    def test_format_academic_paper_from_text_applies_page_layout_and_page_number(self):
        text = "基于多元回归模型的城市化研究\n摘要：这是摘要内容\n关键词：城市化 回归\n1 引言\n正文内容"

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)

        try:
            summary = format_academic_paper_from_text(text, str(output_path))

            self.assertIsInstance(summary, dict)
            self.assertEqual(summary["page_setup"]["page_size"], "A4")
            self.assertEqual(summary["page_setup"]["header_text"], "基于多元回归模型的城市化研究")

            doc = Document(str(output_path))
            section = doc.sections[0]

            self.assertAlmostEqual(section.page_width.cm, 21.0, places=1)
            self.assertAlmostEqual(section.page_height.cm, 29.7, places=1)
            self.assertAlmostEqual(section.top_margin.cm, 2.54, places=1)
            self.assertAlmostEqual(section.left_margin.cm, 3.18, places=1)
            self.assertEqual(section.header.paragraphs[0].text, "基于多元回归模型的城市化研究")
            self.assertIn("PAGE", section.footer._element.xml)
        finally:
            output_path.unlink(missing_ok=True)

    def test_format_academic_paper_from_text_formats_heading_l3_and_references(self):
        text = (
            "基于多元回归模型的城市化研究\n"
            "摘要：这是摘要内容\n"
            "关键词：城市化 回归\n"
            "1 引言\n"
            "1.1 研究背景\n"
            "1.1.1 研究假设\n"
            "正文内容\n"
            "参考文献\n"
            "[1] 张三. 学术论文写作规范[J]. 高教研究, 2024.\n"
            "[2] 李四. 文献整理方法[M]. 北京: 科学出版社, 2023."
        )

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)

        try:
            summary = format_academic_paper_from_text(text, str(output_path))

            self.assertIsInstance(summary, dict)
            self.assertEqual(summary["stats"]["heading_l3"], 1)
            self.assertEqual(summary["stats"]["references_heading"], 1)
            self.assertEqual(summary["stats"]["reference_entry"], 2)
            self.assertTrue(any(item["level"] == "h3" for item in summary["outline"]))
            self.assertTrue(any(item["level"] == "references" for item in summary["outline"]))

            doc = Document(str(output_path))
            heading_l3 = doc.paragraphs[5]
            references_heading = doc.paragraphs[7]
            reference_entry = doc.paragraphs[8]

            self.assertEqual(heading_l3.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.LEFT)
            self.assertTrue(heading_l3.runs[0].font.bold)
            self.assertEqual(heading_l3.runs[0].font.size.pt, 12.0)
            self.assertEqual(heading_l3.paragraph_format.first_line_indent.pt, 0.0)

            self.assertEqual(references_heading.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertAlmostEqual(reference_entry.paragraph_format.left_indent.pt, 24.0, places=1)
            self.assertAlmostEqual(reference_entry.paragraph_format.first_line_indent.pt, -24.0, places=1)
        finally:
            output_path.unlink(missing_ok=True)

    def test_format_academic_paper_from_text_stops_references_before_acknowledgements(self):
        text = (
            "基于多元回归模型的城市化研究\n"
            "摘要：这是摘要内容\n"
            "关键词：城市化 回归\n"
            "参考文献\n"
            "[1] 张三. 学术论文写作规范[J]. 高教研究, 2024.\n"
            "致谢\n"
            "感谢导师和同学们的帮助。"
        )

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)

        try:
            summary = format_academic_paper_from_text(text, str(output_path))

            self.assertIsInstance(summary, dict)
            self.assertEqual(summary["stats"]["reference_entry"], 1)
            self.assertEqual(summary["stats"]["section_heading"], 1)
            self.assertTrue(any(item["level"] == "section" for item in summary["outline"]))

            doc = Document(str(output_path))
            acknowledgements = doc.paragraphs[5]
            acknowledgement_body = doc.paragraphs[6]

            self.assertEqual(acknowledgements.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertEqual(acknowledgement_body.paragraph_format.first_line_indent.pt, 24.0)
        finally:
            output_path.unlink(missing_ok=True)

    def test_format_academic_paper_from_text_shortens_running_header_for_long_title(self):
        text = (
            "基于多源异构数据融合与深度学习模型的中国城市高质量发展测度及影响机制研究\n"
            "摘要：这是摘要内容\n"
            "关键词：城市化 回归\n"
            "正文内容"
        )

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)

        try:
            summary = format_academic_paper_from_text(text, str(output_path))

            self.assertIsInstance(summary, dict)
            self.assertTrue(summary["page_setup"]["header_text"].endswith("..."))
            self.assertLessEqual(len(summary["page_setup"]["header_text"]), 28)

            doc = Document(str(output_path))
            self.assertEqual(doc.sections[0].header.paragraphs[0].text, summary["page_setup"]["header_text"])
        finally:
            output_path.unlink(missing_ok=True)

    def test_format_academic_paper_handles_complex_doc_with_images_table_and_lists(self):
        tiny_png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            image_path = temp_path / "pixel.png"
            input_path = temp_path / "complex_input.docx"
            output_path = temp_path / "complex_output.docx"
            image_path.write_bytes(tiny_png)

            doc = Document()
            doc.add_paragraph("跨境电商场景下 Python 模型驱动的供应链韧性研究")
            doc.add_paragraph("摘要：本文使用 Python 3.11 和 Cross-border E-commerce 数据进行分析。")
            doc.add_paragraph("关键词：Python Cross-border E-commerce 2024")

            heading = doc.add_paragraph()
            heading.add_run("1")
            heading.add_run().add_break()
            heading.add_run("引言")

            doc.add_paragraph("本文基于 Python 3.11 与 Cross-border E-commerce 2024 数据进行实证分析。")
            doc.add_paragraph("First bullet uses Python 3.11", style="List Bullet")

            picture_1 = doc.add_paragraph()
            picture_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            picture_1.add_run().add_picture(str(image_path), width=Inches(0.2))
            doc.add_paragraph("【图9】 系统架构图")

            picture_2 = doc.add_paragraph()
            picture_2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            picture_2.add_run().add_picture(str(image_path), width=Inches(0.2))
            doc.add_paragraph("图8 模型流程图")

            doc.add_paragraph("表88 样本描述统计")
            table = doc.add_table(rows=2, cols=2)
            table.cell(0, 0).text = "变量"
            table.cell(0, 1).text = "Python 3.11"
            table.cell(1, 0).text = "平台"
            table.cell(1, 1).text = "Cross-border E-commerce 2024"

            doc.add_paragraph("参考文献")
            doc.add_paragraph("[9] Smith, John. Python-based trade analytics[J]. 2024.")
            doc.save(str(input_path))

            summary = format_academic_paper(str(input_path), str(output_path))
            self.assertIsInstance(summary, dict)
            self.assertEqual(summary["stats"]["figure_caption"], 2)
            self.assertEqual(summary["stats"]["table_caption"], 1)
            self.assertEqual(summary["table_paragraphs"], 4)

            output_doc = Document(str(output_path))
            self.assertEqual(len(output_doc.inline_shapes), 2)

            self.assertEqual(output_doc.paragraphs[3].text, "1 引言")
            self.assertEqual(output_doc.paragraphs[7].text, "图 1 系统架构图")
            self.assertEqual(output_doc.paragraphs[9].text, "图 2 模型流程图")
            self.assertEqual(output_doc.paragraphs[10].text, "表 1 样本描述统计")
            self.assertEqual(output_doc.paragraphs[5].style.name, "List Bullet")
            self.assertEqual(output_doc.paragraphs[6].paragraph_format.first_line_indent.pt, 0.0)

            body_run = output_doc.paragraphs[4].runs[0]
            table_run = output_doc.tables[0].cell(1, 1).paragraphs[0].runs[0]
            self.assert_run_uses_mixed_font_pair(self, body_run)
            self.assert_run_uses_mixed_font_pair(self, table_run)

            reference_entry = output_doc.paragraphs[12]
            self.assertAlmostEqual(reference_entry.paragraph_format.left_indent.pt, 24.0, places=1)
            self.assertAlmostEqual(reference_entry.paragraph_format.first_line_indent.pt, -24.0, places=1)


if __name__ == "__main__":
    unittest.main()
