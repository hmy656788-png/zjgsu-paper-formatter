import base64
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.oxml.ns import nsdecls
from docx.shared import Inches
from lxml import etree

from format_paper import (
    apply_document_layout,
    ensure_document_ends_with_page_break,
    find_title_paragraph_index,
    format_academic_paper,
    format_academic_paper_from_text,
    generate_cover_page,
    split_text_to_paragraphs,
)


class FormatPaperFromTextTestCase(unittest.TestCase):
    CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
    PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    FOOTNOTE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"

    @staticmethod
    def assert_run_uses_mixed_font_pair(test_case, run):
        r_fonts = run._element.rPr.rFonts
        test_case.assertEqual(r_fonts.get(qn("w:eastAsia")), "宋体")
        test_case.assertEqual(r_fonts.get(qn("w:ascii")), "Times New Roman")
        test_case.assertEqual(r_fonts.get(qn("w:hAnsi")), "Times New Roman")

    def assert_paragraph_has_numbering(self, paragraph, ilvl: int) -> int:
        num_pr = paragraph._element.pPr.find(qn("w:numPr"))
        self.assertIsNotNone(num_pr)
        self.assertEqual(num_pr.find(qn("w:ilvl")).get(qn("w:val")), str(ilvl))
        return int(num_pr.find(qn("w:numId")).get(qn("w:val")))

    def assert_numbering_overrides(self, doc, num_id: int, expected_starts: dict[int, int]):
        numbering = doc.part.numbering_part.numbering_definitions._numbering
        num = numbering.num_having_numId(num_id)

        for ilvl, start_value in expected_starts.items():
            lvl_override = next(
                override
                for override in num.findall("./" + qn("w:lvlOverride"))
                if override.get(qn("w:ilvl")) == str(ilvl)
            )
            start_override = lvl_override.find(qn("w:startOverride"))
            self.assertIsNotNone(start_override)
            self.assertEqual(start_override.get(qn("w:val")), str(start_value))

    def inject_simple_footnote(self, docx_path: Path, footnote_text: str):
        with ZipFile(docx_path, "r") as source:
            entries = source.infolist()
            payloads = {entry.filename: source.read(entry.filename) for entry in entries}

        content_types = etree.fromstring(payloads["[Content_Types].xml"])
        override_tag = f"{{{self.CONTENT_TYPES_NS}}}Override"
        if not any(node.get("PartName") == "/word/footnotes.xml" for node in content_types.findall(override_tag)):
            override = etree.Element(override_tag)
            override.set("PartName", "/word/footnotes.xml")
            override.set(
                "ContentType",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            )
            content_types.append(override)
        payloads["[Content_Types].xml"] = etree.tostring(
            content_types,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        rels = etree.fromstring(payloads["word/_rels/document.xml.rels"])
        relationship_tag = f"{{{self.PACKAGE_REL_NS}}}Relationship"
        if not any(node.get("Type") == self.FOOTNOTE_REL_TYPE for node in rels.findall(relationship_tag)):
            relationship = etree.Element(relationship_tag)
            relationship.set("Id", "rIdFootnotes")
            relationship.set("Type", self.FOOTNOTE_REL_TYPE)
            relationship.set("Target", "footnotes.xml")
            rels.append(relationship)
        payloads["word/_rels/document.xml.rels"] = etree.tostring(
            rels,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        document_xml = etree.fromstring(payloads["word/document.xml"])
        body = document_xml.find(qn("w:body"))
        target_paragraph = body.findall(qn("w:p"))[-1]
        target_paragraph.append(
            parse_xml(
                f'<w:r {nsdecls("w")}>'
                '<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                '<w:footnoteReference w:id="2"/>'
                "</w:r>"
            )
        )
        payloads["word/document.xml"] = etree.tostring(
            document_xml,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        payloads["word/footnotes.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:footnotes {nsdecls("w")}>'
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
            '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
            '<w:footnote w:id="2"><w:p>'
            '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
            f"<w:r><w:t xml:space=\"preserve\"> {footnote_text}</w:t></w:r>"
            "</w:p></w:footnote>"
            "</w:footnotes>"
        ).encode("utf-8")

        with ZipFile(docx_path, "w") as target:
            written = set()
            for entry in entries:
                target.writestr(entry, payloads[entry.filename])
                written.add(entry.filename)
            for name, data in payloads.items():
                if name not in written:
                    target.writestr(name, data)

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

    def test_find_title_paragraph_index_accepts_english_abstract_heading(self):
        doc = Document()
        doc.add_paragraph("跨境电商场景下供应链韧性研究")
        doc.add_paragraph("Abstract")
        doc.add_paragraph("This paper studies supply chain resilience.")

        self.assertEqual(find_title_paragraph_index(doc.paragraphs), 0)

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

    def test_format_academic_paper_from_text_generates_cover_from_cover_info(self):
        text = (
            "企业数字化转型对绿色技术创新的影响研究\n"
            "摘要：这是摘要内容\n"
            "关键词：数字化 创新\n"
            "正文内容"
        )

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)

        try:
            summary = format_academic_paper_from_text(
                text,
                str(output_path),
                cover_info={
                    "cover_title": "《大数据挖掘》期末大作业",
                    "college": "工商管理学院",
                    "teacher": "刘璇",
                    "class_name": "国商2301",
                    "student_name": "何旻洋",
                    "student_id": "2320100731",
                },
            )

            self.assertTrue(summary["cover_generated"])
            output_doc = Document(str(output_path))
            self.assertEqual(output_doc.paragraphs[2].text, "《大数据挖掘》期末大作业")
            self.assertEqual(output_doc.tables[0].cell(0, 1).text, "工商管理学院")
            self.assertEqual(output_doc.tables[0].cell(4, 1).text, "2320100731")
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
            heading_l1 = next(paragraph for paragraph in doc.paragraphs if paragraph.text == "引言")
            heading_l2 = next(paragraph for paragraph in doc.paragraphs if paragraph.text == "研究背景")
            heading_l3 = next(paragraph for paragraph in doc.paragraphs if paragraph.text == "研究假设")
            references_heading = next(paragraph for paragraph in doc.paragraphs if paragraph.text == "参考文献：")
            reference_entry = next(paragraph for paragraph in doc.paragraphs if paragraph.text.startswith("[1] 张三."))
            toc_heading = next(paragraph for paragraph in doc.paragraphs if paragraph.text == "目录")
            toc_field = next(
                paragraph for paragraph in doc.paragraphs
                if 'TOC \\o "1-3" \\h \\z \\u' in paragraph._element.xml
            )
            toc_page_break = next(
                paragraph for paragraph in doc.paragraphs
                if 'w:type="page"' in paragraph._element.xml
            )

            heading_l1_num_id = self.assert_paragraph_has_numbering(heading_l1, 0)
            heading_l2_num_id = self.assert_paragraph_has_numbering(heading_l2, 1)
            heading_l3_num_id = self.assert_paragraph_has_numbering(heading_l3, 2)

            self.assertEqual(heading_l3.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.LEFT)
            self.assertTrue(heading_l3.runs[0].font.bold)
            self.assertEqual(heading_l3.runs[0].font.size.pt, 12.0)
            self.assertEqual(heading_l3.paragraph_format.first_line_indent.pt, 0.0)
            self.assertEqual(heading_l3._element.pPr.find(qn("w:outlineLvl")).get(qn("w:val")), "2")
            self.assert_numbering_overrides(doc, heading_l1_num_id, {0: 1})
            self.assert_numbering_overrides(doc, heading_l2_num_id, {0: 1, 1: 1})
            self.assert_numbering_overrides(doc, heading_l3_num_id, {0: 1, 1: 1, 2: 1})

            self.assertEqual(references_heading.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertEqual(references_heading.text, "参考文献：")
            self.assertEqual(references_heading.runs[0].font.size.pt, 12.0)
            self.assertEqual(reference_entry.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.LEFT)
            self.assertAlmostEqual(reference_entry.paragraph_format.left_indent.pt, 0.0, places=1)
            self.assertAlmostEqual(reference_entry.paragraph_format.first_line_indent.pt, 21.0, places=1)
            self.assertEqual(reference_entry.paragraph_format.line_spacing, 1.0)
            self.assertEqual(reference_entry.runs[0].font.size.pt, 10.0)
            self.assertEqual(toc_heading.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertIn('TOC \\o "1-3" \\h \\z \\u', toc_field._element.xml)
            self.assertIn('w:type="page"', toc_page_break._element.xml)
            self.assertIn("w:updateFields", doc.settings.element.xml)
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

    def test_format_academic_paper_from_text_formats_english_abstract_keywords(self):
        text = (
            "跨境电商场景下供应链韧性研究\n"
            "Abstract\n"
            "This paper studies supply chain resilience under cross-border e-commerce settings.\n"
            "Keywords: supply chain resilience; cross-border e-commerce\n"
            "1 Introduction\n"
            "正文内容"
        )

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)

        try:
            summary = format_academic_paper_from_text(text, str(output_path))

            self.assertIsInstance(summary, dict)
            self.assertEqual(summary["stats"]["english_abstract_heading"], 1)
            self.assertEqual(summary["stats"]["english_abstract"], 1)
            self.assertEqual(summary["stats"]["english_keywords"], 1)

            doc = Document(str(output_path))
            abstract_heading = next(paragraph for paragraph in doc.paragraphs if paragraph.text == "Abstract")
            abstract_body = next(
                paragraph for paragraph in doc.paragraphs
                if paragraph.text.startswith("This paper studies supply chain resilience")
            )
            keywords = next(paragraph for paragraph in doc.paragraphs if paragraph.text.startswith("Keywords:"))
            introduction = next(paragraph for paragraph in doc.paragraphs if paragraph.text == "Introduction")

            self.assertEqual(abstract_heading.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertTrue(abstract_heading.runs[0].font.bold)
            self.assertEqual(abstract_heading.runs[0].font.size.pt, 12.0)

            self.assertEqual(abstract_body.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)
            self.assertEqual(abstract_body.paragraph_format.first_line_indent.pt, 0.0)
            self.assertEqual(abstract_body.runs[0].font.size.pt, 12.0)

            self.assertEqual(keywords.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.LEFT)
            self.assertTrue(keywords.runs[0].font.bold)
            self.assertEqual(keywords.text, "Keywords: supply chain resilience; cross-border e-commerce")
            self.assert_paragraph_has_numbering(introduction, 0)
        finally:
            output_path.unlink(missing_ok=True)

    def test_format_academic_paper_from_text_preserves_explicit_heading_number_offsets(self):
        text = (
            "数字化转型场景下企业韧性研究\n"
            "摘要：这是摘要内容\n"
            "关键词：数字化 韧性\n"
            "3 研究设计\n"
            "3.2 数据来源\n"
            "3.2.4 稳健性检验\n"
            "正文内容"
        )

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)

        try:
            summary = format_academic_paper_from_text(text, str(output_path))

            self.assertIsInstance(summary, dict)

            doc = Document(str(output_path))
            heading_l1 = next(paragraph for paragraph in doc.paragraphs if paragraph.text == "研究设计")
            heading_l2 = next(paragraph for paragraph in doc.paragraphs if paragraph.text == "数据来源")
            heading_l3 = next(paragraph for paragraph in doc.paragraphs if paragraph.text == "稳健性检验")

            heading_l1_num_id = self.assert_paragraph_has_numbering(heading_l1, 0)
            heading_l2_num_id = self.assert_paragraph_has_numbering(heading_l2, 1)
            heading_l3_num_id = self.assert_paragraph_has_numbering(heading_l3, 2)

            self.assert_numbering_overrides(doc, heading_l1_num_id, {0: 3})
            self.assert_numbering_overrides(doc, heading_l2_num_id, {0: 3, 1: 2})
            self.assert_numbering_overrides(doc, heading_l3_num_id, {0: 3, 1: 2, 2: 4})
        finally:
            output_path.unlink(missing_ok=True)

    def test_format_academic_paper_preserves_equation_paragraphs(self):
        omml = parse_xml(
            r"""
            <m:oMathPara %s>
              <m:oMath>
                <m:r>
                  <m:t>x=1</m:t>
                </m:r>
              </m:oMath>
            </m:oMathPara>
            """ % nsdecls("m")
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "equation_input.docx"
            output_path = temp_path / "equation_output.docx"

            doc = Document()
            doc.add_paragraph("含公式的论文标题")
            doc.add_paragraph("摘要：这是摘要内容")
            doc.add_paragraph("关键词：公式 测试")
            equation_paragraph = doc.add_paragraph()
            equation_paragraph._element.append(omml)
            doc.save(str(input_path))

            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertEqual(summary["equation_paragraphs"], 1)

            output_doc = Document(str(output_path))
            preserved_equation = output_doc.paragraphs[3]
            self.assertIn("oMath", preserved_equation._element.xml)
            self.assertEqual(preserved_equation.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertEqual(preserved_equation.paragraph_format.first_line_indent.pt, 0.0)

    def test_format_academic_paper_unifies_footnote_fonts_and_sizes(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "footnote_input.docx"
            output_path = temp_path / "footnote_output.docx"

            doc = Document()
            doc.add_paragraph("含脚注的论文标题")
            doc.add_paragraph("摘要：这是摘要内容")
            doc.add_paragraph("关键词：脚注 测试")
            doc.add_paragraph("正文里有一个脚注引用")
            doc.save(str(input_path))
            self.inject_simple_footnote(input_path, "脚注内容 Footnote 123")

            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertEqual(summary["formatted_footnotes"], 1)

            with ZipFile(output_path, "r") as archive:
                footnotes_xml = archive.read("word/footnotes.xml")

            footnotes_root = etree.fromstring(footnotes_xml)
            footnote = next(
                node
                for node in footnotes_root.findall(qn("w:footnote"))
                if node.get(qn("w:id")) == "2"
            )
            runs = footnote.findall(".//" + qn("w:r"))
            reference_rpr = runs[0].find(qn("w:rPr"))
            content_rpr = runs[1].find(qn("w:rPr"))

            self.assertEqual(reference_rpr.find(qn("w:rStyle")).get(qn("w:val")), "FootnoteReference")
            self.assertEqual(reference_rpr.find(qn("w:vertAlign")).get(qn("w:val")), "superscript")
            self.assertEqual(reference_rpr.find(qn("w:sz")).get(qn("w:val")), "20")
            self.assertEqual(content_rpr.find(qn("w:rFonts")).get(qn("w:eastAsia")), "宋体")
            self.assertEqual(content_rpr.find(qn("w:rFonts")).get(qn("w:ascii")), "Times New Roman")
            self.assertEqual(content_rpr.find(qn("w:sz")).get(qn("w:val")), "20")

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

            self.assertEqual(output_doc.paragraphs[3].text, "引言")
            self.assertEqual(output_doc.paragraphs[7].text, "图 1 系统架构图")
            self.assertEqual(output_doc.paragraphs[9].text, "图 2 模型流程图")
            self.assertEqual(output_doc.paragraphs[10].text, "表 1 样本描述统计")
            self.assertEqual(output_doc.paragraphs[5].style.name, "List Bullet")
            self.assertEqual(output_doc.paragraphs[6].paragraph_format.first_line_indent.pt, 0.0)
            self.assert_paragraph_has_numbering(output_doc.paragraphs[3], 0)

            body_run = output_doc.paragraphs[4].runs[0]
            table_run = output_doc.tables[0].cell(1, 1).paragraphs[0].runs[0]
            self.assert_run_uses_mixed_font_pair(self, body_run)
            self.assert_run_uses_mixed_font_pair(self, table_run)

            reference_entry = output_doc.paragraphs[12]
            self.assertAlmostEqual(reference_entry.paragraph_format.left_indent.pt, 0.0, places=1)
            self.assertAlmostEqual(reference_entry.paragraph_format.first_line_indent.pt, 21.0, places=1)
            self.assertEqual(reference_entry.paragraph_format.line_spacing, 1.0)
            self.assertEqual(reference_entry.runs[0].font.size.pt, 10.0)

            tbl_borders = output_doc.tables[0]._tbl.tblPr.find(qn("w:tblBorders"))
            top_left_borders = output_doc.tables[0].cell(0, 0)._tc.tcPr.find(qn("w:tcBorders"))
            bottom_left_borders = output_doc.tables[0].cell(1, 0)._tc.tcPr.find(qn("w:tcBorders"))
            self.assertIsNone(tbl_borders)
            self.assertEqual(top_left_borders.find(qn("w:top")).get(qn("w:val")), "single")
            self.assertEqual(top_left_borders.find(qn("w:bottom")).get(qn("w:val")), "single")
            self.assertEqual(top_left_borders.find(qn("w:left")).get(qn("w:val")), "none")
            self.assertEqual(top_left_borders.find(qn("w:right")).get(qn("w:val")), "none")
            self.assertEqual(bottom_left_borders.find(qn("w:top")).get(qn("w:val")), "none")
            self.assertEqual(bottom_left_borders.find(qn("w:bottom")).get(qn("w:val")), "single")

    def test_format_academic_paper_scales_oversized_images_to_page_width(self):
        tiny_png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            image_path = temp_path / "oversized.png"
            input_path = temp_path / "oversized_input.docx"
            output_path = temp_path / "oversized_output.docx"
            image_path.write_bytes(tiny_png)

            doc = Document()
            doc.add_paragraph("超宽图片版式测试")
            doc.add_paragraph("摘要：这是摘要内容。")
            doc.add_paragraph("关键词：图片 缩放 测试")

            picture = doc.add_paragraph()
            picture.alignment = WD_ALIGN_PARAGRAPH.CENTER
            picture.add_run().add_picture(str(image_path), width=Inches(8))
            doc.add_paragraph("图 1 超宽测试图")
            doc.save(str(input_path))

            summary = format_academic_paper(str(input_path), str(output_path))
            self.assertIsInstance(summary, dict)
            self.assertEqual(summary["resized_images"], 1)

            output_doc = Document(str(output_path))
            self.assertEqual(len(output_doc.inline_shapes), 1)

            printable_width = (
                int(output_doc.sections[0].page_width)
                - int(output_doc.sections[0].left_margin)
                - int(output_doc.sections[0].right_margin)
            )
            resized_shape = output_doc.inline_shapes[0]

            self.assertLessEqual(int(resized_shape.width), printable_width)
            self.assertEqual(int(resized_shape.width), int(resized_shape.height))
            self.assertLess(int(resized_shape.width), int(Inches(8)))

    def test_generate_cover_page_inserts_cover_table_and_page_break(self):
        info_dict = {
            "title": "企业数字化转型对绿色技术创新的影响研究",
            "cover_title": "《大数据挖掘》期末大作业",
            "college": "工商管理学院",
            "teacher": "刘璇",
            "class_name": "国商2301",
            "student_name": "何旻洋",
            "student_id": "2320100731",
        }

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)

        try:
            doc = Document()
            doc.add_paragraph("这是已经完成正文排版的第一页内容。")
            doc.add_paragraph("这是正文第二段。")
            apply_document_layout(doc, "这是正文运行页眉")

            self.assertTrue(generate_cover_page(doc, info_dict))
            doc.save(str(output_path))

            output_doc = Document(str(output_path))

            self.assertEqual(output_doc.paragraphs[0].text, "")
            self.assertEqual(output_doc.paragraphs[1].text, "")
            self.assertEqual(output_doc.paragraphs[2].text, info_dict["cover_title"])
            self.assertEqual(output_doc.paragraphs[3].text, "")
            self.assertEqual(output_doc.paragraphs[5].text, "这是已经完成正文排版的第一页内容。")
            self.assertIn('w:type="page"', output_doc.paragraphs[4]._element.xml)
            self.assertEqual(len(output_doc.inline_shapes), 2)

            title_run = output_doc.paragraphs[2].runs[0]
            self.assertEqual(output_doc.paragraphs[2].paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertTrue(title_run.font.bold)
            self.assertEqual(title_run.font.size.pt, 26.0)

            cover_table = output_doc.tables[0]
            self.assertEqual(len(cover_table.rows), 5)
            self.assertEqual(cover_table.cell(0, 0).text, "学院")
            self.assertEqual(cover_table.cell(0, 1).text, "工商管理学院")
            self.assertEqual(cover_table.cell(4, 1).text, "2320100731")

            label_borders = cover_table.cell(0, 0)._tc.tcPr.find(qn("w:tcBorders"))
            value_borders = cover_table.cell(0, 1)._tc.tcPr.find(qn("w:tcBorders"))
            self.assertEqual(label_borders.find(qn("w:bottom")).get(qn("w:val")), "none")
            self.assertEqual(value_borders.find(qn("w:top")).get(qn("w:val")), "none")
            self.assertEqual(value_borders.find(qn("w:left")).get(qn("w:val")), "none")
            self.assertEqual(value_borders.find(qn("w:right")).get(qn("w:val")), "none")
            self.assertEqual(value_borders.find(qn("w:bottom")).get(qn("w:val")), "single")

            paragraph_borders = cover_table.cell(0, 1).paragraphs[0]._element.pPr.find(qn("w:pBdr"))
            self.assertIsNotNone(paragraph_borders)
            self.assertEqual(paragraph_borders.find(qn("w:bottom")).get(qn("w:val")), "single")

            label_run = cover_table.cell(0, 0).paragraphs[0].runs[0]
            value_run = cover_table.cell(4, 1).paragraphs[0].runs[0]
            self.assert_run_uses_mixed_font_pair(self, label_run)
            self.assert_run_uses_mixed_font_pair(self, value_run)

            section = output_doc.sections[0]
            self.assertTrue(section.different_first_page_header_footer)
            self.assertEqual(section.first_page_header.paragraphs[0].text, "")
            self.assertEqual(section.first_page_footer.paragraphs[0].text, "")
            self.assertEqual(section.header.paragraphs[0].text, "这是正文运行页眉")
        finally:
            output_path.unlink(missing_ok=True)

    def test_generate_cover_page_skips_when_title_is_missing(self):
        doc = Document()
        doc.add_paragraph("正文内容")

        self.assertFalse(generate_cover_page(doc, {"student_name": "何旻洋"}))
        self.assertEqual([paragraph.text for paragraph in doc.paragraphs], ["正文内容"])

    def test_ensure_document_ends_with_page_break_appends_break(self):
        doc = Document()
        doc.add_paragraph("封面内容")

        ensure_document_ends_with_page_break(doc)

        self.assertTrue(doc.paragraphs)
        self.assertIn('w:type="page"', doc.paragraphs[-1]._element.xml)


if __name__ == "__main__":
    unittest.main()
