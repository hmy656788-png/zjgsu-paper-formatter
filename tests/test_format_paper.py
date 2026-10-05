import base64
import re
import tempfile
import unittest
from dataclasses import replace
from io import BytesIO
from pathlib import Path
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.oxml.ns import nsdecls
from docx.shared import Cm, Inches
from docxcompose.properties import CustomProperties
from lxml import etree

import format_paper as format_paper_module
from format_paper import (
    apply_document_layout,
    concatenate_documents,
    DocumentConcatError,
    OutputSizeLimitExceeded,
    ensure_document_ends_with_page_break,
    find_title_paragraph_index,
    format_academic_paper,
    format_academic_paper_from_text,
    generate_cover_page,
    merge_cover_and_body,
    split_text_to_paragraphs,
)


class FormatPaperFromTextTestCase(unittest.TestCase):
    CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
    PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    FOOTNOTE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
    ENDNOTE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
    CHART_STYLE_REQUIRED_ENTRIES = (
        "axisTitle",
        "categoryAxis",
        "chartArea",
        "dataLabel",
        "dataPoint",
        "dataPoint3D",
        "dataPointLine",
        "dataPointMarker",
        "dataPointWireframe",
        "dataTable",
        "downBar",
        "dropLine",
        "errorBar",
        "floor",
        "gridlineMajor",
        "gridlineMinor",
        "hiLoLine",
        "leaderLine",
        "legend",
        "plotArea",
        "plotArea3D",
        "seriesAxis",
        "seriesLine",
        "title",
        "trendline",
        "trendlineLabel",
        "upBar",
        "valueAxis",
        "wall",
    )

    @staticmethod
    def assert_run_uses_mixed_font_pair(test_case, run):
        r_fonts = run._element.rPr.rFonts
        test_case.assertEqual(r_fonts.get(qn("w:eastAsia")), "宋体")
        test_case.assertEqual(r_fonts.get(qn("w:ascii")), "Times New Roman")
        test_case.assertEqual(r_fonts.get(qn("w:hAnsi")), "Times New Roman")

    def assert_xml_children_follow_order(self, element, ordered_names):
        order = {qn(name): index for index, name in enumerate(ordered_names)}
        observed = [order[child.tag] for child in element if child.tag in order]
        self.assertEqual(observed, sorted(observed), etree.tostring(element, encoding="unicode"))

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

    @staticmethod
    def customize_test_theme(
        document,
        *,
        accent1,
        major_latin,
        format_name,
        background_mapping,
    ):
        drawing_namespace = (
            "http://schemas.openxmlformats.org/drawingml/2006/main"
        )
        theme_part = document.part.part_related_by(
            format_paper_module.RT.THEME
        )
        theme_root = etree.fromstring(theme_part.blob)
        accent = theme_root.find(
            ".//a:clrScheme/a:accent1/*",
            namespaces={"a": drawing_namespace},
        )
        accent.tag = f"{{{drawing_namespace}}}srgbClr"
        accent.attrib.clear()
        accent.set("val", accent1)
        theme_root.find(
            ".//a:fontScheme/a:majorFont/a:latin",
            namespaces={"a": drawing_namespace},
        ).set("typeface", major_latin)
        theme_root.find(
            ".//a:fmtScheme",
            namespaces={"a": drawing_namespace},
        ).set("name", format_name)
        theme_part._blob = etree.tostring(
            theme_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        color_mapping = document.settings.element.find(
            qn("w:clrSchemeMapping")
        )
        color_mapping.set(qn("w:bg1"), background_mapping)

    @staticmethod
    def make_test_theme_override_blob(
        document,
        *,
        included_schemes=("clrScheme", "fontScheme", "fmtScheme"),
        accent1=None,
    ):
        drawing_namespace = (
            "http://schemas.openxmlformats.org/drawingml/2006/main"
        )
        theme_part = document.part.part_related_by(
            format_paper_module.RT.THEME
        )
        theme_root = etree.fromstring(theme_part.blob)
        override_root = etree.Element(
            f"{{{drawing_namespace}}}themeOverride",
            nsmap={"a": drawing_namespace},
        )
        for scheme_name in included_schemes:
            source_scheme = theme_root.find(
                f".//a:{scheme_name}",
                namespaces={"a": drawing_namespace},
            )
            copied_scheme = etree.fromstring(etree.tostring(source_scheme))
            if scheme_name == "clrScheme" and accent1 is not None:
                color = copied_scheme.find(
                    "a:accent1/*",
                    namespaces={"a": drawing_namespace},
                )
                color.tag = f"{{{drawing_namespace}}}srgbClr"
                color.attrib.clear()
                color.set("val", accent1)
            override_root.append(copied_scheme)
        return etree.tostring(
            override_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    @classmethod
    def make_test_chart_style_blob(cls):
        namespace = format_paper_module._CHART_STYLE_NAMESPACE
        drawing_namespace = format_paper_module._THEME_NAMESPACE
        root = etree.Element(
            f"{{{namespace}}}chartStyle",
            nsmap={"cs": namespace, "a": drawing_namespace},
            id="201",
        )
        for entry_name in cls.CHART_STYLE_REQUIRED_ENTRIES:
            entry = etree.SubElement(root, f"{{{namespace}}}{entry_name}")
            etree.SubElement(
                entry,
                f"{{{namespace}}}lnRef",
                idx="0",
            )
            etree.SubElement(
                entry,
                f"{{{namespace}}}fillRef",
                idx="0",
            )
            etree.SubElement(
                entry,
                f"{{{namespace}}}effectRef",
                idx="0",
            )
            font_reference = etree.SubElement(
                entry,
                f"{{{namespace}}}fontRef",
                idx="minor",
            )
            etree.SubElement(
                font_reference,
                f"{{{drawing_namespace}}}schemeClr",
                val="tx1",
            )
        return etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    @staticmethod
    def make_test_chart_color_style_blob():
        namespace = format_paper_module._CHART_STYLE_NAMESPACE
        drawing_namespace = format_paper_module._THEME_NAMESPACE
        root = etree.Element(
            f"{{{namespace}}}colorStyle",
            nsmap={"cs": namespace, "a": drawing_namespace},
            meth="cycle",
            id="10",
        )
        etree.SubElement(
            root,
            f"{{{drawing_namespace}}}schemeClr",
            val="accent1",
        )
        return etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    @staticmethod
    def add_test_chart(
        document,
        label,
        *,
        extended=False,
        override_specs=(),
        add_sidecar=False,
        sidecar_specs=None,
    ):
        chart_namespace = (
            format_paper_module._CHARTEX_NAMESPACE
            if extended
            else format_paper_module._CHART_NAMESPACE
        )
        chart_prefix = "cx" if extended else "c"
        if extended:
            chart_xml = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<cx:chartSpace xmlns:cx="{chart_namespace}" '
                f'xmlns:a="{format_paper_module._THEME_NAMESPACE}">'
                '<cx:chartData><cx:data id="0"/></cx:chartData>'
                '<cx:chart><cx:plotArea><cx:plotAreaRegion>'
                '<cx:series/></cx:plotAreaRegion></cx:plotArea></cx:chart>'
                '<cx:spPr>'
                '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>'
                '</cx:spPr></cx:chartSpace>'
            ).encode("utf-8")
            relationship_type = format_paper_module._CHARTEX_RELATIONSHIP_TYPE
            content_type = format_paper_module._CHARTEX_CONTENT_TYPE
        else:
            chart_xml = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<c:chartSpace xmlns:c="{chart_namespace}" '
                f'xmlns:a="{format_paper_module._THEME_NAMESPACE}">'
                '<c:chart><c:plotArea><c:layout/><c:pieChart>'
                '<c:varyColors val="0"/>'
                '</c:pieChart></c:plotArea></c:chart><c:spPr>'
                '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>'
                '</c:spPr></c:chartSpace>'
            ).encode("utf-8")
            relationship_type = format_paper_module.RT.CHART
            content_type = format_paper_module.CT.DML_CHART

        chart_part = format_paper_module.Part(
            format_paper_module.PackURI(
                "/word/charts/chartEx1.xml"
                if extended
                else "/word/charts/chart1.xml"
            ),
            content_type,
            chart_xml,
            document.part.package,
        )
        chart_rid = document.part.relate_to(
            chart_part,
            relationship_type,
        )
        paragraph = document.add_paragraph(label)
        paragraph._element.append(
            parse_xml(
                f'<w:r {nsdecls("w", "a", "r")} '
                f'xmlns:{chart_prefix}="{chart_namespace}"><w:drawing>'
                '<a:graphic><a:graphicData uri="urn:test:chart">'
                f'<{chart_prefix}:chart r:id="{chart_rid}"/>'
                '</a:graphicData></a:graphic>'
                '</w:drawing></w:r>'
            )
        )

        for index, (blob, part_content_type, external) in enumerate(
            override_specs,
            start=1,
        ):
            if external:
                chart_part.rels.add_relationship(
                    format_paper_module.RT.THEME_OVERRIDE,
                    f"https://example.test/theme-override-{index}.xml",
                    f"rIdTheme{index}",
                    is_external=True,
                )
                continue
            override_part = format_paper_module.Part(
                format_paper_module.PackURI(
                    f"/word/theme/themeOverride{index}.xml"
                ),
                part_content_type,
                blob,
                document.part.package,
            )
            chart_part.rels.add_relationship(
                format_paper_module.RT.THEME_OVERRIDE,
                override_part,
                f"rIdTheme{index}",
            )

        if add_sidecar or sidecar_specs is not None:
            if sidecar_specs is None:
                sidecar_specs = (
                    (
                        format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                        FormatPaperFromTextTestCase.make_test_chart_style_blob(),
                        format_paper_module._CHART_STYLE_CONTENT_TYPE,
                        False,
                    ),
                    (
                        format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                        FormatPaperFromTextTestCase.make_test_chart_color_style_blob(),
                        format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                        False,
                    ),
                )
            for index, (
                sidecar_relationship_type,
                sidecar_blob,
                sidecar_content_type,
                external,
            ) in enumerate(sidecar_specs, start=1):
                relationship_id = f"rIdSidecar{index}"
                if external:
                    chart_part.rels.add_relationship(
                        sidecar_relationship_type,
                        f"https://example.test/chart-sidecar-{index}.xml",
                        relationship_id,
                        is_external=True,
                    )
                    continue
                sidecar_part = format_paper_module.Part(
                    format_paper_module.PackURI(
                        f"/word/charts/sidecar{index}.xml"
                    ),
                    sidecar_content_type,
                    sidecar_blob,
                    document.part.package,
                )
                chart_part.rels.add_relationship(
                    sidecar_relationship_type,
                    sidecar_part,
                    relationship_id,
                )
        return chart_part

    def inject_simple_footnote(
        self,
        docx_path: Path,
        footnote_text: str,
        footnote_id: int = 2,
        footnote_content_xml: str | None = None,
    ):
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
                f'<w:footnoteReference w:id="{footnote_id}"/>'
                "</w:r>"
            )
        )
        payloads["word/document.xml"] = etree.tostring(
            document_xml,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        if footnote_content_xml is None:
            footnote_content_xml = (
                f"<w:r><w:t xml:space=\"preserve\"> {footnote_text}</w:t></w:r>"
            )

        payloads["word/footnotes.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:footnotes {nsdecls("w")}>'
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
            '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
            f'<w:footnote w:id="{footnote_id}"><w:p>'
            '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
            f"{footnote_content_xml}"
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

    def inject_simple_endnote(
        self,
        docx_path: Path,
        endnote_text: str,
        endnote_id: int = 2,
        separator_style: str | None = None,
    ):
        with ZipFile(docx_path, "r") as source:
            entries = source.infolist()
            payloads = {entry.filename: source.read(entry.filename) for entry in entries}

        content_types = etree.fromstring(payloads["[Content_Types].xml"])
        override_tag = f"{{{self.CONTENT_TYPES_NS}}}Override"
        override = etree.Element(override_tag)
        override.set("PartName", "/word/endnotes.xml")
        override.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
        )
        content_types.append(override)
        payloads["[Content_Types].xml"] = etree.tostring(
            content_types,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        relationships = etree.fromstring(payloads["word/_rels/document.xml.rels"])
        relationship_tag = f"{{{self.PACKAGE_REL_NS}}}Relationship"
        relationship = etree.Element(relationship_tag)
        relationship.set("Id", "rIdEndnotes")
        relationship.set("Type", self.ENDNOTE_REL_TYPE)
        relationship.set("Target", "endnotes.xml")
        relationships.append(relationship)
        payloads["word/_rels/document.xml.rels"] = etree.tostring(
            relationships,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        document_xml = etree.fromstring(payloads["word/document.xml"])
        target_paragraph = document_xml.find(qn("w:body")).findall(qn("w:p"))[-1]
        target_paragraph.append(
            parse_xml(
                f'<w:r {nsdecls("w")}>'
                f'<w:endnoteReference w:id="{endnote_id}"/>'
                "</w:r>"
            )
        )
        payloads["word/document.xml"] = etree.tostring(
            document_xml,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        separator_properties = (
            f'<w:pPr><w:pStyle w:val="{separator_style}"/></w:pPr>'
            if separator_style
            else ""
        )
        payloads["word/endnotes.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:endnotes {nsdecls("w")}>'
            '<w:endnote w:type="separator" w:id="-1"><w:p>'
            f"{separator_properties}"
            '<w:r><w:separator/></w:r></w:p></w:endnote>'
            '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>'
            f'<w:endnote w:id="{endnote_id}"><w:p>'
            f'<w:r><w:t>{endnote_text}</w:t></w:r>'
            "</w:p></w:endnote>"
            "</w:endnotes>"
        ).encode("utf-8")

        with ZipFile(docx_path, "w") as target:
            for entry in entries:
                target.writestr(entry, payloads[entry.filename])
            target.writestr("word/endnotes.xml", payloads["word/endnotes.xml"])

    def relocate_simple_footnote_part(
        self,
        docx_path: Path,
        relocated_name: str = "word/notes/footnotes2.xml",
    ):
        with ZipFile(docx_path, "r") as source:
            entries = source.infolist()
            payloads = {entry.filename: source.read(entry.filename) for entry in entries}

        original_name = "word/footnotes.xml"
        content_types = etree.fromstring(payloads["[Content_Types].xml"])
        override_tag = f"{{{self.CONTENT_TYPES_NS}}}Override"
        footnote_override = next(
            node
            for node in content_types.findall(override_tag)
            if node.get("PartName") == f"/{original_name}"
        )
        footnote_override.set("PartName", f"/{relocated_name}")
        payloads["[Content_Types].xml"] = etree.tostring(
            content_types,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        relationships = etree.fromstring(payloads["word/_rels/document.xml.rels"])
        relationship_tag = f"{{{self.PACKAGE_REL_NS}}}Relationship"
        footnote_relationship = next(
            node
            for node in relationships.findall(relationship_tag)
            if node.get("Type") == self.FOOTNOTE_REL_TYPE
        )
        footnote_relationship.set("Target", relocated_name.removeprefix("word/"))
        payloads["word/_rels/document.xml.rels"] = etree.tostring(
            relationships,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        payloads[relocated_name] = payloads.pop(original_name)
        with ZipFile(docx_path, "w") as target:
            for entry in entries:
                if entry.filename == original_name:
                    target.writestr(relocated_name, payloads[relocated_name])
                else:
                    target.writestr(entry, payloads[entry.filename])

    @staticmethod
    def rewrite_docx_member(docx_path: Path, member_name: str, transform):
        with ZipFile(docx_path, "r") as source:
            entries = source.infolist()
            payloads = {entry.filename: source.read(entry.filename) for entry in entries}

        payloads[member_name] = transform(payloads[member_name])
        with ZipFile(docx_path, "w") as target:
            for entry in entries:
                target.writestr(entry, payloads[entry.filename])

    def test_formatter_preserves_special_hyphens_and_structural_breaks(self):
        for tag, attributes, visible in (
            ("w:noBreakHyphen", {}, "\u2011"),
            ("w:softHyphen", {}, "\u00ad"),
            ("w:br", {"w:type": "page"}, "\n"),
            ("w:br", {"w:type": "column"}, "\n"),
            ("w:br", {"w:clear": "all"}, "\n"),
        ):
            with self.subTest(tag=tag, attributes=attributes), tempfile.TemporaryDirectory() as temp_dir:
                document = Document()
                document.add_paragraph("行内特殊字符保留测试")
                abstract = document.add_paragraph("摘要：成本")
                special = OxmlElement(tag)
                for name, value in attributes.items():
                    special.set(qn(name), value)
                abstract.add_run()._r.append(special)
                abstract.add_run("收益分析。")
                document.add_paragraph("关键词：排版 内容")
                document.add_paragraph("正文内容。")
                source = Path(temp_dir) / "input.docx"
                output = Path(temp_dir) / "output.docx"
                document.save(source)

                self.assertTrue(format_academic_paper(str(source), str(output)))
                result = Document(output)
                result_abstract = next(p for p in result.paragraphs if p.text.startswith("摘要："))
                preserved = result_abstract._element.find(".//" + qn(tag))
                self.assertIsNotNone(preserved)
                for name, value in attributes.items():
                    self.assertEqual(preserved.get(qn(name)), value)
                self.assertEqual(
                    format_paper_module._paragraph_text_for_matching(result_abstract),
                    "摘要：成本" + visible + "收益分析。",
                )

    def test_regular_text_wrapping_break_remains_rewrite_safe(self):
        paragraph = Document().add_paragraph("摘要：正文")
        paragraph.add_run()._r.append(OxmlElement("w:br"))
        paragraph.add_run("续行")
        self.assertFalse(format_paper_module._has_rewrite_sensitive_inline_content(paragraph))

    def test_split_text_to_paragraphs_normalizes_line_endings(self):
        text = "第一段\r\n\r\n第二段\r第三段\n"

        self.assertEqual(
            split_text_to_paragraphs(text),
            ["第一段", "", "第二段", "第三段"],
        )

    def test_text_paragraph_limit_accepts_exact_default_boundary(self):
        text = "\n".join(["x"] * format_paper_module.MAX_TEXT_PARAGRAPHS)

        self.assertFalse(format_paper_module.text_exceeds_paragraph_limit(text))
        self.assertEqual(
            len(split_text_to_paragraphs(text)),
            format_paper_module.MAX_TEXT_PARAGRAPHS,
        )

    def test_text_paragraph_limit_rejects_one_over_default_boundary(self):
        text = "\n".join(["x"] * (format_paper_module.MAX_TEXT_PARAGRAPHS + 1))

        self.assertTrue(format_paper_module.text_exceeds_paragraph_limit(text))

    def test_text_paragraph_limit_handles_mixed_newlines_and_trailing_breaks(self):
        cases = (
            ("a\r\nb\rc\nd", 3, True),
            ("a\r\nb\rc\nd", 4, False),
            ("\r\n\rtext", 3, False),
            ("a\r\n\r\n\n", 1, False),
            ("a\r\n\r\n\nb", 3, True),
        )

        for text, max_paragraphs, expected in cases:
            with self.subTest(text=repr(text), max_paragraphs=max_paragraphs):
                self.assertEqual(
                    format_paper_module.text_exceeds_paragraph_limit(
                        text,
                        max_paragraphs=max_paragraphs,
                    ),
                    expected,
                )

    def test_text_paragraph_limit_rejects_invalid_maximum(self):
        for invalid_maximum in (0, -1, True, False, 1.5, "1", None):
            with self.subTest(max_paragraphs=invalid_maximum):
                with self.assertRaisesRegex(ValueError, "positive integer"):
                    format_paper_module.text_exceeds_paragraph_limit(
                        "text",
                        max_paragraphs=invalid_maximum,
                    )

    def test_split_text_to_paragraphs_raises_when_limit_is_exceeded(self):
        with self.assertRaises(format_paper_module.TextParagraphLimitExceeded) as caught:
            split_text_to_paragraphs("one\ntwo\nthree", max_paragraphs=2)

        self.assertEqual(caught.exception.max_paragraphs, 2)

    def test_format_from_text_rejects_paragraph_amplification_before_document_creation(self):
        text = "\n".join(["x"] * (format_paper_module.MAX_TEXT_PARAGRAPHS + 1))

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "too-many-paragraphs.docx"
            with (
                patch.object(format_paper_module, "Document") as document_factory,
                patch.object(format_paper_module, "_run_document_processing") as processor,
            ):
                result = format_academic_paper_from_text(text, str(output_path))

            self.assertFalse(result)
            document_factory.assert_not_called()
            processor.assert_not_called()
            self.assertFalse(output_path.exists())

    def test_format_log_path_sanitizes_paths_and_control_chars(self):
        long_name = "a" * (format_paper_module.MAX_LOG_PATH_LENGTH + 20) + ".docx"

        self.assertEqual(format_paper_module.format_log_path("..\\secret\r\nX.docx"), "secretX.docx")
        self.assertEqual(format_paper_module.format_log_path("/tmp/private/job-input.docx"), "job-input.docx")
        self.assertEqual(format_paper_module.format_log_path(None), "unknown")
        self.assertEqual(len(format_paper_module.format_log_path(long_name)), format_paper_module.MAX_LOG_PATH_LENGTH)

    def test_bounded_output_stream_limits_extent_not_cumulative_writes(self):
        class SpyBytesIO(BytesIO):
            def __init__(self):
                super().__init__()
                self.forwarded_writes = []

            def write(self, data):
                self.forwarded_writes.append(bytes(data))
                return super().write(data)

        raw = SpyBytesIO()
        stream = format_paper_module._BoundedOutputStream(raw, 5)
        self.assertEqual(stream.write(b"abc"), 3)
        stream.seek(0)
        self.assertEqual(stream.write(b"XY"), 2)
        stream.seek(3)

        with self.assertRaises(OutputSizeLimitExceeded):
            stream.write(b"123")

        self.assertEqual(raw.getvalue(), b"XYc")
        self.assertEqual(raw.forwarded_writes, [b"abc", b"XY"])
        self.assertTrue(stream.limit_exceeded)
        self.assertEqual(stream.write(b"discarded cleanup"), len(b"discarded cleanup"))
        self.assertEqual(raw.getvalue(), b"XYc")

    def test_output_limit_rejects_numeric_lookalikes(self):
        for invalid in (True, False, 1.5, "1024", b"1024"):
            with self.subTest(invalid=invalid):
                with self.assertRaises(ValueError):
                    format_paper_module._normalize_output_limit(invalid)
        self.assertEqual(format_paper_module._normalize_output_limit(1024), 1024)

    def test_save_with_output_limit_enforces_exact_zip_boundary(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            baseline_path = temp_path / "baseline.docx"
            exact_path = temp_path / "exact.docx"
            short_path = temp_path / "short.docx"
            doc = Document()
            doc.add_paragraph("边界测试")

            exact_size = format_paper_module._save_with_output_limit(
                doc.save,
                baseline_path,
            )
            saved_size = format_paper_module._save_with_output_limit(
                doc.save,
                exact_path,
                max_output_bytes=exact_size,
            )

            self.assertEqual(saved_size, exact_size)
            with ZipFile(exact_path) as archive:
                self.assertIsNone(archive.testzip())

            with self.assertRaises(OutputSizeLimitExceeded):
                format_paper_module._save_with_output_limit(
                    doc.save,
                    short_path,
                    max_output_bytes=exact_size - 1,
                )
            self.assertFalse(short_path.exists())

    def test_save_with_output_limit_preserves_existing_target_on_failure(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "existing.docx"
            sentinel = b"ORIGINAL-SENTINEL"
            output_path.write_bytes(sentinel)

            doc = Document()
            doc.add_paragraph("输出上限失败")
            with self.assertRaises(OutputSizeLimitExceeded):
                format_paper_module._save_with_output_limit(
                    doc.save,
                    output_path,
                    max_output_bytes=1024,
                )
            self.assertEqual(output_path.read_bytes(), sentinel)

            def failing_save(destination):
                Path(destination).write_bytes(b"partial")
                raise RuntimeError("save failed")

            with self.assertRaises(RuntimeError):
                format_paper_module._save_with_output_limit(
                    failing_save,
                    output_path,
                )
            self.assertEqual(output_path.read_bytes(), sentinel)

    def test_save_with_output_validation_preserves_existing_target(self):
        """结构校验必须发生在原子替换前，拒绝坏包且保留旧产物。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "existing.docx"
            sentinel = b"ORIGINAL-SENTINEL"
            output_path.write_bytes(sentinel)

            def save_invalid(destination):
                Path(destination).write_bytes(b"not-a-docx")

            with self.assertRaises(format_paper_module.InvalidGeneratedDocxError):
                format_paper_module._save_with_output_limit(
                    save_invalid,
                    output_path,
                    validate_callable=lambda path: False,
                )
            self.assertEqual(output_path.read_bytes(), sentinel)

    def test_save_with_output_limit_leases_staging_path_until_publish(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "leased.docx"
            leases_before = format_paper_module.get_active_temp_paths()
            observed_staging_paths = []

            def save_bytes(destination):
                staging_path = Path(destination)
                observed_staging_paths.append(staging_path)
                self.assertTrue(format_paper_module.is_active_temp_path(staging_path))
                self.assertIn(staging_path, format_paper_module.get_active_temp_paths())
                staging_path.write_bytes(b"complete-document")

            format_paper_module._save_with_output_limit(save_bytes, output_path)

            self.assertEqual(output_path.read_bytes(), b"complete-document")
            self.assertEqual(len(observed_staging_paths), 1)
            self.assertFalse(observed_staging_paths[0].exists())
            self.assertEqual(format_paper_module.get_active_temp_paths(), leases_before)

    def test_save_with_output_limit_rejects_replaced_staging_symlink(self):
        """发布前不得跟随保存回调替换出的外部符号链接。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_path = temp_path / "existing.docx"
            output_path.write_bytes(b"ORIGINAL")
            external_path = temp_path / "external.docx"
            external_path.write_bytes(b"EXTERNAL")

            def replace_with_symlink(destination):
                staging_path = Path(destination)
                staging_path.unlink()
                staging_path.symlink_to(external_path)

            with self.assertRaises(format_paper_module.InvalidGeneratedDocxError):
                format_paper_module._save_with_output_limit(
                    replace_with_symlink,
                    output_path,
                )
            self.assertEqual(output_path.read_bytes(), b"ORIGINAL")
            self.assertEqual(external_path.read_bytes(), b"EXTERNAL")

    def test_save_with_output_limit_rejects_symlinked_output_parent(self):
        """暂存目录不得通过符号链接跳出调用方指定的输出根目录。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            real_parent = temp_path / "real-output"
            real_parent.mkdir()
            linked_parent = temp_path / "linked-output"
            linked_parent.symlink_to(real_parent, target_is_directory=True)
            with self.assertRaises(format_paper_module.InvalidGeneratedDocxError):
                format_paper_module._save_with_output_limit(
                    lambda destination: Path(destination).write_bytes(b"data"),
                    linked_parent / "output.docx",
                )
            self.assertFalse((real_parent / "output.docx").exists())

    def test_save_with_output_limit_rejects_parent_swap_during_tempfile_creation(self):
        """暂存文件创建期间父目录换 inode 时不得越过输出边界。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_parent = temp_path / "output"
            output_parent.mkdir()
            external_parent = temp_path / "external"
            external_parent.mkdir()
            output_path = output_parent / "result.docx"
            real_named_temporary_file = tempfile.NamedTemporaryFile

            def swapping_named_temporary_file(*args, **kwargs):
                context = real_named_temporary_file(*args, **kwargs)

                class SwappingContext:
                    def __enter__(self):
                        handle = context.__enter__()
                        backup = temp_path / "output-backup"
                        output_parent.rename(backup)
                        output_parent.symlink_to(external_parent, target_is_directory=True)
                        self.backup = backup
                        return handle

                    def __exit__(self, *exc_info):
                        output_parent.unlink(missing_ok=True)
                        self.backup.rename(output_parent)
                        return context.__exit__(*exc_info)

                return SwappingContext()

            with patch.object(
                format_paper_module.tempfile,
                "NamedTemporaryFile",
                swapping_named_temporary_file,
            ):
                with self.assertRaises(format_paper_module.InvalidGeneratedDocxError):
                    format_paper_module._save_with_output_limit(
                        lambda destination: Path(destination).write_bytes(b"data"),
                        output_path,
                    )

            self.assertFalse((external_parent / "result.docx").exists())
            self.assertFalse(output_path.exists())

    def test_parent_swap_cleanup_never_unlinks_through_remaining_symlink(self):
        """异常清理不得沿仍存在的父目录符号链接删除外部文件。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_parent = temp_path / "output"
            output_parent.mkdir()
            external_parent = temp_path / "external"
            external_parent.mkdir()
            sentinel = external_parent / "sentinel.docx"
            sentinel.write_bytes(b"KEEP")
            output_path = output_parent / "result.docx"
            real_named_temporary_file = tempfile.NamedTemporaryFile
            created_staging = {}
            backup = temp_path / "output-backup"

            def swapping_named_temporary_file(*args, **kwargs):
                context = real_named_temporary_file(*args, **kwargs)

                class SwappingContext:
                    def __enter__(self):
                        handle = context.__enter__()
                        created_staging["path"] = Path(handle.name)
                        output_parent.rename(backup)
                        output_parent.symlink_to(
                            external_parent, target_is_directory=True
                        )
                        return handle

                    def __exit__(self, *exc_info):
                        return context.__exit__(*exc_info)

                return SwappingContext()

            try:
                with patch.object(
                    format_paper_module.tempfile,
                    "NamedTemporaryFile",
                    swapping_named_temporary_file,
                ):
                    with self.assertRaises(
                        format_paper_module.InvalidGeneratedDocxError
                    ):
                        format_paper_module._save_with_output_limit(
                            lambda destination: Path(destination).write_bytes(
                                b"data"
                            ),
                            output_path,
                        )

                self.assertTrue(sentinel.exists())
                self.assertTrue(
                    (backup / created_staging["path"].name).exists()
                )
            finally:
                output_parent.unlink(missing_ok=True)
                backup.rename(output_parent)
                staged_path = created_staging.get("path")
                if staged_path is not None:
                    staged_path.unlink(missing_ok=True)
            self.assertEqual(list(output_parent.iterdir()), [])

    def test_save_with_output_limit_rejects_replaced_staging_hardlink(self):
        """发布前不得接受指向外部 inode 的硬链接暂存文件。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_path = temp_path / "existing.docx"
            output_path.write_bytes(b"ORIGINAL")
            external_path = temp_path / "external.docx"
            external_path.write_bytes(b"EXTERNAL")

            def replace_with_hardlink(destination):
                staging_path = Path(destination)
                staging_path.unlink()
                staging_path.hardlink_to(external_path)

            with self.assertRaises(format_paper_module.InvalidGeneratedDocxError):
                format_paper_module._save_with_output_limit(
                    replace_with_hardlink,
                    output_path,
                )
            self.assertEqual(output_path.read_bytes(), b"ORIGINAL")
            self.assertEqual(external_path.read_bytes(), b"EXTERNAL")

    def test_save_with_output_limit_does_not_mask_replaced_staging_directory(self):
        """异常清理不得因回调替换成目录而覆盖真实校验错误。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_path = temp_path / "existing.docx"
            output_path.write_bytes(b"ORIGINAL")
            replaced_directory = None

            def replace_with_directory(destination):
                nonlocal replaced_directory
                staging_path = Path(destination)
                staging_path.unlink()
                replaced_directory = staging_path
                staging_path.mkdir()

            with self.assertRaises(format_paper_module.InvalidGeneratedDocxError):
                format_paper_module._save_with_output_limit(
                    replace_with_directory,
                    output_path,
                )
            self.assertEqual(output_path.read_bytes(), b"ORIGINAL")
            self.assertIsNotNone(replaced_directory)
            self.assertTrue(replaced_directory.is_dir())
            self.assertFalse(
                format_paper_module.is_active_temp_path(replaced_directory)
            )

    def test_save_with_output_validation_rechecks_staging_before_publish(self):
        """校验回调替换暂存路径后不得发布外部符号链接。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_path = temp_path / "existing.docx"
            output_path.write_bytes(b"ORIGINAL")
            external_path = temp_path / "external.docx"
            external_path.write_bytes(b"EXTERNAL")

            def validator(staging_path):
                staging_path = Path(staging_path)
                staging_path.unlink()
                staging_path.symlink_to(external_path)
                return True

            with self.assertRaises(format_paper_module.InvalidGeneratedDocxError):
                format_paper_module._save_with_output_limit(
                    lambda destination: Path(destination).write_bytes(b"GENERATED"),
                    output_path,
                    validate_callable=validator,
                )

            self.assertEqual(output_path.read_bytes(), b"ORIGINAL")
            self.assertEqual(external_path.read_bytes(), b"EXTERNAL")

    def test_active_temp_path_registry_reference_counts_nested_leases(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            leased_path = Path(temp_dir) / "nested.tmp"
            leases_before = format_paper_module.get_active_temp_paths()

            normalized = format_paper_module.register_active_temp_path(leased_path)
            format_paper_module.register_active_temp_path(leased_path)
            self.assertTrue(format_paper_module.is_active_temp_path(leased_path))
            self.assertIn(normalized, format_paper_module.get_active_temp_paths())

            format_paper_module.release_active_temp_path(leased_path)
            self.assertTrue(format_paper_module.is_active_temp_path(leased_path))
            format_paper_module.release_active_temp_path(leased_path)

            self.assertFalse(format_paper_module.is_active_temp_path(leased_path))
            self.assertEqual(format_paper_module.get_active_temp_paths(), leases_before)

    def test_format_log_exception_sanitizes_paths_without_spreading_characters(self):
        secret_path = "/tmp/private/job-input.docx"
        message = format_paper_module.format_log_exception(
            RuntimeError(f"failed at {secret_path}\nplease retry"),
            secret_path,
        )

        self.assertEqual(message, "RuntimeError: failed at job-input.docx please retry")
        self.assertNotIn(secret_path, message)
        self.assertNotIn("f a i l e d", message)

        inferred_path_message = format_paper_module.format_log_exception(
            RuntimeError("failed at /tmp/private/other-input.docx\r\nplease retry")
        )
        self.assertEqual(inferred_path_message, "RuntimeError: failed at other-input.docx please retry")
        self.assertNotIn("/tmp/private", inferred_path_message)

    def test_emit_progress_logs_sanitized_callback_errors(self):
        def failing_callback(_payload):
            raise RuntimeError("callback failed at /tmp/private/progress-input.docx\r\ntry again")

        with self.assertLogs(format_paper_module.logger, level="WARNING") as logs:
            format_paper_module.emit_progress(failing_callback, 1, "正在处理")

        log_text = "\n".join(logs.output)
        self.assertIn("RuntimeError: callback failed at progress-input.docx try again", log_text)
        self.assertNotIn("/tmp/private", log_text)

    def test_format_docx_footnotes_logs_sanitized_rewrite_errors(self):
        def failing_rewrite(path, *_args, **_kwargs):
            raise RuntimeError(f"rewrite failed for {path}\r\ntry again")

        with tempfile.TemporaryDirectory() as temp_dir:
            private_dir = Path(temp_dir) / "private"
            private_dir.mkdir()
            secret_path = private_dir / "footnotes-input.docx"
            Document().save(secret_path)

            with patch.object(
                format_paper_module,
                "_rewrite_docx_part",
                side_effect=failing_rewrite,
            ):
                with self.assertLogs(
                    format_paper_module.logger,
                    level="WARNING",
                ) as logs:
                    result = format_paper_module.format_docx_footnotes(
                        secret_path
                    )

        log_text = "\n".join(logs.output)
        self.assertEqual(result, 0)
        self.assertIn("RuntimeError: rewrite failed for footnotes-input.docx try again", log_text)
        self.assertNotIn(str(private_dir), log_text)

    def test_format_docx_footnotes_uses_input_limits_and_one_validated_stream(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            docx_path = Path(temp_dir) / "direct-footnotes.docx"
            Document().save(docx_path)
            real_validator = format_paper_module.is_valid_docx_stream

            with (
                patch.object(
                    format_paper_module,
                    "is_valid_docx_stream",
                    side_effect=real_validator,
                ) as validator,
                patch.object(
                    format_paper_module,
                    "_rewrite_docx_part",
                    return_value=0,
                ) as rewrite,
            ):
                self.assertEqual(
                    format_paper_module.format_docx_footnotes(docx_path),
                    0,
                )

            validated_stream = validator.call_args.args[0]
            self.assertIs(
                validator.call_args.kwargs["limits"],
                format_paper_module.INPUT_DOCX_LIMITS,
            )
            self.assertIs(
                rewrite.call_args.kwargs["source_stream"],
                validated_stream,
            )
            self.assertTrue(validated_stream.closed)

    def test_format_docx_footnotes_rejects_over_budget_package_before_rewrite(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            docx_path = Path(temp_dir) / "oversized-footnotes.docx"
            document = Document()
            document.add_paragraph("footnote anchor")
            document.save(docx_path)
            self.inject_simple_footnote(
                docx_path,
                "x" * (512 * 1024),
            )
            original_bytes = docx_path.read_bytes()
            with ZipFile(docx_path, "r") as archive:
                footnote_size = archive.getinfo("word/footnotes.xml").file_size
                largest_other_xml = max(
                    member.file_size
                    for member in archive.infolist()
                    if member.filename != "word/footnotes.xml"
                    and member.filename.lower().endswith((".xml", ".rels"))
                )
            self.assertGreater(footnote_size, largest_other_xml)
            limits = replace(
                format_paper_module.INPUT_DOCX_LIMITS,
                max_xml_member_uncompressed_bytes=largest_other_xml,
            )

            with patch.object(
                format_paper_module,
                "_rewrite_docx_part",
            ) as rewrite:
                with self.assertLogs(
                    format_paper_module.logger,
                    level="WARNING",
                ):
                    with self.assertRaises(
                        format_paper_module.InvalidInputDocxError
                    ):
                        format_paper_module.format_docx_footnotes(
                            docx_path,
                            limits=limits,
                            raise_on_error=True,
                        )

            rewrite.assert_not_called()
            self.assertEqual(docx_path.read_bytes(), original_bytes)

    def test_footnote_amplification_does_not_replace_validated_input(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            docx_path = Path(temp_dir) / "amplified-footnotes.docx"
            document = Document()
            document.add_paragraph("footnote anchor")
            document.save(docx_path)
            self.inject_simple_footnote(
                docx_path,
                "unused",
                footnote_content_xml="<w:r/>" * 5_000,
            )
            original_bytes = docx_path.read_bytes()
            with ZipFile(docx_path, "r") as archive:
                original_footnote_size = archive.getinfo(
                    "word/footnotes.xml"
                ).file_size
                largest_other_xml = max(
                    member.file_size
                    for member in archive.infolist()
                    if member.filename != "word/footnotes.xml"
                    and member.filename.lower().endswith((".xml", ".rels"))
                )
            self.assertLess(original_footnote_size, largest_other_xml)
            generated_limits = replace(
                format_paper_module.GENERATED_DOCX_LIMITS,
                max_xml_member_uncompressed_bytes=largest_other_xml,
            )

            with patch.object(
                format_paper_module,
                "GENERATED_DOCX_LIMITS",
                generated_limits,
            ):
                with self.assertLogs(
                    format_paper_module.logger,
                    level="WARNING",
                ):
                    with self.assertRaises(
                        format_paper_module.InvalidGeneratedDocxError
                    ):
                        format_paper_module.format_docx_footnotes(
                            docx_path,
                            raise_on_error=True,
                        )

            self.assertEqual(docx_path.read_bytes(), original_bytes)
            self.assertEqual(list(Path(temp_dir).glob(".docx-output-*")), [])

    def test_format_docx_footnotes_rejects_dtd_without_modifying_package(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            docx_path = Path(temp_dir) / "footnote-dtd.docx"
            document = Document()
            document.add_paragraph("footnote anchor")
            document.save(str(docx_path))
            self.inject_simple_footnote(docx_path, "placeholder")

            with ZipFile(docx_path, "r") as source:
                entries = source.infolist()
                payloads = {
                    entry.filename: source.read(entry.filename)
                    for entry in entries
                }
            payloads["word/footnotes.xml"] = (
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<!DOCTYPE w:footnotes [<!ENTITY injected "EXPANDED">]>'
                f'<w:footnotes {nsdecls("w")}>'
                '<w:footnote w:id="2"><w:p><w:r><w:t>&injected;</w:t></w:r>'
                "</w:p></w:footnote></w:footnotes>"
            ).encode()
            with ZipFile(docx_path, "w") as target:
                for entry in entries:
                    target.writestr(entry, payloads[entry.filename])
            original_bytes = docx_path.read_bytes()

            with self.assertLogs(format_paper_module.logger, level="WARNING"):
                self.assertEqual(
                    format_paper_module.format_docx_footnotes(docx_path),
                    0,
                )
            self.assertEqual(docx_path.read_bytes(), original_bytes)

            with self.assertRaisesRegex(
                format_paper_module.InvalidInputDocxError,
                "safety validation",
            ):
                format_paper_module.format_docx_footnotes(
                    docx_path,
                    raise_on_error=True,
                )
            self.assertEqual(docx_path.read_bytes(), original_bytes)

    def test_generated_footnote_postprocessing_uses_generated_limits(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "generated.docx"
            document = Document()
            document.add_paragraph("generated content")

            with patch.object(
                format_paper_module,
                "format_docx_footnotes",
                return_value=0,
            ) as formatter:
                self.assertEqual(
                    format_paper_module._save_docx_with_footnote_postprocessing(
                        document.save,
                        output_path,
                    ),
                    0,
                )

            self.assertIs(
                formatter.call_args.kwargs["limits"],
                format_paper_module.GENERATED_DOCX_LIMITS,
            )
            self.assertTrue(output_path.exists())

    def test_generated_footnote_postprocessing_rejects_invalid_final_package(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "generated.docx"

            def save_invalid(destination):
                Path(destination).write_bytes(b"not-a-docx-package")

            with patch.object(
                format_paper_module,
                "format_docx_footnotes",
                return_value=0,
            ):
                with self.assertRaises(
                    format_paper_module.InvalidGeneratedDocxError
                ):
                    format_paper_module._save_docx_with_footnote_postprocessing(
                        save_invalid,
                        output_path,
                    )

            self.assertFalse(output_path.exists())

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

    def test_format_from_text_enforces_output_limit_during_zip_write(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "limited.docx"

            with self.assertRaises(OutputSizeLimitExceeded) as caught:
                format_academic_paper_from_text(
                    "论文标题\n摘要：测试\n关键词：测试\n正文内容",
                    str(output_path),
                    max_output_bytes=1024,
            )

            self.assertEqual(caught.exception.max_bytes, 1024)
            self.assertFalse(output_path.exists())

    def test_format_from_text_accepts_document_within_output_limit(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "bounded.docx"

            result = format_academic_paper_from_text(
                "论文标题\n摘要：测试\n关键词：测试\n正文内容",
                str(output_path),
                max_output_bytes=64 * 1024,
            )

            self.assertTrue(result)
            self.assertLessEqual(output_path.stat().st_size, 64 * 1024)
            self.assertGreater(len(Document(str(output_path)).paragraphs), 0)

    def test_format_pipeline_preserves_existing_target_when_footnote_stage_fails(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_path = temp_path / "existing.docx"
            sentinel = b"ORIGINAL-SENTINEL"
            output_path.write_bytes(sentinel)

            with patch.object(
                format_paper_module,
                "format_docx_footnotes",
                side_effect=OutputSizeLimitExceeded(64 * 1024),
            ):
                with self.assertRaises(OutputSizeLimitExceeded):
                    format_academic_paper_from_text(
                        "论文标题\n摘要：测试\n关键词：测试\n正文内容",
                        str(output_path),
                        max_output_bytes=64 * 1024,
                    )

            self.assertEqual(output_path.read_bytes(), sentinel)
            self.assertEqual(list(temp_path.glob(".docx-output-*")), [])

    def test_format_pipeline_does_not_publish_when_footnote_rewrite_errors(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            output_path = temp_path / "existing.docx"
            sentinel = b"ORIGINAL-SENTINEL"
            output_path.write_bytes(sentinel)

            with patch.object(
                format_paper_module,
                "_rewrite_docx_part",
                side_effect=RuntimeError("rewrite failed"),
            ):
                result = format_academic_paper_from_text(
                    "论文标题\n摘要：测试\n关键词：测试\n正文内容",
                    str(output_path),
                    max_output_bytes=64 * 1024,
                )

            self.assertFalse(result)
            self.assertEqual(output_path.read_bytes(), sentinel)
            self.assertEqual(list(temp_path.glob(".docx-output-*")), [])

    def test_footnote_rewrite_enforces_output_limit_in_place(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "footnotes.docx"
            doc = Document()
            doc.add_paragraph("含脚注的正文")
            doc.save(str(output_path))
            self.inject_simple_footnote(output_path, "脚注内容")
            original_bytes = output_path.read_bytes()

            with self.assertRaises(OutputSizeLimitExceeded):
                format_paper_module.format_docx_footnotes(
                    output_path,
                    max_output_bytes=1024,
                )

            self.assertEqual(output_path.read_bytes(), original_bytes)

    def test_footnote_rewrite_enforces_output_limit_when_part_is_missing(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "without-footnotes.docx"
            document = Document()
            document.add_paragraph("不含脚注的正文")
            document.save(str(output_path))
            original_bytes = output_path.read_bytes()

            with self.assertRaises(OutputSizeLimitExceeded):
                format_paper_module.format_docx_footnotes(
                    output_path,
                    max_output_bytes=len(original_bytes) - 1,
                )

            self.assertEqual(output_path.read_bytes(), original_bytes)

    def test_rewrite_docx_part_enforces_output_limit_before_unchanged_return(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "unchanged.docx"
            document = Document()
            document.add_paragraph("无需改写的正文")
            document.save(str(output_path))
            original_bytes = output_path.read_bytes()

            with self.assertRaises(OutputSizeLimitExceeded):
                format_paper_module._rewrite_docx_part(
                    output_path,
                    "word/document.xml",
                    lambda xml_bytes: (xml_bytes, 0, False),
                    max_output_bytes=len(original_bytes) - 1,
                )

            self.assertEqual(output_path.read_bytes(), original_bytes)

    def test_rewrite_docx_part_streams_unchanged_archive_members(self):
        payload = b"large-binary-payload" * (128 * 1024)

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            leases_before = format_paper_module.get_active_temp_paths()
            for atomic_write in (True, False):
                with self.subTest(atomic_write=atomic_write):
                    docx_path = temp_path / f"streamed-{atomic_write}.docx"
                    with ZipFile(
                        docx_path,
                        "w",
                        compression=ZIP_DEFLATED,
                    ) as archive:
                        archive.comment = b"preserved archive comment"
                        archive.writestr("word/footnotes.xml", b"<old/>")
                        archive.writestr("word/media/large.bin", payload)

                    original_read = format_paper_module.ZipFile.read
                    read_members = []

                    def tracked_read(archive, member, *args, **kwargs):
                        member_name = getattr(member, "filename", member)
                        read_members.append(member_name)
                        if member_name == "word/media/large.bin":
                            raise AssertionError(
                                "unchanged ZIP members must not be materialized"
                            )
                        return original_read(archive, member, *args, **kwargs)

                    with patch.object(
                        format_paper_module.ZipFile,
                        "read",
                        new=tracked_read,
                    ):
                        count = format_paper_module._rewrite_docx_part(
                            docx_path,
                            "word/footnotes.xml",
                            lambda _xml: (b"<new/>", 1, True),
                            atomic_write=atomic_write,
                        )

                    self.assertEqual(count, 1)
                    self.assertEqual(read_members, ["word/footnotes.xml"])
                    with ZipFile(docx_path, "r") as archive:
                        self.assertEqual(archive.comment, b"preserved archive comment")
                        self.assertEqual(archive.read("word/footnotes.xml"), b"<new/>")
                        self.assertEqual(archive.read("word/media/large.bin"), payload)

            self.assertEqual(
                format_paper_module.get_active_temp_paths(),
                leases_before,
            )

    def test_rewrite_docx_part_rechecks_staging_after_output_validation(self):
        """输出校验器替换重写路径后不得触发外部文件替换。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            docx_path = temp_path / "source.docx"
            external_path = temp_path / "external.docx"
            Document().save(docx_path)
            external_path.write_bytes(b"EXTERNAL")
            original_bytes = docx_path.read_bytes()

            def validator(stream, *, limits):
                staging_path = Path(stream.name)
                stream.close()
                staging_path.unlink()
                staging_path.symlink_to(external_path)
                return True

            with patch.object(
                format_paper_module,
                "is_valid_docx_stream",
                side_effect=validator,
            ):
                with self.assertRaises(format_paper_module.InvalidGeneratedDocxError):
                    format_paper_module._rewrite_docx_part(
                        docx_path,
                        "word/document.xml",
                        lambda _xml: (b"<changed/>", 1, True),
                        output_validation_limits=object(),
                    )

            self.assertEqual(docx_path.read_bytes(), original_bytes)
            self.assertEqual(external_path.read_bytes(), b"EXTERNAL")

    def test_format_academic_paper_logs_safe_filenames_without_temp_paths(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            tmp = Path(temp_dir)
            input_path = tmp / "input.docx"
            output_path = tmp / "output.docx"

            doc = Document()
            doc.add_paragraph("基于多元回归模型的城市化研究")
            doc.add_paragraph("摘要：这是摘要内容")
            doc.add_paragraph("关键词：城市化 回归")
            doc.add_paragraph("正文内容")
            doc.save(str(input_path))

            with self.assertLogs(format_paper_module.logger, level="INFO") as logs:
                result = format_academic_paper(str(input_path), str(output_path))

            self.assertTrue(result)
            log_text = "\n".join(logs.output)
            self.assertIn("成功读取文档：input.docx", log_text)
            self.assertIn("排版完成！已保存至：output.docx", log_text)
            self.assertNotIn(temp_dir, log_text)

    def test_format_academic_paper_does_not_duplicate_hyperlinked_heading_text(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            tmp = Path(temp_dir)
            input_path = tmp / "hyperlink_heading_input.docx"
            output_path = tmp / "hyperlink_heading_output.docx"

            doc = Document()
            doc.add_paragraph("基于多元回归模型的城市化研究")
            doc.add_paragraph("摘要：这是摘要内容")
            doc.add_paragraph("关键词：城市化 回归")
            heading = doc.add_paragraph()
            heading.add_run("1 ")
            relationship_id = doc.part.relate_to(
                "https://example.com/heading",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                is_external=True,
            )
            hyperlink = OxmlElement("w:hyperlink")
            hyperlink.set(qn("r:id"), relationship_id)
            hyperlink_run = OxmlElement("w:r")
            hyperlink_text = OxmlElement("w:t")
            hyperlink_text.text = "引言"
            hyperlink_run.append(hyperlink_text)
            hyperlink.append(hyperlink_run)
            heading._element.append(hyperlink)
            doc.add_paragraph("正文内容")
            doc.save(str(input_path))

            self.assertTrue(format_academic_paper(str(input_path), str(output_path)))

            output_doc = Document(str(output_path))
            output_heading = output_doc.paragraphs[3]
            self.assertEqual(output_heading.text, "引言")
            self.assertEqual(
                [node.text for node in output_heading._element.iter(qn("w:t"))],
                ["引言"],
            )
            self.assertIsNone(output_heading._element.find(qn("w:hyperlink")))

    def test_format_academic_paper_applies_fonts_inside_body_hyperlinks(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "hyperlink_body_input.docx"
            output_path = temp_path / "hyperlink_body_output.docx"

            doc = Document()
            doc.add_paragraph("论文标题")
            doc.add_paragraph("摘要：这是摘要内容")
            doc.add_paragraph("关键词：超链接 测试")
            body = doc.add_paragraph("正文 ")
            relationship_id = doc.part.relate_to(
                "https://example.com/body",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                is_external=True,
            )
            hyperlink = OxmlElement("w:hyperlink")
            hyperlink.set(qn("r:id"), relationship_id)
            hyperlink_run = OxmlElement("w:r")
            hyperlink_text = OxmlElement("w:t")
            hyperlink_text.text = "正文链接"
            hyperlink_run.append(hyperlink_text)
            hyperlink.append(hyperlink_run)
            body._element.append(hyperlink)
            doc.save(str(input_path))

            self.assertTrue(format_academic_paper(str(input_path), str(output_path)))

            output_body = Document(str(output_path)).paragraphs[3]
            run_elements = list(output_body._element.iter(qn("w:r")))
            self.assertEqual(
                [node.text for node in output_body._element.iter(qn("w:t"))],
                ["正文 ", "正文链接"],
            )
            self.assertEqual(len(run_elements), 2)
            for run_element in run_elements:
                run_properties = run_element.find(qn("w:rPr"))
                self.assertIsNotNone(run_properties)
                fonts = run_properties.find(qn("w:rFonts"))
                self.assertEqual(fonts.get(qn("w:eastAsia")), "宋体")
                self.assertEqual(fonts.get(qn("w:ascii")), "Times New Roman")
                self.assertEqual(run_properties.find(qn("w:sz")).get(qn("w:val")), "24")

            self.assertIsNotNone(output_body._element.find(qn("w:hyperlink")))

    def test_format_academic_paper_preserves_rewrite_sensitive_inline_semantics(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "inline-semantics-input.docx"
            output_path = temp_path / "inline-semantics-output.docx"

            doc = Document()

            title = doc.add_paragraph()
            bookmark_start = OxmlElement("w:bookmarkStart")
            bookmark_start.set(qn("w:id"), "7")
            bookmark_start.set(qn("w:name"), "PaperTitle")
            title._element.append(bookmark_start)
            title.add_run("行内语义保全测试")
            bookmark_end = OxmlElement("w:bookmarkEnd")
            bookmark_end.set(qn("w:id"), "7")
            title._element.append(bookmark_end)

            abstract = doc.add_paragraph("摘要：包含脚注引用")
            footnote_run = abstract.add_run()
            footnote_reference = OxmlElement("w:footnoteReference")
            footnote_reference.set(qn("w:id"), "2")
            footnote_run._r.append(footnote_reference)

            keywords = doc.add_paragraph("关键词：书签 域 ")
            content_control = OxmlElement("w:sdt")
            content_control_properties = OxmlElement("w:sdtPr")
            content_control_body = OxmlElement("w:sdtContent")
            controlled_run = OxmlElement("w:r")
            controlled_text = OxmlElement("w:t")
            controlled_text.text = "内容控件"
            controlled_run.append(controlled_text)
            content_control_body.append(controlled_run)
            content_control.append(content_control_properties)
            content_control.append(content_control_body)
            keywords._element.append(content_control)

            caption = doc.add_paragraph("图")
            field_begin_run = caption.add_run()
            field_begin = OxmlElement("w:fldChar")
            field_begin.set(qn("w:fldCharType"), "begin")
            field_begin_run._r.append(field_begin)
            instruction_run = caption.add_run()
            instruction = OxmlElement("w:instrText")
            instruction.text = " SEQ Figure "
            instruction_run._r.append(instruction)
            field_separator_run = caption.add_run()
            field_separator = OxmlElement("w:fldChar")
            field_separator.set(qn("w:fldCharType"), "separate")
            field_separator_run._r.append(field_separator)
            caption.add_run("9")
            field_end_run = caption.add_run()
            field_end = OxmlElement("w:fldChar")
            field_end.set(qn("w:fldCharType"), "end")
            field_end_run._r.append(field_end)
            caption.add_run(" 带域图题")

            doc.add_paragraph("正文内容。")
            doc.save(str(input_path))
            self.inject_simple_footnote(input_path, "对应的脚注内容")

            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertIsInstance(summary, dict)
            self.assertEqual(summary["stats"]["title"], 1)
            self.assertEqual(summary["stats"]["abstract"], 1)
            self.assertEqual(summary["stats"]["keywords"], 1)
            self.assertEqual(summary["stats"]["figure_caption"], 1)
            self.assertEqual(summary["formatted_footnotes"], 1)

            output_doc = Document(str(output_path))
            output_title = next(
                paragraph for paragraph in output_doc.paragraphs
                if paragraph.text == "行内语义保全测试"
            )
            output_abstract = next(
                paragraph for paragraph in output_doc.paragraphs
                if paragraph.text == "摘要：包含脚注引用"
            )
            output_keywords = next(
                paragraph for paragraph in output_doc.paragraphs
                if paragraph.text.startswith("关键词：书签 域")
            )
            output_caption = next(
                paragraph for paragraph in output_doc.paragraphs
                if paragraph.text == "图9 带域图题"
            )

            self.assertEqual(output_title._element.find(qn("w:bookmarkStart")).get(qn("w:name")), "PaperTitle")
            self.assertEqual(output_title._element.find(qn("w:bookmarkEnd")).get(qn("w:id")), "7")
            self.assertEqual(len(output_abstract._element.findall(".//" + qn("w:footnoteReference"))), 1)
            self.assertEqual(
                output_keywords._element.find(".//" + qn("w:sdtContent") + "/" + qn("w:r") + "/" + qn("w:t")).text,
                "内容控件",
            )
            self.assertEqual(
                [node.get(qn("w:fldCharType")) for node in output_caption._element.findall(".//" + qn("w:fldChar"))],
                ["begin", "separate", "end"],
            )
            self.assertEqual(output_caption._element.find(".//" + qn("w:instrText")).text.strip(), "SEQ Figure")

            with ZipFile(output_path, "r") as archive:
                self.assertIsNone(archive.testzip())
                self.assertIn("word/footnotes.xml", archive.namelist())

    def test_format_academic_paper_preserves_semantic_run_properties_in_headings(self):
        """Rebuilding numbered headings must not flatten superscript/subscript runs."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "heading-semantic-input.docx"
            output_path = temp_path / "heading-semantic-output.docx"

            doc = Document()
            doc.add_paragraph("论文标题")
            doc.add_paragraph("摘要：测试内容")
            heading = doc.add_paragraph()
            heading.add_run("1 CO")
            superscript_run = heading.add_run("2")
            vert_align = OxmlElement("w:vertAlign")
            vert_align.set(qn("w:val"), "superscript")
            superscript_run._r.get_or_add_rPr().append(vert_align)
            doc.add_paragraph("正文内容")
            doc.save(input_path)

            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertEqual(summary["stats"]["heading_l1"], 1)
            output_doc = Document(output_path)
            output_heading = next(
                paragraph for paragraph in output_doc.paragraphs if "CO2" in paragraph.text
            )
            self.assertEqual(output_heading.text, "1 CO2")
            output_superscript = output_heading._element.find(
                ".//" + qn("w:vertAlign")
            )
            self.assertIsNotNone(output_superscript)
            self.assertEqual(output_superscript.get(qn("w:val")), "superscript")

    def test_format_academic_paper_classifies_content_control_text(self):
        """标题/摘要完全位于内容控件时仍应参与结构识别并保留控件。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "sdt-structure-input.docx"
            output_path = Path(temp_dir) / "sdt-structure-output.docx"

            doc = Document()

            def add_content_control_paragraph(text):
                paragraph = doc.add_paragraph()
                sdt = OxmlElement("w:sdt")
                sdt.append(OxmlElement("w:sdtPr"))
                content = OxmlElement("w:sdtContent")
                run = OxmlElement("w:r")
                text_node = OxmlElement("w:t")
                text_node.text = text
                run.append(text_node)
                content.append(run)
                sdt.append(content)
                paragraph._element.append(sdt)
                return paragraph

            title = add_content_control_paragraph("内容控件论文标题")
            abstract = add_content_control_paragraph("摘要：内容控件摘要")
            doc.add_paragraph("关键词：控件 识别")
            doc.add_paragraph("正文内容。")
            doc.save(input_path)

            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertIsInstance(summary, dict)
            self.assertEqual(summary["stats"]["title"], 1)
            self.assertEqual(summary["stats"]["abstract"], 1)
            output_doc = Document(output_path)
            self.assertEqual(output_doc.paragraphs[0].text, "")
            self.assertEqual(output_doc.paragraphs[1].text, "")
            self.assertIsNotNone(output_doc.paragraphs[0]._element.find(qn("w:sdt")))
            self.assertIsNotNone(output_doc.paragraphs[1]._element.find(qn("w:sdt")))
            self.assertEqual(
                "".join(
                    node.text or ""
                    for node in output_doc.paragraphs[0]._element.iter(qn("w:t"))
                ),
                "内容控件论文标题",
            )
            self.assertEqual(
                output_doc.paragraphs[0].paragraph_format.alignment,
                WD_ALIGN_PARAGRAPH.CENTER,
            )

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

    def test_apply_document_layout_reuses_header_footer_parts_across_sections(self):
        doc = Document()
        doc.add_paragraph("第一节")
        for index in range(1, 41):
            doc.add_section(WD_SECTION.NEW_PAGE)
            doc.add_paragraph(f"第 {index + 1} 节")

        apply_document_layout(doc, "多分节共享页眉")

        self.assertFalse(doc.sections[0].header.is_linked_to_previous)
        self.assertFalse(doc.sections[0].footer.is_linked_to_previous)
        for section in doc.sections[1:]:
            self.assertTrue(section.header.is_linked_to_previous)
            self.assertTrue(section.footer.is_linked_to_previous)

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)
        try:
            doc.save(str(output_path))
            with ZipFile(output_path, "r") as archive:
                member_names = archive.namelist()
            self.assertEqual(len([name for name in member_names if re.fullmatch(r"word/header\d+\.xml", name)]), 1)
            self.assertEqual(len([name for name in member_names if re.fullmatch(r"word/footer\d+\.xml", name)]), 1)
        finally:
            output_path.unlink(missing_ok=True)

    def test_apply_document_layout_removes_all_legacy_header_footer_blocks(self):
        tiny_png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )
        doc = Document()
        doc.add_paragraph("页眉页脚旧结构清理测试")
        section = doc.sections[0]

        def add_legacy_content(container):
            paragraph = container.paragraphs[0]
            paragraph.add_run("旧内容")
            paragraph._element.append(
                parse_xml(
                    r"""
                    <m:oMathPara %s>
                      <m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath>
                    </m:oMathPara>
                    """ % nsdecls("m")
                )
            )
            field = OxmlElement("w:fldSimple")
            field.set(qn("w:instr"), "DATE")
            field_run = OxmlElement("w:r")
            field_text = OxmlElement("w:t")
            field_text.text = "旧日期"
            field_run.append(field_text)
            field.append(field_run)
            paragraph._element.append(field)
            content_control = OxmlElement("w:sdt")
            content_control.append(OxmlElement("w:sdtPr"))
            content_control.append(OxmlElement("w:sdtContent"))
            paragraph._element.append(content_control)
            paragraph.add_run().add_picture(BytesIO(tiny_png), width=Inches(0.2))
            relationship_id = container.part.relate_to(
                "https://example.test/legacy-header-footer",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                is_external=True,
            )
            hyperlink = OxmlElement("w:hyperlink")
            hyperlink.set(qn("r:id"), relationship_id)
            hyperlink.append(OxmlElement("w:r"))
            paragraph._element.append(hyperlink)
            container.add_table(rows=1, cols=1, width=Inches(1)).cell(0, 0).text = "旧表格"
            container.add_paragraph("旧的额外段落")

        add_legacy_content(section.header)
        add_legacy_content(section.footer)

        apply_document_layout(doc, "标准运行页眉")

        self.assertEqual(len(section.header.paragraphs), 1)
        self.assertEqual(len(section.header.tables), 0)
        self.assertEqual(section.header.paragraphs[0].text, "标准运行页眉")
        self.assertEqual(len(section.footer.paragraphs), 1)
        self.assertEqual(len(section.footer.tables), 0)
        self.assertIn("PAGE", section.footer.paragraphs[0]._element.xml)
        for container in (section.header, section.footer):
            self.assertIsNone(container._element.find(".//" + qn("m:oMath")))
            self.assertIsNone(container._element.find(".//" + qn("w:fldSimple")))
            self.assertIsNone(container._element.find(".//" + qn("w:sdt")))
            self.assertEqual(len(container.part.rels), 0)

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)
        try:
            doc.save(str(output_path))
            reopened = Document(str(output_path))
            reopened_section = reopened.sections[0]
            self.assertEqual(reopened_section.header.paragraphs[0].text, "标准运行页眉")
            self.assertEqual(len(reopened_section.header.tables), 0)
            self.assertEqual(len(reopened_section.footer.tables), 0)
            self.assertNotIn("旧内容", reopened_section.header._element.xml)
            self.assertNotIn("旧表格", reopened_section.footer._element.xml)
            with ZipFile(output_path, "r") as archive:
                member_names = archive.namelist()
            self.assertFalse(
                any(
                    re.fullmatch(r"word/_rels/(?:header|footer)\d+\.xml\.rels", name)
                    for name in member_names
                )
            )
            self.assertFalse(any(name.startswith("word/media/") for name in member_names))
        finally:
            output_path.unlink(missing_ok=True)

    def test_apply_document_layout_keeps_media_still_referenced_by_body(self):
        tiny_png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )
        doc = Document()
        body_paragraph = doc.add_paragraph("正文共享图片")
        body_paragraph.add_run().add_picture(BytesIO(tiny_png), width=Inches(0.2))
        section = doc.sections[0]
        section.header.paragraphs[0].add_run().add_picture(
            BytesIO(tiny_png),
            width=Inches(0.2),
        )

        apply_document_layout(doc, "标准运行页眉")

        self.assertEqual(len(section.header.part.rels), 0)
        self.assertEqual(len(doc.inline_shapes), 1)

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)
        try:
            doc.save(str(output_path))
            reopened = Document(str(output_path))
            self.assertEqual(len(reopened.inline_shapes), 1)
            self.assertEqual(
                reopened.sections[0].header.paragraphs[0].text,
                "标准运行页眉",
            )
            with ZipFile(output_path, "r") as archive:
                member_names = archive.namelist()
            self.assertEqual(
                len([name for name in member_names if name.startswith("word/media/")]),
                1,
            )
            self.assertFalse(
                any(re.fullmatch(r"word/_rels/header\d+\.xml\.rels", name) for name in member_names)
            )
        finally:
            output_path.unlink(missing_ok=True)

    def test_apply_document_layout_removes_stale_even_page_headers_and_footers(self):
        doc = Document()
        doc.add_paragraph("奇偶页页眉清理测试")
        doc.settings.odd_and_even_pages_header_footer = True
        section = doc.sections[0]
        section.even_page_header.paragraphs[0].text = "旧偶数页页眉"
        section.even_page_footer.paragraphs[0].text = "旧偶数页页脚"

        apply_document_layout(doc, "新统一页眉")

        self.assertFalse(doc.settings.odd_and_even_pages_header_footer)
        self.assertFalse(
            any(
                reference.get(qn("w:type")) == "even"
                for tag in ("w:headerReference", "w:footerReference")
                for reference in section._sectPr.findall(qn(tag))
            )
        )

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)
        try:
            doc.save(output_path)
            reopened = Document(output_path)
            self.assertFalse(
                reopened.settings.odd_and_even_pages_header_footer
            )
            self.assertEqual(
                reopened.sections[0].header.paragraphs[0].text,
                "新统一页眉",
            )
            with ZipFile(output_path) as archive:
                xml_payload = b"\n".join(
                    archive.read(name)
                    for name in archive.namelist()
                    if re.fullmatch(r"word/(?:header|footer)\d+\.xml", name)
                )
            self.assertNotIn("旧偶数页".encode(), xml_payload)
        finally:
            output_path.unlink(missing_ok=True)

    def test_format_academic_paper_preserves_shared_header_footer_relationships(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "shared_header_footer_input.docx"
            output_path = temp_path / "shared_header_footer_output.docx"

            doc = Document()
            doc.add_paragraph("共享页眉页脚关系修复测试")
            doc.add_paragraph("摘要：这是摘要内容。")
            doc.add_paragraph("关键词：页眉 页脚")
            doc.add_paragraph("第一节正文。")
            first_section = doc.sections[0]
            first_section.header.paragraphs[0].text = "原共享页眉"
            first_section.footer.paragraphs[0].text = "原共享页脚"
            doc.add_section(WD_SECTION.NEW_PAGE)
            doc.add_paragraph("第二节正文。")

            first_section = doc.sections[0]
            second_section = doc.sections[1]
            for index, tag in enumerate(("w:headerReference", "w:footerReference")):
                first_reference = first_section._sectPr.find(qn(tag))
                self.assertIsNotNone(first_reference)
                shared_reference = OxmlElement(tag)
                shared_reference.set(qn("w:type"), "default")
                shared_reference.set(qn("r:id"), first_reference.get(qn("r:id")))
                second_section._sectPr.insert(index, shared_reference)
            doc.save(str(input_path))

            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertIsInstance(summary, dict)
            with ZipFile(output_path, "r") as archive:
                document_root = etree.fromstring(archive.read("word/document.xml"))
                relationships_root = etree.fromstring(archive.read("word/_rels/document.xml.rels"))

            relationship_ids = {
                relationship.get("Id")
                for relationship in relationships_root.findall(
                    f"{{{self.PACKAGE_REL_NS}}}Relationship"
                )
            }
            retained_reference_ids = {
                reference.get(qn("r:id"))
                for tag in ("w:headerReference", "w:footerReference")
                for reference in document_root.findall(".//" + qn(tag))
            }
            self.assertTrue(retained_reference_ids)
            self.assertTrue(retained_reference_ids.issubset(relationship_ids))

            output_doc = Document(str(output_path))
            self.assertEqual(len(output_doc.sections), 2)
            for section in output_doc.sections:
                self.assertEqual(
                    section.header.paragraphs[0].text,
                    summary["page_setup"]["header_text"],
                )
                self.assertIn("PAGE", section.footer.paragraphs[0]._element.xml)

    def test_format_academic_paper_replaces_stale_first_page_headers_in_every_section(self):
        for with_cover in (False, True):
            with self.subTest(with_cover=with_cover), tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                input_path = temp_path / "first-page-input.docx"
                output_path = temp_path / "first-page-output.docx"
                doc = Document()
                doc.add_paragraph("不同首页排版测试")
                doc.add_paragraph("摘要：测试原模板中的分节首页")
                doc.add_paragraph("第一节正文。")
                doc.add_section(WD_SECTION.NEW_PAGE)
                doc.add_paragraph("第二节正文。")

                for index, section in enumerate(doc.sections):
                    section.different_first_page_header_footer = True
                    section.first_page_header.is_linked_to_previous = False
                    section.first_page_footer.is_linked_to_previous = False
                    section.first_page_header.paragraphs[0].text = f"旧首页页眉 {index}"
                    section.first_page_footer.paragraphs[0].text = f"旧首页页脚 {index}"
                doc.save(input_path)

                result = format_academic_paper(
                    str(input_path),
                    str(output_path),
                    cover_info={"title": "新封面论文"} if with_cover else None,
                )

                self.assertIsInstance(result, dict)
                self.assertEqual(result["cover_generated"], with_cover)
                reopened = Document(output_path)
                self.assertEqual(len(reopened.sections), 2)
                for index, section in enumerate(reopened.sections):
                    is_cover_section = with_cover and index == 0
                    self.assertEqual(
                        section.different_first_page_header_footer,
                        is_cover_section,
                    )
                    self.assertEqual(section.header.paragraphs[0].text, "不同首页排版测试")
                    self.assertIn("PAGE", section.footer._element.xml)
                    first_references = [
                        reference
                        for tag in ("w:headerReference", "w:footerReference")
                        for reference in section._sectPr.findall(qn(tag))
                        if reference.get(qn("w:type")) == "first"
                    ]
                    self.assertEqual(len(first_references), 2 if is_cover_section else 0)
                    if is_cover_section:
                        self.assertEqual(section.first_page_header.paragraphs[0].text, "")
                        self.assertEqual(section.first_page_footer.paragraphs[0].text, "")

                with ZipFile(output_path) as archive:
                    header_footer_xml = b"\n".join(
                        archive.read(name)
                        for name in archive.namelist()
                        if re.fullmatch(r"word/(?:header|footer)\d+\.xml", name)
                    )
                self.assertNotIn("旧首页".encode(), header_footer_xml)

    def test_apply_document_layout_keeps_default_relationships_shared_with_first_page(self):
        doc = Document()
        doc.add_paragraph("共享首页关系测试")
        section = doc.sections[0]
        section.header.paragraphs[0].text = "旧默认页眉"
        section.footer.paragraphs[0].text = "旧默认页脚"
        section.different_first_page_header_footer = True
        default_ids = {}
        for tag in ("w:headerReference", "w:footerReference"):
            default_reference = section._sectPr.find(qn(tag))
            relationship_id = default_reference.get(qn("r:id"))
            default_ids[tag] = relationship_id
            first_reference = OxmlElement(tag)
            first_reference.set(qn("w:type"), "first")
            first_reference.set(qn("r:id"), relationship_id)
            section._sectPr.append(first_reference)

        apply_document_layout(doc, "标准运行页眉")

        self.assertFalse(section.different_first_page_header_footer)
        for tag, relationship_id in default_ids.items():
            references = section._sectPr.findall(qn(tag))
            self.assertEqual(len(references), 1)
            self.assertEqual(references[0].get(qn("w:type")), "default")
            self.assertEqual(references[0].get(qn("r:id")), relationship_id)
            self.assertIn(relationship_id, doc.part.rels)
        self.assertEqual(section.header.paragraphs[0].text, "标准运行页眉")
        self.assertIn("PAGE", section.footer._element.xml)

    def test_format_academic_paper_rebuilds_invalid_header_footer_reference_types(self):
        allowed_types = {"default", "even", "first"}
        header_footer_rel_types = {
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
        }

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            for index, invalid_type in enumerate(("bogus", None)):
                with self.subTest(invalid_type=invalid_type):
                    input_path = temp_path / f"invalid-reference-type-{index}-input.docx"
                    output_path = temp_path / f"invalid-reference-type-{index}-output.docx"

                    doc = Document()
                    doc.add_paragraph("页眉页脚引用类型修复测试")
                    doc.add_paragraph("摘要：这是摘要内容。")
                    doc.add_paragraph("关键词：页眉 页脚")
                    doc.sections[0].header.paragraphs[0].text = "旧页眉"
                    doc.sections[0].footer.paragraphs[0].text = "旧页脚"
                    for tag in ("w:headerReference", "w:footerReference"):
                        reference = doc.sections[0]._sectPr.find(qn(tag))
                        self.assertIsNotNone(reference)
                        if invalid_type is None:
                            del reference.attrib[qn("w:type")]
                        else:
                            reference.set(qn("w:type"), invalid_type)
                    doc.save(str(input_path))

                    summary = format_academic_paper(str(input_path), str(output_path))

                    self.assertIsInstance(summary, dict)
                    with ZipFile(output_path, "r") as archive:
                        document_root = etree.fromstring(archive.read("word/document.xml"))
                        relationships_root = etree.fromstring(
                            archive.read("word/_rels/document.xml.rels")
                        )

                    references = [
                        reference
                        for tag in ("w:headerReference", "w:footerReference")
                        for reference in document_root.findall(".//" + qn(tag))
                    ]
                    self.assertTrue(references)
                    self.assertTrue(
                        all(reference.get(qn("w:type")) in allowed_types for reference in references)
                    )
                    self.assertNotIn(invalid_type, {reference.get(qn("w:type")) for reference in references})

                    referenced_ids = {reference.get(qn("r:id")) for reference in references}
                    relationship_ids = {
                        relationship.get("Id")
                        for relationship in relationships_root.findall(
                            f"{{{self.PACKAGE_REL_NS}}}Relationship"
                        )
                        if relationship.get("Type") in header_footer_rel_types
                    }
                    self.assertEqual(relationship_ids, referenced_ids)

    def test_format_academic_paper_collapses_duplicate_header_footer_reference_types(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "duplicate-header-footer-input.docx"
            output_path = temp_path / "duplicate-header-footer-output.docx"

            doc = Document()
            doc.add_paragraph("页眉页脚重复引用修复测试")
            doc.add_paragraph("摘要：这是摘要内容。")
            doc.add_paragraph("关键词：页眉 页脚")
            section = doc.sections[0]
            section.header.paragraphs[0].text = "旧页眉"
            section.footer.paragraphs[0].text = "旧页脚"
            for tag in ("w:headerReference", "w:footerReference"):
                reference = section._sectPr.find(qn(tag))
                self.assertIsNotNone(reference)
                duplicate = etree.fromstring(etree.tostring(reference))
                reference.addnext(duplicate)
            doc.save(str(input_path))

            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertIsInstance(summary, dict)
            with ZipFile(output_path, "r") as archive:
                document_root = etree.fromstring(archive.read("word/document.xml"))
            for tag in ("w:headerReference", "w:footerReference"):
                references = document_root.findall(".//" + qn(tag))
                self.assertEqual(len(references), 1)
                self.assertEqual(references[0].get(qn("w:type")), "default")

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

    def test_format_academic_paper_places_update_fields_before_compat(self):
        text = (
            "自动目录域更新顺序测试\n"
            "摘要：这是摘要内容。\n"
            "关键词：目录 域更新\n"
            "1 引言\n"
            "1.1 研究背景\n"
            "正文内容。"
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "update_fields_order.docx"
            summary = format_academic_paper_from_text(text, str(output_path))

            self.assertIsInstance(summary, dict)
            with ZipFile(output_path, "r") as archive:
                settings_root = etree.fromstring(archive.read("word/settings.xml"))

            update_fields = settings_root.find(qn("w:updateFields"))
            compat = settings_root.find(qn("w:compat"))
            self.assertIsNotNone(update_fields)
            self.assertIsNotNone(compat)
            self.assertEqual(update_fields.get(qn("w:val")), "true")
            self.assertLess(settings_root.index(update_fields), settings_root.index(compat))

    def test_enable_field_updates_collapses_duplicate_settings(self):
        doc = Document()
        settings = doc.settings.element
        for value in ("false", "0", "true"):
            update_fields = OxmlElement("w:updateFields")
            update_fields.set(qn("w:val"), value)
            settings.append(update_fields)

        format_paper_module._enable_field_updates_on_open(doc)

        update_fields_nodes = settings.findall(qn("w:updateFields"))
        self.assertEqual(len(update_fields_nodes), 1)
        self.assertEqual(update_fields_nodes[0].get(qn("w:val")), "true")
        compat = settings.find(qn("w:compat"))
        self.assertIsNotNone(compat)
        self.assertLess(settings.index(update_fields_nodes[0]), settings.index(compat))

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
            self.assertEqual(abstract_body._element.pPr.find(qn("w:widowControl")).get(qn("w:val")), "true")

            self.assertEqual(keywords.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.LEFT)
            self.assertTrue(keywords.runs[0].font.bold)
            self.assertEqual(keywords.text, "Keywords: supply chain resilience; cross-border e-commerce")
            self.assertEqual(keywords._element.pPr.find(qn("w:widowControl")).get(qn("w:val")), "true")
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

    def test_format_academic_paper_creates_missing_optional_numbering_part(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "missing_numbering_input.docx"
            output_path = temp_path / "missing_numbering_output.docx"

            doc = Document()
            doc.add_paragraph("缺少编号部件的论文标题")
            doc.add_paragraph("摘要：这是摘要内容")
            doc.add_paragraph("关键词：编号 兼容")
            doc.add_paragraph("1 引言")
            doc.add_paragraph("1.1 研究背景")
            doc.add_paragraph("正文内容")
            doc.save(str(input_path))

            with ZipFile(input_path, "r") as source:
                entries = source.infolist()
                payloads = {
                    entry.filename: source.read(entry.filename)
                    for entry in entries
                }

            content_types = etree.fromstring(payloads["[Content_Types].xml"])
            override_tag = f"{{{self.CONTENT_TYPES_NS}}}Override"
            numbering_overrides = [
                node
                for node in content_types.findall(override_tag)
                if node.get("PartName") == "/word/numbering.xml"
            ]
            self.assertEqual(len(numbering_overrides), 1)
            content_types.remove(numbering_overrides[0])
            payloads["[Content_Types].xml"] = etree.tostring(
                content_types,
                encoding="UTF-8",
                xml_declaration=True,
                standalone=True,
            )

            relationships = etree.fromstring(payloads["word/_rels/document.xml.rels"])
            relationship_tag = f"{{{self.PACKAGE_REL_NS}}}Relationship"
            numbering_relationships = [
                node
                for node in relationships.findall(relationship_tag)
                if node.get("Type")
                == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
            ]
            self.assertEqual(len(numbering_relationships), 1)
            relationships.remove(numbering_relationships[0])
            payloads["word/_rels/document.xml.rels"] = etree.tostring(
                relationships,
                encoding="UTF-8",
                xml_declaration=True,
                standalone=True,
            )
            del payloads["word/numbering.xml"]

            with ZipFile(input_path, "w") as target:
                for entry in entries:
                    if entry.filename == "word/numbering.xml":
                        continue
                    target.writestr(entry, payloads[entry.filename])

            resource_relationship_type = (
                "https://example.test/uppercase-numbering-resource"
            )
            resource_payload = b"<resource>NOT A NUMBERING STORY</resource>"
            missing_numbering_doc = Document(input_path)
            custom_xml_part = (
                missing_numbering_doc.part.rels.part_with_reltype(
                    format_paper_module.RT.CUSTOM_XML
                )
            )
            custom_xml_part.relate_to(
                format_paper_module.Part(
                    format_paper_module.PackURI("/word/NUMBERING.xml"),
                    "application/xml",
                    resource_payload,
                    missing_numbering_doc.part.package,
                ),
                resource_relationship_type,
            )
            missing_numbering_doc.save(input_path)

            # numbering Part 未被正文引用时本就是可选项，python-docx 应能打开输入。
            self.assertEqual(len(Document(str(input_path)).paragraphs), 6)

            summary = format_academic_paper(
                str(input_path),
                str(output_path),
                max_output_bytes=64 * 1024,
            )

            self.assertIsInstance(summary, dict)
            self.assertEqual(summary["stats"]["heading_l1"], 1)
            self.assertEqual(summary["stats"]["heading_l2"], 1)

            output_doc = Document(str(output_path))
            heading_l1 = next(
                paragraph for paragraph in output_doc.paragraphs
                if paragraph.text == "引言"
            )
            heading_l2 = next(
                paragraph for paragraph in output_doc.paragraphs
                if paragraph.text == "研究背景"
            )
            heading_l1_num_id = self.assert_paragraph_has_numbering(heading_l1, 0)
            heading_l2_num_id = self.assert_paragraph_has_numbering(heading_l2, 1)
            self.assert_numbering_overrides(output_doc, heading_l1_num_id, {0: 1})
            self.assert_numbering_overrides(
                output_doc,
                heading_l2_num_id,
                {0: 1, 1: 1},
            )

            with ZipFile(output_path, "r") as archive:
                member_names = archive.namelist()
                self.assertIsNone(archive.testzip())
                self.assertEqual(len(member_names), len(set(member_names)))
                self.assertEqual(
                    len(member_names),
                    len({name.casefold() for name in member_names}),
                )
                numbering_member = str(
                    output_doc.part.part_related_by(
                        format_paper_module.RT.NUMBERING
                    ).partname
                ).lstrip("/")
                self.assertIn(numbering_member, member_names)
                self.assertNotEqual(
                    numbering_member.casefold(),
                    "word/numbering.xml",
                )
            copied_resource = next(
                relationship.target_part
                for custom_relationship in output_doc.part.rels.values()
                if custom_relationship.reltype == format_paper_module.RT.CUSTOM_XML
                for relationship in custom_relationship.target_part.rels.values()
                if relationship.reltype == resource_relationship_type
            )
            self.assertEqual(copied_resource.blob, resource_payload)

    def test_extract_heading_numbering_bounds_extreme_decimal_components(self):
        enormous_heading = f"{'9' * 5000} 极端编号标题"
        above_limit_heading = (
            f"{format_paper_module.MAX_HEADING_NUMBER_COMPONENT + 1} 超限编号标题"
        )

        self.assertEqual(
            format_paper_module.extract_heading_numbering(
                enormous_heading,
                format_paper_module.ParagraphType.HEADING_L1,
            ),
            (enormous_heading, ()),
        )
        self.assertEqual(
            format_paper_module.extract_heading_numbering(
                above_limit_heading,
                format_paper_module.ParagraphType.HEADING_L1,
            ),
            (above_limit_heading, ()),
        )
        self.assertEqual(
            format_paper_module.extract_heading_numbering(
                f"{format_paper_module.MAX_HEADING_NUMBER_COMPONENT} 最大编号标题",
                format_paper_module.ParagraphType.HEADING_L1,
            ),
            ("最大编号标题", (format_paper_module.MAX_HEADING_NUMBER_COMPONENT,)),
        )

    def test_ooxml_numeric_helpers_bound_outline_levels_without_parsing_footnote_ids(self):
        enormous_value = "9" * 5000
        p_pr = OxmlElement("w:pPr")
        outline_level = OxmlElement("w:outlineLvl")
        outline_level.set(qn("w:val"), enormous_value)
        p_pr.append(outline_level)

        self.assertIsNone(format_paper_module._read_outline_level(p_pr))
        self.assertFalse(format_paper_module._is_negative_ooxml_decimal(enormous_value))
        self.assertTrue(format_paper_module._is_negative_ooxml_decimal(f"-{enormous_value}"))
        self.assertFalse(format_paper_module._is_negative_ooxml_decimal("-0000"))

    def test_paragraph_style_metadata_is_indexed_once_per_document(self):
        doc = Document()
        paragraphs = [
            doc.add_paragraph(f"标题 {index}", style="Heading 1")
            for index in range(200)
        ]

        with patch.object(
            format_paper_module,
            "_build_paragraph_style_metadata_index",
            wraps=format_paper_module._build_paragraph_style_metadata_index,
        ) as build_index:
            for paragraph in paragraphs:
                self.assertEqual(format_paper_module._get_paragraph_outline_level_hint(paragraph), 0)
                self.assertFalse(format_paper_module._is_list_paragraph(paragraph))

        self.assertEqual(build_index.call_count, 1)

    def test_clear_paragraph_style_uses_cached_normal_style_id(self):
        doc = Document()
        body_paragraph = doc.add_paragraph("正文", style="Heading 1")
        header_paragraph = doc.sections[0].header.paragraphs[0]
        header_paragraph.style = "Heading 1"
        normal_style = format_paper_module._ensure_normal_paragraph_style(doc)

        with patch.object(
            doc.part,
            "get_style_id",
            side_effect=AssertionError("不应逐段落重复查找 Normal 样式"),
        ):
            format_paper_module._clear_paragraph_style(body_paragraph)
            format_paper_module._clear_paragraph_style(header_paragraph)

        for paragraph in (body_paragraph, header_paragraph):
            p_style = paragraph._element.pPr.find(qn("w:pStyle"))
            self.assertIsNotNone(p_style)
            self.assertEqual(p_style.get(qn("w:val")), normal_style.style_id)

    def test_format_academic_paper_infers_numbering_from_heading_styles(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "style_heading_input.docx"
            output_path = temp_path / "style_heading_output.docx"

            doc = Document()
            doc.add_paragraph("平台治理视角下数字化协同研究")
            doc.add_paragraph("摘要：这是摘要内容")
            doc.add_paragraph("关键词：平台治理 协同")
            doc.add_paragraph("引言", style="Heading 1")
            doc.add_paragraph("研究背景", style="Heading 2")
            doc.add_paragraph("研究假设", style="Heading 3")
            doc.add_paragraph("正文内容")
            doc.save(str(input_path))

            summary = format_academic_paper(
                str(input_path),
                str(output_path),
                max_output_bytes=64 * 1024,
            )

            self.assertEqual(summary["stats"]["heading_l1"], 1)
            self.assertEqual(summary["stats"]["heading_l2"], 1)
            self.assertEqual(summary["stats"]["heading_l3"], 1)

            output_doc = Document(str(output_path))
            heading_l1 = next(paragraph for paragraph in output_doc.paragraphs if paragraph.text == "引言")
            heading_l2 = next(paragraph for paragraph in output_doc.paragraphs if paragraph.text == "研究背景")
            heading_l3 = next(paragraph for paragraph in output_doc.paragraphs if paragraph.text == "研究假设")

            heading_l1_num_id = self.assert_paragraph_has_numbering(heading_l1, 0)
            heading_l2_num_id = self.assert_paragraph_has_numbering(heading_l2, 1)
            heading_l3_num_id = self.assert_paragraph_has_numbering(heading_l3, 2)

            self.assert_numbering_overrides(output_doc, heading_l1_num_id, {0: 1})
            self.assert_numbering_overrides(output_doc, heading_l2_num_id, {0: 1, 1: 1})
            self.assert_numbering_overrides(output_doc, heading_l3_num_id, {0: 1, 1: 1, 2: 1})

    def test_format_academic_paper_repairs_missing_or_wrong_type_normal_style(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)

            for variant in ("missing", "character", "bogus"):
                with self.subTest(variant=variant):
                    input_path = temp_path / f"normal_style_{variant}_input.docx"
                    output_path = temp_path / f"normal_style_{variant}_output.docx"
                    doc = Document()
                    doc.add_paragraph("默认样式修复测试")
                    doc.add_paragraph("摘要：这是摘要内容。")
                    doc.add_paragraph("关键词：样式 修复")
                    doc.add_paragraph("正文内容。")
                    doc.save(str(input_path))

                    def mutate_styles(xml_bytes):
                        root = etree.fromstring(xml_bytes)
                        normal = next(
                            style
                            for style in root.findall(qn("w:style"))
                            if style.get(qn("w:styleId")) == "Normal"
                        )
                        if variant == "missing":
                            root.remove(normal)
                        else:
                            normal.set(qn("w:type"), variant)
                        return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)

                    self.rewrite_docx_member(input_path, "word/styles.xml", mutate_styles)
                    summary = format_academic_paper(str(input_path), str(output_path))

                    self.assertIsInstance(summary, dict)
                    with ZipFile(output_path, "r") as archive:
                        styles_root = etree.fromstring(archive.read("word/styles.xml"))
                    repaired_normal = next(
                        style
                        for style in styles_root.findall(qn("w:style"))
                        if style.get(qn("w:styleId")) == "Normal"
                    )
                    self.assertEqual(repaired_normal.get(qn("w:type")), "paragraph")

    def test_format_academic_paper_returns_false_for_wrong_styles_root(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "wrong_styles_root_input.docx"
            output_path = temp_path / "wrong_styles_root_output.docx"
            doc = Document()
            doc.add_paragraph("样式根节点异常测试")
            doc.add_paragraph("摘要：这是摘要内容。")
            doc.add_paragraph("关键词：根节点")
            doc.add_paragraph("正文内容。")
            doc.save(str(input_path))
            self.rewrite_docx_member(input_path, "word/styles.xml", lambda _xml: b"<root/>")

            self.assertFalse(format_academic_paper(str(input_path), str(output_path)))

    def test_format_academic_paper_rebuilds_broken_header_footer_references(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "broken_header_footer_input.docx"
            output_path = temp_path / "broken_header_footer_output.docx"

            doc = Document()
            doc.add_paragraph("页眉页脚关系修复测试")
            doc.add_paragraph("摘要：这是摘要内容。")
            doc.add_paragraph("关键词：页眉 页脚")
            doc.add_paragraph("正文内容。")
            sect_pr = doc.sections[0]._sectPr
            header_reference = OxmlElement("w:headerReference")
            header_reference.set(qn("w:type"), "default")
            header_reference.set(qn("r:id"), "rIdMissingHeader")
            footer_reference = OxmlElement("w:footerReference")
            footer_reference.set(qn("w:type"), "default")
            footer_reference.set(qn("r:id"), "rIdMissingFooter")
            sect_pr.insert(0, footer_reference)
            sect_pr.insert(0, header_reference)
            doc.save(str(input_path))

            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertIsInstance(summary, dict)
            output_doc = Document(str(output_path))
            self.assertEqual(
                output_doc.sections[0].header.paragraphs[0].text,
                summary["page_setup"]["header_text"],
            )
            self.assertIn("PAGE", output_doc.sections[0].footer.paragraphs[0]._element.xml)

    def test_format_academic_paper_rejects_wrong_root_header_footer_parts(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "wrong_root_header_footer_input.docx"
            output_path = temp_path / "wrong_root_header_footer_output.docx"

            doc = Document()
            doc.add_paragraph("页眉页脚 XML 根节点修复测试")
            doc.add_paragraph("摘要：这是摘要内容。")
            doc.add_paragraph("关键词：根节点 修复")
            doc.add_paragraph("正文内容。")
            doc.sections[0].header.paragraphs[0].text = "错误根节点页眉"
            doc.sections[0].footer.paragraphs[0].text = "错误根节点页脚"
            doc.save(str(input_path))

            with ZipFile(input_path, "r") as archive:
                header_members = [
                    name for name in archive.namelist()
                    if re.fullmatch(r"word/header\d+\.xml", name)
                ]
                footer_members = [
                    name for name in archive.namelist()
                    if re.fullmatch(r"word/footer\d+\.xml", name)
                ]
            self.assertEqual(len(header_members), 1)
            self.assertEqual(len(footer_members), 1)

            def replace_root(xml_bytes, root_tag):
                root = etree.fromstring(xml_bytes)
                root.tag = qn(root_tag)
                return etree.tostring(
                    root,
                    encoding="UTF-8",
                    xml_declaration=True,
                    standalone=True,
                )

            self.rewrite_docx_member(
                input_path,
                header_members[0],
                lambda xml_bytes: replace_root(xml_bytes, "w:ftr"),
            )
            self.rewrite_docx_member(
                input_path,
                footer_members[0],
                lambda xml_bytes: replace_root(xml_bytes, "w:hdr"),
            )

            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertFalse(summary)
            self.assertFalse(output_path.exists())

    def test_format_academic_paper_normalizes_invalid_run_and_alignment_values(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "invalid_typed_values_input.docx"
            output_path = temp_path / "invalid_typed_values_output.docx"

            doc = Document()
            doc.add_paragraph("布尔与对齐属性修复测试")
            doc.add_paragraph("摘要：这是摘要内容。")
            doc.add_paragraph("关键词：布尔 对齐")
            body = doc.add_paragraph()
            body_run = body.add_run("正文内容。")
            invalid_bold = OxmlElement("w:b")
            invalid_bold.set(qn("w:val"), "bogus")
            body_run._r.get_or_add_rPr().append(invalid_bold)

            drawing_paragraph = doc.add_paragraph()
            invalid_alignment = OxmlElement("w:jc")
            invalid_alignment.set(qn("w:val"), "bogus")
            drawing_paragraph._element.get_or_add_pPr().append(invalid_alignment)
            drawing = OxmlElement("w:drawing")
            anchor = OxmlElement("wp:anchor")
            extent = OxmlElement("wp:extent")
            extent.set("cx", "914400")
            extent.set("cy", "914400")
            anchor.append(extent)
            drawing.append(anchor)
            drawing_paragraph.add_run()._r.append(drawing)
            doc.save(str(input_path))

            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertIsInstance(summary, dict)
            output_doc = Document(str(output_path))
            output_body = next(paragraph for paragraph in output_doc.paragraphs if paragraph.text == "正文内容。")
            self.assertIsNone(output_body.runs[0].font.bold)
            output_drawing = next(
                paragraph
                for paragraph in output_doc.paragraphs
                if "w:drawing" in paragraph._element.xml
            )
            self.assertEqual(output_drawing.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)

    def test_format_academic_paper_ignores_extreme_existing_numbering_ids(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "extreme_numbering_input.docx"
            output_path = temp_path / "extreme_numbering_output.docx"
            enormous_id = "9" * 5000

            doc = Document()
            doc.add_paragraph("平台治理视角下数字化协同研究")
            doc.add_paragraph("摘要：这是摘要内容")
            doc.add_paragraph("关键词：平台治理 协同")
            doc.add_paragraph("引言", style="Heading 1")
            doc.add_paragraph("正文内容")
            doc.save(str(input_path))

            with ZipFile(input_path, "r") as source:
                entries = source.infolist()
                payloads = {entry.filename: source.read(entry.filename) for entry in entries}

            numbering_root = etree.fromstring(payloads["word/numbering.xml"])
            first_abstract_num = numbering_root.find(qn("w:abstractNum"))
            first_num = numbering_root.find(qn("w:num"))
            self.assertIsNotNone(first_abstract_num)
            self.assertIsNotNone(first_num)

            extreme_abstract_num = etree.fromstring(etree.tostring(first_abstract_num))
            extreme_abstract_num.set(qn("w:abstractNumId"), enormous_id)
            extreme_nsid = extreme_abstract_num.find(qn("w:nsid"))
            self.assertIsNotNone(extreme_nsid)
            extreme_nsid.set(qn("w:val"), format_paper_module.HEADING_NUMBERING_NSID)
            numbering_root.insert(numbering_root.index(first_num), extreme_abstract_num)

            extreme_num = etree.fromstring(etree.tostring(first_num))
            extreme_num.set(qn("w:numId"), enormous_id)
            extreme_num.find(qn("w:abstractNumId")).set(qn("w:val"), enormous_id)
            numbering_root.append(extreme_num)

            cleanup = etree.Element(qn("w:numIdMacAtCleanup"))
            cleanup.set(qn("w:val"), "1")
            numbering_root.append(cleanup)
            payloads["word/numbering.xml"] = etree.tostring(
                numbering_root,
                encoding="UTF-8",
                xml_declaration=True,
                standalone=True,
            )

            with ZipFile(input_path, "w") as target:
                for entry in entries:
                    target.writestr(entry, payloads[entry.filename])

            summary = format_academic_paper(
                str(input_path),
                str(output_path),
                max_output_bytes=64 * 1024,
            )

            self.assertEqual(summary["stats"]["heading_l1"], 1)
            output_doc = Document(str(output_path))
            heading = next(paragraph for paragraph in output_doc.paragraphs if paragraph.text == "引言")
            generated_num_id = heading._element.pPr.find(qn("w:numPr")).find(qn("w:numId")).get(qn("w:val"))
            self.assertTrue(generated_num_id.isdecimal())
            self.assertLessEqual(int(generated_num_id), format_paper_module.MAX_OOXML_DECIMAL_NUMBER)

            with ZipFile(output_path, "r") as archive:
                self.assertIsNone(archive.testzip())
                output_numbering = etree.fromstring(archive.read("word/numbering.xml"))

            generated_nums = [
                num
                for num in output_numbering.findall(qn("w:num"))
                if num.get(qn("w:numId")) == generated_num_id
            ]
            self.assertEqual(len(generated_nums), 1)
            generated_abstract_id = generated_nums[0].find(qn("w:abstractNumId")).get(qn("w:val"))
            self.assertTrue(generated_abstract_id.isdecimal())
            self.assertLessEqual(int(generated_abstract_id), format_paper_module.MAX_OOXML_DECIMAL_NUMBER)
            self.assertEqual(
                len(
                    [
                        abstract_num
                        for abstract_num in output_numbering.findall(qn("w:abstractNum"))
                        if abstract_num.get(qn("w:abstractNumId")) == generated_abstract_id
                    ]
                ),
                1,
            )
            cleanup_element = output_numbering.find(qn("w:numIdMacAtCleanup"))
            self.assertIsNotNone(cleanup_element)
            generated_abstract = next(
                abstract_num
                for abstract_num in output_numbering.findall(qn("w:abstractNum"))
                if abstract_num.get(qn("w:abstractNumId")) == generated_abstract_id
            )
            self.assertLess(output_numbering.index(generated_abstract), output_numbering.index(cleanup_element))
            self.assertLess(output_numbering.index(generated_nums[0]), output_numbering.index(cleanup_element))

    def test_heading_numbering_rejects_incompatible_nsid_collision(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "numbering_nsid_collision_input.docx"
            output_path = temp_path / "numbering_nsid_collision_output.docx"

            doc = Document()
            doc.add_paragraph("编号定义冲突修复测试")
            doc.add_paragraph("摘要：这是摘要内容。")
            doc.add_paragraph("关键词：编号 冲突")
            doc.add_paragraph("引言", style="Heading 1")
            numbering_root = doc.part.numbering_part.numbering_definitions._numbering
            malformed_abstract = OxmlElement("w:abstractNum")
            malformed_abstract.set(qn("w:abstractNumId"), "999")
            nsid = OxmlElement("w:nsid")
            nsid.set(qn("w:val"), format_paper_module.HEADING_NUMBERING_NSID)
            malformed_abstract.append(nsid)
            first_num = numbering_root.find(qn("w:num"))
            numbering_root.insert(numbering_root.index(first_num), malformed_abstract)
            doc.save(str(input_path))

            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertIsInstance(summary, dict)
            output_doc = Document(str(output_path))
            heading = next(paragraph for paragraph in output_doc.paragraphs if paragraph.text == "引言")
            num_id = self.assert_paragraph_has_numbering(heading, 0)
            output_numbering = output_doc.part.numbering_part.numbering_definitions._numbering
            generated_num = output_numbering.num_having_numId(num_id)
            abstract_num_id = generated_num.find(qn("w:abstractNumId")).get(qn("w:val"))
            self.assertNotEqual(abstract_num_id, "999")
            generated_abstract = next(
                abstract_num
                for abstract_num in output_numbering.findall(qn("w:abstractNum"))
                if abstract_num.get(qn("w:abstractNumId")) == abstract_num_id
            )
            self.assertTrue(format_paper_module._is_compatible_heading_numbering_abstract(generated_abstract))

    def test_heading_numbering_num_id_allocation_uses_document_cache(self):
        doc = Document()
        format_paper_module._create_heading_numbering_instance(doc, (1,))

        with patch.object(
            format_paper_module,
            "_parse_bounded_decimal",
            wraps=format_paper_module._parse_bounded_decimal,
        ) as parser:
            for index in range(1, 201):
                format_paper_module._create_heading_numbering_instance(doc, (index,))

        self.assertEqual(parser.call_count, 0)
        numbering_root = doc.part.numbering_part.numbering_definitions._numbering
        self.assertGreaterEqual(len(numbering_root.findall(qn("w:num"))), 201)

    def test_format_academic_paper_does_not_auto_number_plain_short_body(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "plain_short_body_input.docx"
            output_path = temp_path / "plain_short_body_output.docx"

            doc = Document()
            doc.add_paragraph("短标题误判保护测试")
            doc.add_paragraph("摘要：这是摘要内容")
            doc.add_paragraph("关键词：误判 测试")
            doc.add_paragraph("研究意义")
            doc.add_paragraph("这里是正文展开内容。")
            doc.save(str(input_path))

            summary = format_academic_paper(
                str(input_path),
                str(output_path),
                max_output_bytes=64 * 1024,
            )

            self.assertEqual(summary["stats"]["heading_l1"], 0)
            self.assertEqual(summary["stats"]["heading_l2"], 0)
            self.assertEqual(summary["stats"]["heading_l3"], 0)

            output_doc = Document(str(output_path))
            paragraph = next(item for item in output_doc.paragraphs if item.text == "研究意义")
            num_pr = paragraph._element.pPr.find(qn("w:numPr")) if paragraph._element.pPr is not None else None
            self.assertIsNone(num_pr)
            self.assertEqual(paragraph.paragraph_format.first_line_indent.pt, 24.0)

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

            summary = format_academic_paper(
                str(input_path),
                str(output_path),
                max_output_bytes=64 * 1024,
            )

            self.assertEqual(summary["formatted_footnotes"], 1)

            with ZipFile(output_path, "r") as archive:
                footnotes_xml = archive.read("word/footnotes.xml")

            footnotes_root = etree.fromstring(footnotes_xml)
            footnote = next(
                node
                for node in footnotes_root.findall(qn("w:footnote"))
                if node.get(qn("w:id")) == "2"
            )
            footnote_paragraph = footnote.find(qn("w:p"))
            runs = footnote.findall(".//" + qn("w:r"))
            reference_rpr = runs[0].find(qn("w:rPr"))
            content_rpr = runs[1].find(qn("w:rPr"))
            footnote_ppr = footnote_paragraph.find(qn("w:pPr"))
            footnote_spacing = footnote_ppr.find(qn("w:spacing"))
            footnote_indent = footnote_ppr.find(qn("w:ind"))
            footnote_jc = footnote_ppr.find(qn("w:jc"))

            self.assertEqual(reference_rpr.find(qn("w:rStyle")).get(qn("w:val")), "FootnoteReference")
            self.assertEqual(reference_rpr.find(qn("w:vertAlign")).get(qn("w:val")), "superscript")
            self.assertEqual(reference_rpr.find(qn("w:sz")).get(qn("w:val")), "20")
            self.assertEqual(content_rpr.find(qn("w:rFonts")).get(qn("w:eastAsia")), "宋体")
            self.assertEqual(content_rpr.find(qn("w:rFonts")).get(qn("w:ascii")), "Times New Roman")
            self.assertEqual(content_rpr.find(qn("w:sz")).get(qn("w:val")), "20")
            self.assertEqual(footnote_spacing.get(qn("w:before")), "0")
            self.assertEqual(footnote_spacing.get(qn("w:after")), "0")
            self.assertEqual(footnote_spacing.get(qn("w:line")), "240")
            self.assertEqual(footnote_spacing.get(qn("w:lineRule")), "auto")
            self.assertEqual(footnote_indent.get(qn("w:left")), "0")
            self.assertEqual(footnote_indent.get(qn("w:right")), "0")
            self.assertEqual(footnote_indent.get(qn("w:firstLine")), "0")
            self.assertEqual(footnote_jc.get(qn("w:val")), "left")
            self.assertEqual(footnote_ppr.find(qn("w:widowControl")).get(qn("w:val")), "true")

    def test_format_academic_paper_formats_relocated_footnote_part(self):
        relocated_name = "word/notes/footnotes2.xml"
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "relocated-footnote-input.docx"
            output_path = temp_path / "relocated-footnote-output.docx"

            doc = Document()
            doc.add_paragraph("重定位脚注排版测试")
            doc.add_paragraph("摘要：这是摘要内容")
            doc.add_paragraph("关键词：脚注 重定位")
            doc.add_paragraph("正文里有一个脚注引用")
            doc.save(str(input_path))
            self.inject_simple_footnote(input_path, "重定位脚注内容")
            self.relocate_simple_footnote_part(input_path, relocated_name)

            self.assertEqual(len(Document(str(input_path)).paragraphs), 4)
            self.assertEqual(
                format_paper_module._resolve_footnote_part_name(input_path),
                relocated_name,
            )

            summary = format_academic_paper(
                str(input_path),
                str(output_path),
                max_output_bytes=64 * 1024,
            )

            self.assertIsInstance(summary, dict)
            self.assertEqual(summary["formatted_footnotes"], 1)
            self.assertEqual(len(Document(str(output_path)).paragraphs), 4)

            with ZipFile(output_path, "r") as archive:
                self.assertIsNone(archive.testzip())
                self.assertIn(relocated_name, archive.namelist())
                self.assertNotIn("word/footnotes.xml", archive.namelist())
                footnotes_root = etree.fromstring(archive.read(relocated_name))

            content_run = next(
                footnote.findall(".//" + qn("w:r"))[1]
                for footnote in footnotes_root.findall(qn("w:footnote"))
                if footnote.get(qn("w:id")) == "2"
            )
            run_properties = content_run.find(qn("w:rPr"))
            self.assertEqual(run_properties.find(qn("w:sz")).get(qn("w:val")), "20")
            self.assertEqual(
                run_properties.find(qn("w:rFonts")).get(qn("w:eastAsia")),
                "宋体",
            )

    def test_format_academic_paper_writes_ooxml_properties_in_schema_order(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "property_order_input.docx"
            output_path = temp_path / "property_order_output.docx"

            doc = Document()
            doc.add_paragraph("学术论文属性顺序测试")
            doc.add_paragraph("摘要：这是摘要内容。")
            doc.add_paragraph("关键词：属性 顺序")
            doc.add_paragraph("1 引言")
            doc.add_paragraph("正文内容。")
            table = doc.add_table(rows=2, cols=2)
            table.cell(0, 0).text = "变量"
            table.cell(0, 1).text = "取值"
            table.cell(1, 0).text = "A"
            table.cell(1, 1).text = "1"
            doc.add_section(WD_SECTION.NEW_PAGE)
            doc.add_paragraph("第二节正文。")
            doc.save(str(input_path))
            self.inject_simple_footnote(input_path, "属性顺序脚注")

            def scramble_footnote_property_order(xml_bytes):
                root = etree.fromstring(xml_bytes)
                footnote = next(
                    node
                    for node in root.findall(qn("w:footnote"))
                    if node.get(qn("w:id")) == "2"
                )
                paragraph = footnote.find(qn("w:p"))
                paragraph.insert(
                    0,
                    parse_xml(
                        f'<w:pPr {nsdecls("w")}>'
                        '<w:jc w:val="left"/><w:spacing w:before="0" w:after="0" '
                        'w:line="240" w:lineRule="auto"/>'
                        '<w:spacing w:before="99"/></w:pPr>'
                    ),
                )
                content_run = paragraph.findall(qn("w:r"))[1]
                content_run.insert(
                    0,
                    parse_xml(
                        f'<w:rPr {nsdecls("w")}><w:sz w:val="20"/>'
                        '<w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" '
                        'w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
                        '<w:sz w:val="99"/></w:rPr>'
                    ),
                )
                return etree.tostring(
                    root,
                    encoding="UTF-8",
                    xml_declaration=True,
                    standalone=True,
                )

            self.rewrite_docx_member(
                input_path,
                "word/footnotes.xml",
                scramble_footnote_property_order,
            )

            summary = format_academic_paper(
                str(input_path),
                str(output_path),
                cover_info={
                    "cover_title": "课程论文",
                    "title": "学术论文属性顺序测试",
                    "student_name": "测试学生",
                },
            )

            self.assertIsInstance(summary, dict)
            with ZipFile(output_path, "r") as archive:
                document_root = etree.fromstring(archive.read("word/document.xml"))
                footnotes_root = etree.fromstring(archive.read("word/footnotes.xml"))

            p_pr_order = (
                "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore",
                "w:framePr", "w:widowControl", "w:numPr", "w:suppressLineNumbers",
                "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
                "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
                "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
                "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
                "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
                "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
                "w:rPr", "w:sectPr", "w:pPrChange",
            )
            r_pr_order = (
                "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs",
                "w:caps", "w:smallCaps", "w:strike", "w:dstrike", "w:outline",
                "w:shadow", "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid",
                "w:vanish", "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern",
                "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
                "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs",
                "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath",
            )
            tc_pr_order = (
                "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge",
                "w:tcBorders", "w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
                "w:tcFitText", "w:vAlign", "w:hideMark", "w:headers", "w:cellIns",
                "w:cellDel", "w:cellMerge", "w:tcPrChange",
            )
            tr_pr_order = (
                "w:cnfStyle", "w:divId", "w:gridBefore", "w:gridAfter", "w:wBefore",
                "w:wAfter", "w:cantSplit", "w:trHeight", "w:tblHeader",
                "w:tblCellSpacing", "w:jc", "w:hidden", "w:ins", "w:del", "w:trPrChange",
            )
            tbl_pr_order = (
                "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual",
                "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc",
                "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
                "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription", "w:tblPrChange",
            )
            border_order = (
                "w:top", "w:start", "w:left", "w:bottom", "w:end", "w:right",
                "w:insideH", "w:insideV", "w:between", "w:bar", "w:tl2br", "w:tr2bl",
            )

            for p_pr in document_root.findall(".//" + qn("w:pPr")):
                self.assert_xml_children_follow_order(p_pr, p_pr_order)
            for tc_pr in document_root.findall(".//" + qn("w:tcPr")):
                self.assert_xml_children_follow_order(tc_pr, tc_pr_order)
            for tr_pr in document_root.findall(".//" + qn("w:trPr")):
                self.assert_xml_children_follow_order(tr_pr, tr_pr_order)
            for tbl_pr in document_root.findall(".//" + qn("w:tblPr")):
                self.assert_xml_children_follow_order(tbl_pr, tbl_pr_order)
            for border_tag in ("w:tblBorders", "w:tcBorders", "w:pBdr"):
                for borders in document_root.findall(".//" + qn(border_tag)):
                    self.assert_xml_children_follow_order(borders, border_order)
            for p_pr in footnotes_root.findall(".//" + qn("w:pPr")):
                self.assert_xml_children_follow_order(p_pr, p_pr_order)
            for r_pr in footnotes_root.findall(".//" + qn("w:rPr")):
                self.assert_xml_children_follow_order(r_pr, r_pr_order)

            formatted_footnote = next(
                node
                for node in footnotes_root.findall(qn("w:footnote"))
                if node.get(qn("w:id")) == "2"
            )
            formatted_paragraph = formatted_footnote.find(qn("w:p"))
            self.assertEqual(len(formatted_paragraph.findall(qn("w:pPr"))), 1)
            self.assertEqual(
                len(formatted_paragraph.find(qn("w:pPr")).findall(qn("w:spacing"))),
                1,
            )
            content_run = formatted_paragraph.findall(qn("w:r"))[1]
            self.assertEqual(len(content_run.findall(qn("w:rPr"))), 1)
            self.assertEqual(
                len(content_run.find(qn("w:rPr")).findall(qn("w:sz"))),
                1,
            )

    def test_format_academic_paper_repairs_invalid_table_grid_style_reference(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "invalid_table_grid_style_input.docx"
            output_path = temp_path / "invalid_table_grid_style_output.docx"

            doc = Document()
            doc.add_paragraph("三线表样式修复测试")
            doc.add_paragraph("摘要：这是摘要内容。")
            doc.add_paragraph("关键词：三线表 样式")
            table = doc.add_table(rows=2, cols=2)
            table.style = "Table Grid"
            table.cell(0, 0).text = "变量"
            table.cell(0, 1).text = "取值"
            table.cell(1, 0).text = "A"
            table.cell(1, 1).text = "1"
            doc.save(str(input_path))

            def invalidate_table_grid_style(xml_bytes):
                root = etree.fromstring(xml_bytes)
                table_grid = next(
                    style
                    for style in root.findall(qn("w:style"))
                    if style.get(qn("w:styleId")) == "TableGrid"
                )
                table_grid.set(qn("w:type"), "bogus")
                return etree.tostring(
                    root,
                    encoding="UTF-8",
                    xml_declaration=True,
                    standalone=True,
                )

            self.rewrite_docx_member(
                input_path,
                "word/styles.xml",
                invalidate_table_grid_style,
            )
            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertIsInstance(summary, dict)
            output_doc = Document(str(output_path))
            output_table = output_doc.tables[0]
            self.assertIsNone(output_table._tbl.tblPr.find(qn("w:tblBorders")))
            self.assertEqual(
                output_table.cell(0, 0)._tc.tcPr.find(qn("w:tcBorders"))
                .find(qn("w:top")).get(qn("w:val")),
                "single",
            )
            self.assertEqual(
                output_table.cell(1, 0)._tc.tcPr.find(qn("w:tcBorders"))
                .find(qn("w:bottom")).get(qn("w:val")),
                "single",
            )

            with ZipFile(output_path, "r") as archive:
                document_root = etree.fromstring(archive.read("word/document.xml"))
                styles_root = etree.fromstring(archive.read("word/styles.xml"))
            tbl_style = document_root.find(".//" + qn("w:tblStyle"))
            if tbl_style is not None:
                referenced_style_id = tbl_style.get(qn("w:val"))
                referenced_styles = [
                    style
                    for style in styles_root.findall(qn("w:style"))
                    if style.get(qn("w:styleId")) == referenced_style_id
                ]
                self.assertEqual(len(referenced_styles), 1)
                self.assertEqual(referenced_styles[0].get(qn("w:type")), "table")

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
            doc.add_paragraph("注：样本区间为 2020-2024 年。")
            doc.add_paragraph("来源：Python 爬取与企业年报整理。")

            doc.add_paragraph("参考文献")
            doc.add_paragraph("[9] Smith, John. Python-based trade analytics[J]. 2024.")
            doc.save(str(input_path))

            summary = format_academic_paper(str(input_path), str(output_path))
            self.assertIsInstance(summary, dict)
            self.assertEqual(summary["stats"]["figure_caption"], 2)
            self.assertEqual(summary["stats"]["table_caption"], 1)
            self.assertEqual(summary["stats"]["caption_note"], 2)
            self.assertEqual(summary["table_paragraphs"], 4)

            output_doc = Document(str(output_path))
            self.assertEqual(len(output_doc.inline_shapes), 2)

            self.assertEqual(output_doc.paragraphs[3].text, "引言")
            self.assertEqual(output_doc.paragraphs[7].text, "图 1 系统架构图")
            self.assertEqual(output_doc.paragraphs[9].text, "图 2 模型流程图")
            self.assertEqual(output_doc.paragraphs[10].text, "表 1 样本描述统计")
            self.assertEqual(output_doc.paragraphs[11].text, "注：样本区间为 2020-2024 年。")
            self.assertEqual(output_doc.paragraphs[12].text, "来源：Python 爬取与企业年报整理。")
            self.assertEqual(output_doc.paragraphs[5].style.name, "List Bullet")
            self.assertEqual(output_doc.paragraphs[6].paragraph_format.first_line_indent.pt, 0.0)
            self.assert_paragraph_has_numbering(output_doc.paragraphs[3], 0)
            self.assertEqual(output_doc.paragraphs[3]._element.pPr.find(qn("w:keepNext")).get(qn("w:val")), "true")
            self.assertEqual(output_doc.paragraphs[3]._element.pPr.find(qn("w:keepLines")).get(qn("w:val")), "true")
            self.assertEqual(output_doc.paragraphs[6]._element.pPr.find(qn("w:keepNext")).get(qn("w:val")), "true")
            self.assertEqual(output_doc.paragraphs[7]._element.pPr.find(qn("w:keepLines")).get(qn("w:val")), "true")
            self.assertIsNone(output_doc.paragraphs[7]._element.pPr.find(qn("w:keepNext")))
            self.assertEqual(output_doc.paragraphs[10]._element.pPr.find(qn("w:keepNext")).get(qn("w:val")), "true")
            self.assertEqual(output_doc.paragraphs[4]._element.pPr.find(qn("w:widowControl")).get(qn("w:val")), "true")
            self.assertEqual(output_doc.paragraphs[11]._element.pPr.find(qn("w:widowControl")).get(qn("w:val")), "true")
            self.assertEqual(output_doc.paragraphs[14]._element.pPr.find(qn("w:widowControl")).get(qn("w:val")), "true")

            body_run = output_doc.paragraphs[4].runs[0]
            table_run = output_doc.tables[0].cell(1, 1).paragraphs[0].runs[0]
            self.assert_run_uses_mixed_font_pair(self, body_run)
            self.assert_run_uses_mixed_font_pair(self, table_run)

            caption_note = output_doc.paragraphs[11]
            source_note = output_doc.paragraphs[12]
            reference_entry = output_doc.paragraphs[14]
            self.assertEqual(caption_note.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.LEFT)
            self.assertAlmostEqual(caption_note.paragraph_format.first_line_indent.pt, 0.0, places=1)
            self.assertEqual(caption_note.paragraph_format.line_spacing, 1.0)
            self.assertTrue(caption_note.runs[0].font.bold)
            self.assertEqual(caption_note.runs[0].font.size.pt, 10.5)
            self.assertEqual(source_note.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.LEFT)
            self.assertTrue(source_note.runs[0].font.bold)
            self.assertEqual(source_note.runs[0].font.size.pt, 10.5)
            self.assertAlmostEqual(reference_entry.paragraph_format.left_indent.pt, 0.0, places=1)
            self.assertAlmostEqual(reference_entry.paragraph_format.first_line_indent.pt, 21.0, places=1)
            self.assertEqual(reference_entry.paragraph_format.line_spacing, 1.0)
            self.assertEqual(reference_entry.runs[0].font.size.pt, 10.0)

            tbl_borders = output_doc.tables[0]._tbl.tblPr.find(qn("w:tblBorders"))
            top_left_borders = output_doc.tables[0].cell(0, 0)._tc.tcPr.find(qn("w:tcBorders"))
            bottom_left_borders = output_doc.tables[0].cell(1, 0)._tc.tcPr.find(qn("w:tcBorders"))
            first_row_tr_pr = output_doc.tables[0].rows[0]._tr.trPr
            second_row_tr_pr = output_doc.tables[0].rows[1]._tr.trPr
            header_paragraph = output_doc.tables[0].cell(0, 0).paragraphs[0]
            self.assertIsNone(tbl_borders)
            self.assertEqual(output_doc.tables[0].alignment, WD_TABLE_ALIGNMENT.CENTER)
            self.assertEqual(output_doc.tables[0].cell(0, 0).vertical_alignment, WD_CELL_VERTICAL_ALIGNMENT.CENTER)
            self.assertEqual(top_left_borders.find(qn("w:top")).get(qn("w:val")), "single")
            self.assertEqual(top_left_borders.find(qn("w:bottom")).get(qn("w:val")), "single")
            self.assertEqual(top_left_borders.find(qn("w:left")).get(qn("w:val")), "none")
            self.assertEqual(top_left_borders.find(qn("w:right")).get(qn("w:val")), "none")
            self.assertEqual(bottom_left_borders.find(qn("w:top")).get(qn("w:val")), "none")
            self.assertEqual(bottom_left_borders.find(qn("w:bottom")).get(qn("w:val")), "single")
            self.assertEqual(header_paragraph.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertTrue(header_paragraph.runs[0].font.bold)
            self.assertEqual(first_row_tr_pr.find(qn("w:tblHeader")).get(qn("w:val")), "true")
            self.assertEqual(first_row_tr_pr.find(qn("w:cantSplit")).get(qn("w:val")), "true")
            self.assertEqual(second_row_tr_pr.find(qn("w:cantSplit")).get(qn("w:val")), "true")

    def test_format_academic_paper_preserves_images_in_mixed_special_paragraphs(self):
        tiny_png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            image_path = temp_path / "pixel.png"
            input_path = temp_path / "mixed-content-input.docx"
            output_path = temp_path / "mixed-content-output.docx"
            image_path.write_bytes(tiny_png)

            doc = Document()
            doc.add_paragraph("混合内容保全测试")

            abstract = doc.add_paragraph()
            abstract.add_run("摘要：正文旁")
            abstract.add_run().add_picture(str(image_path), width=Inches(0.2))
            abstract._element.append(
                parse_xml(
                    r"""
                    <m:oMathPara %s>
                      <m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath>
                    </m:oMathPara>
                    """ % nsdecls("m")
                )
            )
            abstract.add_run("包含内嵌图片。")
            doc.add_paragraph("关键词：图片 保全")

            heading = doc.add_paragraph()
            heading.add_run("1 含")
            heading.add_run().add_picture(str(image_path), width=Inches(0.2))
            heading.add_run("图标题")
            doc.add_paragraph("后续推断标题", style="Heading 2")
            doc.add_paragraph("正文内容。")

            caption = doc.add_paragraph()
            caption.add_run("图9 含")
            caption.add_run().add_picture(str(image_path), width=Inches(0.2))
            caption.add_run("图图题")
            doc.save(str(input_path))

            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertIsInstance(summary, dict)
            self.assertEqual(summary["stats"]["abstract"], 1)
            self.assertEqual(summary["stats"]["heading_l1"], 1)
            self.assertEqual(summary["stats"]["heading_l2"], 1)
            self.assertEqual(summary["stats"]["figure_caption"], 1)
            self.assertEqual(summary["equation_paragraphs"], 1)

            output_doc = Document(str(output_path))
            self.assertEqual(len(output_doc.inline_shapes), 3)

            output_abstract = next(
                paragraph for paragraph in output_doc.paragraphs
                if paragraph.text == "摘要：正文旁包含内嵌图片。"
            )
            output_heading = next(
                paragraph for paragraph in output_doc.paragraphs
                if paragraph.text == "1 含图标题"
            )
            inferred_heading = next(
                paragraph for paragraph in output_doc.paragraphs
                if paragraph.text == "后续推断标题"
            )
            output_caption = next(
                paragraph for paragraph in output_doc.paragraphs
                if paragraph.text == "图9 含图图题"
            )

            for paragraph in (output_abstract, output_heading, output_caption):
                self.assertEqual(len(paragraph._element.findall(".//" + qn("w:drawing"))), 1)

            self.assertEqual(
                [run.text for run in output_abstract.runs],
                ["摘要：正文旁", "", "包含内嵌图片。"],
            )
            self.assertEqual(
                [
                    child.tag
                    for child in output_abstract._element
                    if child.tag in {qn("w:r"), qn("m:oMathPara")}
                ],
                [qn("w:r"), qn("w:r"), qn("m:oMathPara"), qn("w:r")],
            )
            self.assertEqual(
                [run.text for run in output_heading.runs],
                ["1 含", "", "图标题"],
            )
            self.assertEqual(
                [run.text for run in output_caption.runs],
                ["图9 含", "", "图图题"],
            )
            self.assertIsNone(output_heading._element.pPr.find(qn("w:numPr")))
            inferred_num_id = self.assert_paragraph_has_numbering(inferred_heading, 1)
            self.assert_numbering_overrides(output_doc, inferred_num_id, {0: 1, 1: 1})
            self.assertTrue(output_heading.runs[0].font.bold)
            self.assertEqual(output_caption.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)

            with ZipFile(output_path, "r") as archive:
                self.assertIsNone(archive.testzip())

    def test_table_iterators_do_not_repeat_merged_cells_or_nested_tables(self):
        doc = Document()
        outer_table = doc.add_table(rows=1, cols=4)
        merged_outer_cell = outer_table.cell(0, 0).merge(outer_table.cell(0, 3))
        merged_outer_cell.text = "外层合并单元格"
        nested_table = merged_outer_cell.add_table(rows=1, cols=4)
        nested_table.cell(0, 0).merge(nested_table.cell(0, 3)).text = "内层合并单元格"

        unique_outer_cells = list(format_paper_module.iter_unique_row_cells(outer_table.rows[0]))
        all_tables = list(format_paper_module.iter_all_tables(doc.tables))
        table_paragraphs = list(format_paper_module.iter_table_paragraphs(doc.tables))

        self.assertEqual(len(unique_outer_cells), 1)
        self.assertEqual(len(all_tables), 2)
        self.assertIs(all_tables[0]._tbl, outer_table._tbl)
        self.assertIs(all_tables[1]._tbl, nested_table._tbl)
        paragraph_elements = [id(paragraph._element) for paragraph in table_paragraphs]
        self.assertEqual(len(paragraph_elements), len(set(paragraph_elements)))

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

    def test_constrain_inline_images_skips_extreme_anchor_coordinates(self):
        doc = Document()
        paragraph = doc.add_paragraph()
        enormous_coordinate = "9" * 5000
        max_width = format_paper_module.get_max_printable_width(doc)

        def append_anchor_extent(cx: str, cy: str):
            drawing = OxmlElement("w:drawing")
            anchor = OxmlElement("wp:anchor")
            extent = OxmlElement("wp:extent")
            extent.set("cx", cx)
            extent.set("cy", cy)
            anchor.append(extent)
            drawing.append(anchor)
            paragraph.add_run()._r.append(drawing)
            return extent

        extreme_cx = append_anchor_extent(enormous_coordinate, "914400")
        extreme_cy = append_anchor_extent(str(max_width + 1), enormous_coordinate)

        self.assertEqual(format_paper_module.constrain_inline_images(doc), 0)
        self.assertEqual(extreme_cx.get("cx"), enormous_coordinate)
        self.assertEqual(extreme_cx.get("cy"), "914400")
        self.assertEqual(extreme_cy.get("cx"), str(max_width + 1))
        self.assertEqual(extreme_cy.get("cy"), enormous_coordinate)

    def test_format_academic_paper_skips_extreme_anchor_coordinates_end_to_end(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "extreme_anchor_input.docx"
            output_path = temp_path / "extreme_anchor_output.docx"
            enormous_coordinate = "9" * 5000

            doc = Document()
            doc.add_paragraph("浮动图坐标边界测试")
            doc.add_paragraph("摘要：这是摘要内容。")
            doc.add_paragraph("关键词：浮动图 坐标 测试")
            drawing_paragraph = doc.add_paragraph()
            drawing = OxmlElement("w:drawing")
            anchor = OxmlElement("wp:anchor")
            extent = OxmlElement("wp:extent")
            extent.set("cx", enormous_coordinate)
            extent.set("cy", "914400")
            anchor.append(extent)
            drawing.append(anchor)
            drawing_paragraph.add_run()._r.append(drawing)
            doc.save(str(input_path))

            summary = format_academic_paper(
                str(input_path),
                str(output_path),
                max_output_bytes=64 * 1024,
            )

            self.assertIsInstance(summary, dict)
            self.assertEqual(summary["resized_images"], 0)
            output_doc = Document(str(output_path))
            output_extent = output_doc._element.find(".//" + qn("wp:anchor") + "/" + qn("wp:extent"))
            self.assertIsNotNone(output_extent)
            self.assertEqual(output_extent.get("cx"), enormous_coordinate)
            self.assertEqual(output_extent.get("cy"), "914400")

    def test_constrain_inline_images_skips_invalid_inline_and_vml_dimensions(self):
        doc = Document()
        paragraph = doc.add_paragraph()
        enormous_value = "9" * 5000
        max_width = format_paper_module.get_max_printable_width(doc)

        def append_inline_extent(cx: str, cy: str):
            drawing = OxmlElement("w:drawing")
            inline = OxmlElement("wp:inline")
            extent = OxmlElement("wp:extent")
            extent.set("cx", cx)
            extent.set("cy", cy)
            inline.append(extent)
            drawing.append(inline)
            paragraph.add_run()._r.append(drawing)
            return extent

        enormous_extent = append_inline_extent(enormous_value, "914400")
        out_of_range_extent = append_inline_extent(
            str(max_width + 1),
            str(format_paper_module.MAX_OOXML_COORDINATE + 1),
        )

        vml_shape_tag = "{urn:schemas-microsoft-com:vml}shape"
        malformed_vml = etree.Element(vml_shape_tag)
        malformed_vml.set("style", "width:1..2pt;height:10pt")
        paragraph.add_run()._r.append(malformed_vml)
        enormous_vml = etree.Element(vml_shape_tag)
        enormous_vml_style = f"width:{enormous_value}pt;height:10pt"
        enormous_vml.set("style", enormous_vml_style)
        paragraph.add_run()._r.append(enormous_vml)
        enormous_height_vml = etree.Element(vml_shape_tag)
        enormous_height_vml_style = f"width:10pt;height:{enormous_value}pt"
        enormous_height_vml.set("style", enormous_height_vml_style)
        paragraph.add_run()._r.append(enormous_height_vml)

        self.assertEqual(format_paper_module.constrain_inline_images(doc), 0)
        self.assertEqual(enormous_extent.get("cx"), enormous_value)
        self.assertEqual(out_of_range_extent.get("cy"), str(format_paper_module.MAX_OOXML_COORDINATE + 1))
        self.assertEqual(malformed_vml.get("style"), "width:1..2pt;height:10pt")
        self.assertEqual(enormous_vml.get("style"), enormous_vml_style)
        self.assertEqual(enormous_height_vml.get("style"), enormous_height_vml_style)

    def test_constrain_inline_images_resizes_incomplete_inline_without_graphic(self):
        doc = Document()
        paragraph = doc.add_paragraph()
        max_width = format_paper_module.get_max_printable_width(doc)
        drawing = OxmlElement("w:drawing")
        inline = OxmlElement("wp:inline")
        extent = OxmlElement("wp:extent")
        extent.set("cx", str(max_width * 2))
        extent.set("cy", str(max_width))
        inline.append(extent)
        drawing.append(inline)
        paragraph.add_run()._r.append(drawing)

        self.assertEqual(format_paper_module.constrain_inline_images(doc), 1)
        self.assertEqual(extent.get("cx"), str(max_width))
        self.assertEqual(extent.get("cy"), str(max_width // 2))

    def test_constrain_inline_images_scales_valid_vml_dimensions(self):
        doc = Document()
        paragraph = doc.add_paragraph()
        vml_shape = etree.Element("{urn:schemas-microsoft-com:vml}shape")
        vml_shape.set("style", "position:absolute;width:1000PT;height:500PT")
        paragraph.add_run()._r.append(vml_shape)

        self.assertEqual(format_paper_module.constrain_inline_images(doc), 1)

        style = vml_shape.get("style")
        width_match = re.search(r"width:([0-9.]+)pt", style)
        height_match = re.search(r"height:([0-9.]+)pt", style)
        self.assertIsNotNone(width_match)
        self.assertIsNotNone(height_match)
        width = float(width_match.group(1))
        height = float(height_match.group(1))
        self.assertAlmostEqual(width, format_paper_module.get_max_printable_width(doc) / 12700, places=2)
        self.assertAlmostEqual(height / width, 0.5, places=2)

    def test_format_academic_paper_does_not_misclassify_body_explanations_as_caption_notes(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / "body_explanation_input.docx"
            output_path = temp_path / "body_explanation_output.docx"

            doc = Document()
            doc.add_paragraph("平台治理视角下数字化协同研究")
            doc.add_paragraph("摘要：这是摘要内容。")
            doc.add_paragraph("关键词：平台治理 协同")
            doc.add_paragraph("说明：这是正文中的说明句，不应被当作图表附注。")
            doc.save(str(input_path))

            summary = format_academic_paper(str(input_path), str(output_path))

            self.assertEqual(summary["stats"]["caption_note"], 0)

            output_doc = Document(str(output_path))
            body_paragraph = output_doc.paragraphs[3]
            self.assertAlmostEqual(body_paragraph.paragraph_format.first_line_indent.pt, 24.0, places=1)
            self.assertEqual(body_paragraph.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)
            self.assertEqual(body_paragraph.runs[0].font.size.pt, 12.0)
            self.assertFalse(body_paragraph.runs[0].font.bold)

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

    def test_generate_cover_page_detaches_shared_first_header_footer_parts(self):
        doc = Document()
        doc.add_paragraph("正文内容")
        apply_document_layout(doc, "正文运行页眉")
        section = doc.sections[0]

        original_default_ids = {}
        for tag in ("w:headerReference", "w:footerReference"):
            default_reference = next(
                reference
                for reference in section._sectPr.findall(qn(tag))
                if reference.get(qn("w:type")) == "default"
            )
            original_default_ids[tag] = default_reference.get(qn("r:id"))
            shared_first_reference = OxmlElement(tag)
            shared_first_reference.set(qn("w:type"), "first")
            shared_first_reference.set(
                qn("r:id"),
                original_default_ids[tag],
            )
            section._sectPr.append(shared_first_reference)

        self.assertIs(section.first_page_header.part, section.header.part)
        self.assertIs(section.first_page_footer.part, section.footer.part)

        self.assertTrue(
            generate_cover_page(
                doc,
                {
                    "title": "共享页眉页脚测试论文",
                    "cover_title": "课程论文",
                },
            )
        )

        self.assertIsNot(section.first_page_header.part, section.header.part)
        self.assertIsNot(section.first_page_footer.part, section.footer.part)
        self.assertEqual(section.header.paragraphs[0].text, "正文运行页眉")
        self.assertIn("PAGE", section.footer._element.xml)
        self.assertEqual(section.first_page_header.paragraphs[0].text, "")
        self.assertEqual(section.first_page_footer.paragraphs[0].text, "")

        for tag in ("w:headerReference", "w:footerReference"):
            references = {
                reference.get(qn("w:type")): reference.get(qn("r:id"))
                for reference in section._sectPr.findall(qn(tag))
            }
            self.assertEqual(references["default"], original_default_ids[tag])
            self.assertNotEqual(references["first"], references["default"])
            self.assertIn(references["default"], section._document_part.rels)
            self.assertIn(references["first"], section._document_part.rels)

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            output_path = Path(handle.name)
        try:
            doc.save(str(output_path))
            reopened = Document(str(output_path))
            reopened_section = reopened.sections[0]
            self.assertEqual(
                reopened_section.header.paragraphs[0].text,
                "正文运行页眉",
            )
            self.assertIn("PAGE", reopened_section.footer._element.xml)
            self.assertEqual(reopened_section.first_page_header.paragraphs[0].text, "")
            self.assertEqual(reopened_section.first_page_footer.paragraphs[0].text, "")
        finally:
            output_path.unlink(missing_ok=True)

    def test_generate_cover_page_detaches_distinct_relationships_targeting_shared_parts(self):
        doc = Document()
        doc.add_paragraph("正文内容")
        apply_document_layout(doc, "正文运行页眉")
        section = doc.sections[0]
        document_part = section._document_part
        obsolete_first_relationships = []

        for tag in ("w:headerReference", "w:footerReference"):
            default_reference = next(
                reference
                for reference in section._sectPr.findall(qn(tag))
                if reference.get(qn("w:type")) == "default"
            )
            default_relationship = document_part.rels[
                default_reference.get(qn("r:id"))
            ]
            shared_first_rid = document_part.rels._next_rId
            document_part.rels.add_relationship(
                default_relationship.reltype,
                default_relationship.target_part,
                shared_first_rid,
            )
            obsolete_first_relationships.append(
                (shared_first_rid, document_part.rels[shared_first_rid])
            )

            shared_first_reference = OxmlElement(tag)
            shared_first_reference.set(qn("w:type"), "first")
            shared_first_reference.set(qn("r:id"), shared_first_rid)
            section._sectPr.append(shared_first_reference)

        self.assertIs(section.first_page_header.part, section.header.part)
        self.assertIs(section.first_page_footer.part, section.footer.part)

        self.assertTrue(
            generate_cover_page(
                doc,
                {
                    "title": "不同关系共享 Part 测试论文",
                    "cover_title": "课程论文",
                },
            )
        )

        self.assertIsNot(section.first_page_header.part, section.header.part)
        self.assertIsNot(section.first_page_footer.part, section.footer.part)
        self.assertEqual(section.header.paragraphs[0].text, "正文运行页眉")
        self.assertIn("PAGE", section.footer._element.xml)
        self.assertEqual(section.first_page_header.paragraphs[0].text, "")
        self.assertEqual(section.first_page_footer.paragraphs[0].text, "")
        self.assertTrue(
            all(
                document_part.rels.get(rid) is not obsolete_relationship
                for rid, obsolete_relationship in obsolete_first_relationships
            )
        )

    def test_ensure_document_ends_with_page_break_appends_break(self):
        doc = Document()
        doc.add_paragraph("封面内容")

        ensure_document_ends_with_page_break(doc)

        self.assertTrue(doc.paragraphs)
        self.assertIn('w:type="page"', doc.paragraphs[-1]._element.xml)

    def test_ensure_document_ends_with_page_break_checks_real_terminal_content(self):
        break_then_text = Document()
        mixed_paragraph = break_then_text.add_paragraph()
        mixed_paragraph.add_run().add_break(format_paper_module.WD_BREAK.PAGE)
        mixed_paragraph.add_run("分页后的封面内容")

        ensure_document_ends_with_page_break(break_then_text)

        self.assertEqual(len(break_then_text.paragraphs), 2)
        self.assertEqual(break_then_text.paragraphs[0].text, "分页后的封面内容")
        mixed_content_tags = [
            node.tag
            for node in break_then_text.paragraphs[0]._element.iter()
            if node.tag in {qn("w:br"), qn("w:t")}
        ]
        self.assertEqual(mixed_content_tags, [qn("w:br"), qn("w:t")])
        self.assertIn('w:type="page"', break_then_text.paragraphs[-1]._element.xml)

        table_after_break = Document()
        table_after_break.add_paragraph().add_run().add_break(format_paper_module.WD_BREAK.PAGE)
        table_after_break.add_table(rows=1, cols=1).cell(0, 0).text = "封面末表"

        ensure_document_ends_with_page_break(table_after_break)

        terminal_blocks = list(format_paper_module._iter_document_body_blocks(table_after_break))
        self.assertEqual(terminal_blocks[-2].tag, qn("w:tbl"))
        self.assertEqual(terminal_blocks[-1].tag, qn("w:p"))
        self.assertIn('w:type="page"', terminal_blocks[-1].xml)

        paragraph_count = len(table_after_break.paragraphs)
        ensure_document_ends_with_page_break(table_after_break)
        self.assertEqual(len(table_after_break.paragraphs), paragraph_count)

    def test_strip_trailing_blank_paragraphs_stops_at_real_terminal_table(self):
        doc = Document()
        doc.add_paragraph("封面内容")
        non_trailing_blank = doc.add_paragraph()
        doc.add_table(rows=1, cols=1).cell(0, 0).text = "末尾表格"

        self.assertEqual(format_paper_module._strip_trailing_blank_paragraphs(doc), 0)
        self.assertIn(non_trailing_blank._element, doc._element.body)

    def test_concatenate_documents_preserves_body_and_restarts_page_number(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "cover.docx"     # 封面
            second_path = tmp / "body.docx"     # 已排版好的正文
            output_path = tmp / "merged.docx"

            cover = Document()
            cover_run = cover.add_paragraph().add_run("论文封面标题")
            cover_run.font.name = "华文行楷"
            cover.save(str(first_path))

            body = Document()
            body_section = body.sections[0]
            body_section.left_margin = Cm(1.5)
            body.add_paragraph("第一段正文内容")
            body.add_paragraph("第二段正文内容")
            footer_run = body_section.footer.paragraphs[0].add_run()
            page_field = OxmlElement("w:fldSimple")
            page_field.set(qn("w:instr"), "PAGE")
            footer_run._r.addprevious(page_field)
            body.save(str(second_path))

            result = concatenate_documents(
                str(first_path),
                str(second_path),
                str(output_path),
                max_output_bytes=64 * 1024,
            )

            self.assertIsInstance(result, dict)
            self.assertTrue(result["concatenated"])
            self.assertTrue(result["page_number_restarted"])
            self.assertTrue(output_path.exists())
            with ZipFile(output_path) as archive:
                self.assertIsNone(archive.testzip())

            merged = Document(str(output_path))
            merged_texts = [p.text for p in merged.paragraphs]
            # 封面在前、正文在后，两份内容都完整保留
            self.assertIn("论文封面标题", merged_texts)
            self.assertIn("第一段正文内容", merged_texts)
            self.assertIn("第二段正文内容", merged_texts)
            self.assertLess(
                merged_texts.index("论文封面标题"),
                merged_texts.index("第一段正文内容"),
            )
            # 封面原有字体形态未被改写
            cover_idx = merged_texts.index("论文封面标题")
            self.assertEqual(merged.paragraphs[cover_idx].runs[0].font.name, "华文行楷")

            # 合并后分为两节：封面一节、正文一节
            self.assertEqual(len(merged.sections), 2)
            cover_section, body_first_section = merged.sections[0], merged.sections[1]
            # 正文原有页边距与页码字段完整保留
            self.assertAlmostEqual(body_first_section.left_margin.cm, 1.5, places=1)
            self.assertIn("PAGE", body_first_section.footer.paragraphs[0]._p.xml)
            # 正文页码从第 1 页重新编号
            pg_num_type = body_first_section._sectPr.find(qn("w:pgNumType"))
            self.assertIsNotNone(pg_num_type)
            self.assertEqual(pg_num_type.get(qn("w:start")), "1")
            # 封面不计入页码：去掉了页脚引用
            self.assertEqual(len(cover_section._sectPr.findall(qn("w:footerReference"))), 0)

    def test_concatenate_documents_preserves_first_headers_when_restart_is_disabled(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "first-with-header.docx"
            second_path = tmp / "second-with-header.docx"
            output_path = tmp / "combined-with-headers.docx"

            first = Document()
            first.sections[0].header.paragraphs[0].text = "第一份文档页眉"
            first.sections[0].footer.paragraphs[0].text = "第一份文档页脚"
            first.add_paragraph("第一份文档")
            first.save(str(first_path))

            second = Document()
            second.sections[0].header.paragraphs[0].text = "第二份文档页眉"
            second.sections[0].footer.paragraphs[0].text = "第二份文档页脚"
            second.add_paragraph("第二份文档")
            second.save(str(second_path))

            result = concatenate_documents(
                str(first_path),
                str(second_path),
                str(output_path),
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(str(output_path))
            self.assertEqual(len(merged.sections), 2)
            self.assertEqual(
                merged.sections[0].header.paragraphs[0].text,
                "第一份文档页眉",
            )
            self.assertEqual(
                merged.sections[0].footer.paragraphs[0].text,
                "第一份文档页脚",
            )
            self.assertEqual(
                merged.sections[1].header.paragraphs[0].text,
                "第二份文档页眉",
            )
            self.assertEqual(
                merged.sections[1].footer.paragraphs[0].text,
                "第二份文档页脚",
            )

    def test_concatenate_documents_preserves_odd_even_header_semantics(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "first-even-header.docx"
            second_path = tmp / "second-default-header.docx"
            output_path = tmp / "combined-even-header.docx"

            first = Document()
            first.settings.odd_and_even_pages_header_footer = True
            first.sections[0].header.paragraphs[0].text = "第一份默认页眉"
            first.sections[0].even_page_header.paragraphs[0].text = (
                "第一份偶数页页眉"
            )
            first.add_paragraph("第一份文档")
            first.save(first_path)

            second = Document()
            second.sections[0].header.paragraphs[0].text = "第二份默认页眉"
            second.sections[0].footer.paragraphs[0].text = "第二份默认页脚"
            second.add_paragraph("第二份文档")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            self.assertTrue(
                merged.settings.odd_and_even_pages_header_footer
            )
            first_section, second_section = merged.sections
            self.assertEqual(
                first_section.even_page_header.paragraphs[0].text,
                "第一份偶数页页眉",
            )
            # 第二份原未启用 odd/even，启用全局开关后其偶数页
            # 仍应显示自己的 default story，不得继承封面的 even story。
            self.assertEqual(
                second_section.even_page_header.paragraphs[0].text,
                "第二份默认页眉",
            )
            self.assertEqual(
                second_section.even_page_footer.paragraphs[0].text,
                "第二份默认页脚",
            )

    def test_concatenate_documents_preserves_sparse_header_relationship_ids(self):
        tiny_png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "first-sparse-header.docx"
            second_path = tmp / "second.docx"
            output_path = tmp / "combined-sparse-header.docx"

            first = Document()
            first.add_paragraph("第一份文档")
            first.sections[0].header.paragraphs[0].add_run().add_picture(
                BytesIO(tiny_png),
                width=Inches(0.2),
            )
            first.save(first_path)

            with ZipFile(first_path, "r") as source:
                entries = source.infolist()
                payloads = {
                    entry.filename: source.read(entry.filename)
                    for entry in entries
                }
            header_name = next(
                name
                for name in payloads
                if re.fullmatch(r"word/header\d+\.xml", name)
            )
            header_rels_name = (
                f"word/_rels/{Path(header_name).name}.rels"
            )
            relationships = etree.fromstring(payloads[header_rels_name])
            image_relationship = next(
                relationship
                for relationship in relationships
                if relationship.get("Type") == format_paper_module.RT.IMAGE
            )
            old_rid = image_relationship.get("Id")
            image_relationship.set("Id", "assetRel")
            payloads[header_rels_name] = etree.tostring(
                relationships,
                encoding="UTF-8",
                xml_declaration=True,
                standalone=True,
            )
            header_root = etree.fromstring(payloads[header_name])
            for node in header_root.iter():
                for attribute_name, attribute_value in list(node.attrib.items()):
                    if attribute_value == old_rid:
                        node.set(attribute_name, "assetRel")
            payloads[header_name] = etree.tostring(
                header_root,
                encoding="UTF-8",
                xml_declaration=True,
                standalone=True,
            )
            with ZipFile(first_path, "w") as target:
                for entry in entries:
                    target.writestr(entry, payloads[entry.filename])

            second = Document()
            second.add_paragraph("第二份文档")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            with ZipFile(output_path) as archive:
                self.assertIsNone(archive.testzip())
            merged = Document(output_path)
            header_part = merged.sections[0].header.part
            embedded_ids = header_part.element.xpath(".//@r:embed")
            self.assertEqual(embedded_ids, ["assetRel"])
            self.assertIn("assetRel", header_part.rels)
            self.assertEqual(
                header_part.rels["assetRel"].target_part.blob,
                tiny_png,
            )

    def test_concatenate_documents_preserves_nested_bookmark_pairing(self):
        def add_bookmark_paragraph(document, text, bookmark_specs):
            paragraph = document.add_paragraph()
            for name, bookmark_id in bookmark_specs:
                start = OxmlElement("w:bookmarkStart")
                start.set(qn("w:id"), str(bookmark_id))
                start.set(qn("w:name"), name)
                paragraph._element.append(start)
            paragraph.add_run(text)
            for _name, bookmark_id in reversed(bookmark_specs):
                end = OxmlElement("w:bookmarkEnd")
                end.set(qn("w:id"), str(bookmark_id))
                paragraph._element.append(end)

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "nested-bookmarks-first.docx"
            second_path = tmp / "nested-bookmarks-second.docx"
            output_path = tmp / "nested-bookmarks-merged.docx"

            first = Document()
            add_bookmark_paragraph(
                first,
                "嵌套书签",
                (("OuterBookmark", -1), ("InnerBookmark", 2147483648)),
            )
            first.save(first_path)

            second = Document()
            add_bookmark_paragraph(
                second,
                "正文书签",
                (("BodyBookmark", 5),),
            )
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            nested_paragraph = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "嵌套书签"
            )
            semantic_nodes = [
                node
                for node in nested_paragraph._element
                if node.tag in {qn("w:bookmarkStart"), qn("w:bookmarkEnd")}
            ]
            self.assertEqual(
                [
                    (
                        node.tag == qn("w:bookmarkStart"),
                        node.get(qn("w:id")),
                        node.get(qn("w:name")),
                    )
                    for node in semantic_nodes
                ],
                [
                    (True, "0", "OuterBookmark"),
                    (True, "1", "InnerBookmark"),
                    (False, "1", None),
                    (False, "0", None),
                ],
            )

            all_start_ids = [
                node.get(qn("w:id"))
                for node in merged.element.body.findall(
                    ".//" + qn("w:bookmarkStart")
                )
            ]
            all_end_ids = [
                node.get(qn("w:id"))
                for node in merged.element.body.findall(
                    ".//" + qn("w:bookmarkEnd")
                )
            ]
            self.assertEqual(len(all_start_ids), len(set(all_start_ids)))
            self.assertEqual(set(all_start_ids), set(all_end_ids))

    def test_concatenate_documents_renames_colliding_bookmarks_and_references(self):
        def add_bookmark_and_references(document, name, prefix, bookmark_id):
            bookmark_paragraph = document.add_paragraph()
            bookmark_start = OxmlElement("w:bookmarkStart")
            bookmark_start.set(qn("w:id"), str(bookmark_id))
            bookmark_start.set(qn("w:name"), name)
            bookmark_paragraph._element.append(bookmark_start)
            bookmark_paragraph.add_run(f"{prefix} BOOKMARK")
            bookmark_end = OxmlElement("w:bookmarkEnd")
            bookmark_end.set(qn("w:id"), str(bookmark_id))
            bookmark_paragraph._element.append(bookmark_end)

            link_paragraph = document.add_paragraph()
            hyperlink = OxmlElement("w:hyperlink")
            hyperlink.set(qn("w:anchor"), name)
            hyperlink.append(
                parse_xml(
                    f'<w:r {nsdecls("w")}><w:t>{prefix} LINK</w:t></w:r>'
                )
            )
            link_paragraph._element.append(hyperlink)

            simple_paragraph = document.add_paragraph()
            simple_paragraph._element.append(
                parse_xml(
                    f'<w:fldSimple {nsdecls("w")} '
                    f'w:instr=\' REF "{name}" \\h \'> '
                    f"<w:r><w:t>{prefix} REF</w:t></w:r>"
                    "</w:fldSimple>"
                )
            )

            split_paragraph = document.add_paragraph()
            for fragment in (
                f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="begin"/></w:r>',
                f'<w:r {nsdecls("w")}><w:instrText xml:space="preserve"> PAGEREF </w:instrText></w:r>',
                f'<w:r {nsdecls("w")}><w:instrText>"{name}"</w:instrText></w:r>',
                f'<w:r {nsdecls("w")}><w:instrText xml:space="preserve"> \\h </w:instrText></w:r>',
                f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="separate"/></w:r>',
                f'<w:r {nsdecls("w")}><w:t>{prefix} PAGE</w:t></w:r>',
                f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="end"/></w:r>',
            ):
                split_paragraph._element.append(parse_xml(fragment))

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "bookmark-first.docx"
            second_path = tmp / "bookmark-second.docx"
            output_path = tmp / "bookmark-merged.docx"

            first = Document()
            add_bookmark_and_references(
                first,
                "sharedbookmark",
                "FIRST",
                5,
            )
            first.save(first_path)

            second = Document()
            add_bookmark_and_references(
                second,
                "SharedBookmark",
                "SECOND",
                5,
            )
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            bookmark_names = [
                bookmark.get(qn("w:name"))
                for bookmark in merged.element.body.iter(qn("w:bookmarkStart"))
            ]
            self.assertEqual(
                len(bookmark_names),
                len({name.casefold() for name in bookmark_names}),
            )
            self.assertIn("SharedBookmark", bookmark_names)
            renamed_name = next(
                name for name in bookmark_names if name != "SharedBookmark"
            )

            paragraphs_by_text = {
                "".join(
                    node.text or ""
                    for node in paragraph._element.iter(qn("w:t"))
                ): paragraph._element
                for paragraph in merged.paragraphs
            }
            self.assertEqual(
                paragraphs_by_text["FIRST LINK"]
                .find(".//" + qn("w:hyperlink"))
                .get(qn("w:anchor")),
                renamed_name,
            )
            self.assertEqual(
                paragraphs_by_text["SECOND LINK"]
                .find(".//" + qn("w:hyperlink"))
                .get(qn("w:anchor")),
                "SharedBookmark",
            )
            self.assertIn(
                f'REF "{renamed_name}"',
                paragraphs_by_text["FIRST REF"]
                .find(".//" + qn("w:fldSimple"))
                .get(qn("w:instr")),
            )
            first_page_instruction = "".join(
                node.text or ""
                for node in paragraphs_by_text["FIRST PAGE"].iter(
                    qn("w:instrText")
                )
            )
            self.assertIn(f'PAGEREF "{renamed_name}"', first_page_instruction)
            self.assertIn(
                'REF "SharedBookmark"',
                paragraphs_by_text["SECOND REF"]
                .find(".//" + qn("w:fldSimple"))
                .get(qn("w:instr")),
            )

    def test_concatenate_documents_preserves_top_level_alt_chunks(self):
        relationship_type = (
            "http://schemas.openxmlformats.org/officeDocument/2006/"
            "relationships/aFChunk"
        )

        def add_alt_chunk(document, payload):
            part = format_paper_module.Part(
                format_paper_module.PackURI("/word/afchunk1.html"),
                "text/html",
                payload,
                document.part.package,
            )
            relationship_id = document.part.relate_to(part, relationship_type)
            chunk = parse_xml(
                f'<w:altChunk {nsdecls("w", "r")} r:id="{relationship_id}"/>'
            )
            document.element.body.insert(len(document.element.body) - 1, chunk)

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "alt-chunk-first.docx"
            second_path = tmp / "alt-chunk-second.docx"
            output_path = tmp / "alt-chunk-merged.docx"
            first_payload = b"<html><body>FIRST ALT CHUNK</body></html>"
            second_payload = b"<html><body>SECOND ALT CHUNK</body></html>"

            first = Document()
            add_alt_chunk(first, first_payload)
            first.save(first_path)

            second = Document()
            add_alt_chunk(second, second_payload)
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            chunks = merged.element.body.findall(qn("w:altChunk"))
            self.assertEqual(len(chunks), 2)
            relationships = [
                merged.part.rels[chunk.get(qn("r:id"))]
                for chunk in chunks
            ]
            self.assertTrue(
                all(rel.reltype == relationship_type for rel in relationships)
            )
            self.assertEqual(
                {rel.target_part.blob for rel in relationships},
                {first_payload, second_payload},
            )
            self.assertEqual(
                len({rel.target_part.partname for rel in relationships}),
                2,
            )

    def test_concatenate_documents_avoids_colliding_diagram_partnames(self):
        specs = (
            (
                "dm",
                format_paper_module.RT.DIAGRAM_DATA,
                format_paper_module.CT.DML_DIAGRAM_DATA,
                "/word/diagrams/data1.xml",
            ),
            (
                "lo",
                format_paper_module.RT.DIAGRAM_LAYOUT,
                format_paper_module.CT.DML_DIAGRAM_LAYOUT,
                "/word/diagrams/layout1.xml",
            ),
            (
                "qs",
                format_paper_module.RT.DIAGRAM_QUICK_STYLE,
                format_paper_module.CT.DML_DIAGRAM_STYLE,
                "/word/diagrams/quickStyle1.xml",
            ),
            (
                "cs",
                format_paper_module.RT.DIAGRAM_COLORS,
                format_paper_module.CT.DML_DIAGRAM_COLORS,
                "/word/diagrams/colors1.xml",
            ),
        )

        def add_diagram(document, label):
            diagram_namespace = etree.QName(
                qn("dgm:dataModel")
            ).namespace
            root_names = {
                "dm": "dataModel",
                "lo": "layoutDef",
                "qs": "styleDef",
                "cs": "colorsDef",
            }
            relationship_ids = {}
            payloads = {}
            for attribute, relationship_type, content_type, partname in specs:
                payload = (
                    '<?xml version="1.0"?>'
                    f'<dgm:{root_names[attribute]} '
                    f'xmlns:dgm="{diagram_namespace}" '
                    f'xmlns:a="{format_paper_module._THEME_NAMESPACE}" '
                    f'label="{label}-{attribute}">'
                    '<a:solidFill><a:schemeClr val="accent1"/>'
                    '</a:solidFill>'
                    f'</dgm:{root_names[attribute]}>'
                ).encode()
                payloads[attribute] = payload
                relationship_ids[attribute] = document.part.relate_to(
                    format_paper_module.Part(
                        format_paper_module.PackURI(partname),
                        content_type,
                        payload,
                        document.part.package,
                    ),
                    relationship_type,
                )
            paragraph = document.add_paragraph(label)
            paragraph._element.append(
                parse_xml(
                    f'<w:r {nsdecls("w", "dgm", "r")}><w:drawing>'
                    f'<dgm:relIds r:dm="{relationship_ids["dm"]}" '
                    f'r:lo="{relationship_ids["lo"]}" '
                    f'r:qs="{relationship_ids["qs"]}" '
                    f'r:cs="{relationship_ids["cs"]}"/>'
                    "</w:drawing></w:r>"
                )
            )
            return payloads

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "diagram-first.docx"
            second_path = tmp / "diagram-second.docx"
            output_path = tmp / "diagram-merged.docx"

            first = Document()
            first_payloads = add_diagram(first, "FIRST DIAGRAM")
            first.save(first_path)

            second = Document()
            second_payloads = add_diagram(second, "SECOND DIAGRAM")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            with ZipFile(output_path) as archive:
                member_names = archive.namelist()
                self.assertIsNone(archive.testzip())
            self.assertEqual(len(member_names), len(set(member_names)))
            self.assertEqual(
                len(member_names),
                len({name.casefold() for name in member_names}),
            )

            merged = Document(output_path)
            relation_sets = list(
                merged.element.body.iter(qn("dgm:relIds"))
            )
            self.assertEqual(len(relation_sets), 2)
            for attribute, relationship_type, _content_type, _partname in specs:
                relationships = [
                    merged.part.rels[node.get(qn(f"r:{attribute}"))]
                    for node in relation_sets
                ]
                self.assertTrue(
                    all(
                        relationship.reltype == relationship_type
                        for relationship in relationships
                    )
                )
                self.assertEqual(
                    {relationship.target_part.blob for relationship in relationships},
                    {first_payloads[attribute], second_payloads[attribute]},
                )
                self.assertEqual(
                    len(
                        {
                            relationship.target_part.partname
                            for relationship in relationships
                        }
                    ),
                    2,
                )

    def test_concatenate_documents_preserves_multiple_smartart_with_shared_parts(self):
        diagram_namespace = etree.QName(qn("dgm:dataModel")).namespace
        specs = (
            (
                "dm",
                format_paper_module.RT.DIAGRAM_DATA,
                format_paper_module.CT.DML_DIAGRAM_DATA,
                "dataModel",
                "data",
            ),
            (
                "lo",
                format_paper_module.RT.DIAGRAM_LAYOUT,
                format_paper_module.CT.DML_DIAGRAM_LAYOUT,
                "layoutDef",
                "layout",
            ),
            (
                "qs",
                format_paper_module.RT.DIAGRAM_QUICK_STYLE,
                format_paper_module.CT.DML_DIAGRAM_STYLE,
                "styleDef",
                "quickStyle",
            ),
            (
                "cs",
                format_paper_module.RT.DIAGRAM_COLORS,
                format_paper_module.CT.DML_DIAGRAM_COLORS,
                "colorsDef",
                "colors",
            ),
        )

        def add_smartart(document, label, index, shared_parts=None):
            relationship_ids = {}
            created_parts = {}
            for attribute, reltype, content_type, root_name, filename in specs:
                part = None if shared_parts is None else shared_parts.get(attribute)
                if part is None:
                    part_index = index if attribute == "dm" else 1
                    payload = (
                        '<?xml version="1.0"?>'
                        f'<dgm:{root_name} xmlns:dgm="{diagram_namespace}" '
                        f'label="{label}-{attribute}"/>'
                    ).encode("utf-8")
                    part = format_paper_module.Part(
                        format_paper_module.PackURI(
                            f"/word/diagrams/{filename}{part_index}.xml"
                        ),
                        content_type,
                        payload,
                        document.part.package,
                    )
                created_parts[attribute] = part
                relationship_ids[attribute] = document.part.relate_to(part, reltype)

            paragraph = document.add_paragraph(label)
            paragraph._element.append(
                parse_xml(
                    f'<w:r {nsdecls("w", "dgm", "r")}><w:drawing>'
                    f'<dgm:relIds r:dm="{relationship_ids["dm"]}" '
                    f'r:lo="{relationship_ids["lo"]}" '
                    f'r:qs="{relationship_ids["qs"]}" '
                    f'r:cs="{relationship_ids["cs"]}"/>'
                    "</w:drawing></w:r>"
                )
            )
            return created_parts

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "multiple-smartart-first.docx"
            second_path = tmp / "multiple-smartart-second.docx"
            output_path = tmp / "multiple-smartart-merged.docx"

            first = Document()
            first_parts = add_smartart(first, "FIRST SMARTART", 1)
            add_smartart(
                first,
                "SECOND SMARTART",
                2,
                shared_parts={
                    key: value
                    for key, value in first_parts.items()
                    if key != "dm"
                },
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph("BODY")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            with ZipFile(output_path) as archive:
                self.assertIsNone(archive.testzip())
                member_names = archive.namelist()
            self.assertEqual(len(member_names), len(set(member_names)))

            merged = Document(output_path)
            relation_sets = list(merged.element.body.iter(qn("dgm:relIds")))
            self.assertEqual(len(relation_sets), 2)
            data_partnames = {
                merged.part.rels[node.get(qn("r:dm"))].target_part.partname
                for node in relation_sets
            }
            self.assertEqual(len(data_partnames), 2)
            for attribute in ("lo", "qs", "cs"):
                shared_partnames = {
                    merged.part.rels[node.get(qn(f"r:{attribute}"))]
                    .target_part.partname
                    for node in relation_sets
                }
                self.assertEqual(len(shared_partnames), 1)

    def test_concatenate_documents_materializes_diagram_theme_colors_and_rejects_format_drift(self):
        diagram_namespace = etree.QName(qn("dgm:dataModel")).namespace
        specs = (
            (
                "dm",
                format_paper_module.RT.DIAGRAM_DATA,
                format_paper_module.CT.DML_DIAGRAM_DATA,
                "dataModel",
                "/word/diagrams/data1.xml",
            ),
            (
                "lo",
                format_paper_module.RT.DIAGRAM_LAYOUT,
                format_paper_module.CT.DML_DIAGRAM_LAYOUT,
                "layoutDef",
                "/word/diagrams/layout1.xml",
            ),
            (
                "qs",
                format_paper_module.RT.DIAGRAM_QUICK_STYLE,
                format_paper_module.CT.DML_DIAGRAM_STYLE,
                "styleDef",
                "/word/diagrams/quickStyle1.xml",
            ),
            (
                "cs",
                format_paper_module.RT.DIAGRAM_COLORS,
                format_paper_module.CT.DML_DIAGRAM_COLORS,
                "colorsDef",
                "/word/diagrams/colors1.xml",
            ),
        )

        def add_themed_diagram(document, label, *, format_reference):
            relationship_ids = {}
            for attribute, reltype, content_type, root_name, partname in specs:
                themed_content = (
                    '<a:fillRef idx="1"><a:schemeClr val="accent1"/>'
                    "</a:fillRef>"
                    if format_reference and attribute == "qs"
                    else (
                        '<a:solidFill><a:schemeClr val="accent1">'
                        '<a:tint val="20000"/></a:schemeClr></a:solidFill>'
                    )
                )
                payload = (
                    '<?xml version="1.0"?>'
                    f'<dgm:{root_name} xmlns:dgm="{diagram_namespace}" '
                    f'xmlns:a="{format_paper_module._THEME_NAMESPACE}">'
                    f"{themed_content}</dgm:{root_name}>"
                ).encode("utf-8")
                part = format_paper_module.Part(
                    format_paper_module.PackURI(partname),
                    content_type,
                    payload,
                    document.part.package,
                )
                relationship_ids[attribute] = document.part.relate_to(
                    part,
                    reltype,
                )
            paragraph = document.add_paragraph(label)
            paragraph._element.append(
                parse_xml(
                    f'<w:r {nsdecls("w", "dgm", "r")}><w:drawing>'
                    f'<dgm:relIds r:dm="{relationship_ids["dm"]}" '
                    f'r:lo="{relationship_ids["lo"]}" '
                    f'r:qs="{relationship_ids["qs"]}" '
                    f'r:cs="{relationship_ids["cs"]}"/>'
                    "</w:drawing></w:r>"
                )
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "diagram-theme-first.docx"
            second_path = tmp / "diagram-theme-second.docx"
            output_path = tmp / "diagram-theme-merged.docx"

            first = Document()
            self.customize_test_theme(
                first,
                accent1="D01020",
                major_latin="SharedDiagramFont",
                format_name="SharedDiagramFormat",
                background_mapping="accent1",
            )
            add_themed_diagram(
                first,
                "THEMED DIAGRAM",
                format_reference=False,
            )
            first.save(first_path)

            second = Document()
            self.customize_test_theme(
                second,
                accent1="1020D0",
                major_latin="SharedDiagramFont",
                format_name="SharedDiagramFormat",
                background_mapping="accent1",
            )
            second.add_paragraph("BODY")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            colors_part = next(
                relationship.target_part
                for relationship in merged.part.rels.values()
                if relationship.reltype
                == format_paper_module.RT.DIAGRAM_COLORS
            )
            colors_root = etree.fromstring(colors_part.blob)
            self.assertFalse(
                list(
                    colors_root.iter(
                        f"{{{format_paper_module._THEME_NAMESPACE}}}schemeClr"
                    )
                )
            )
            explicit_color = next(
                colors_root.iter(
                    f"{{{format_paper_module._THEME_NAMESPACE}}}srgbClr"
                )
            )
            self.assertEqual(explicit_color.get("val"), "D01020")
            self.assertEqual(
                explicit_color[0].tag,
                f"{{{format_paper_module._THEME_NAMESPACE}}}tint",
            )

            failing_first_path = tmp / "diagram-format-first.docx"
            failing_second_path = tmp / "diagram-format-second.docx"
            failing_output_path = tmp / "diagram-format-merged.docx"
            failing_first = Document()
            self.customize_test_theme(
                failing_first,
                accent1="D01020",
                major_latin="SharedDiagramFont",
                format_name="CoverDiagramFormat",
                background_mapping="accent1",
            )
            add_themed_diagram(
                failing_first,
                "FORMAT-BOUND DIAGRAM",
                format_reference=True,
            )
            failing_first.save(failing_first_path)

            failing_second = Document()
            self.customize_test_theme(
                failing_second,
                accent1="1020D0",
                major_latin="SharedDiagramFont",
                format_name="BodyDiagramFormat",
                background_mapping="accent1",
            )
            failing_second.add_paragraph("BODY")
            failing_second.save(failing_second_path)

            with self.assertRaises(DocumentConcatError):
                concatenate_documents(
                    failing_first_path,
                    failing_second_path,
                    failing_output_path,
                    restart_body_page_number=False,
                )

    def test_concatenate_documents_rejects_invalid_diagram_parts(self):
        diagram_namespace = etree.QName(qn("dgm:dataModel")).namespace
        specs = (
            (
                "dm",
                format_paper_module.RT.DIAGRAM_DATA,
                format_paper_module.CT.DML_DIAGRAM_DATA,
                "dataModel",
                "/word/diagrams/data1.xml",
            ),
            (
                "lo",
                format_paper_module.RT.DIAGRAM_LAYOUT,
                format_paper_module.CT.DML_DIAGRAM_LAYOUT,
                "layoutDef",
                "/word/diagrams/layout1.xml",
            ),
            (
                "qs",
                format_paper_module.RT.DIAGRAM_QUICK_STYLE,
                format_paper_module.CT.DML_DIAGRAM_STYLE,
                "styleDef",
                "/word/diagrams/quickStyle1.xml",
            ),
            (
                "cs",
                format_paper_module.RT.DIAGRAM_COLORS,
                format_paper_module.CT.DML_DIAGRAM_COLORS,
                "colorsDef",
                "/word/diagrams/colors1.xml",
            ),
        )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            for case_name in (
                "missing-data-reference",
                "wrong-root",
                "wrong-content-type",
                "dangling-part-relationship",
            ):
                with self.subTest(case=case_name):
                    first_path = tmp / f"{case_name}-first.docx"
                    second_path = tmp / f"{case_name}-second.docx"
                    output_path = tmp / f"{case_name}-merged.docx"

                    first = Document()
                    relationship_ids = {}
                    for (
                        attribute,
                        relationship_type,
                        content_type,
                        root_name,
                        partname,
                    ) in specs:
                        actual_content_type = (
                            "text/plain"
                            if case_name == "wrong-content-type"
                            and attribute == "dm"
                            else content_type
                        )
                        actual_root = (
                            "layoutDef"
                            if case_name == "wrong-root" and attribute == "dm"
                            else root_name
                        )
                        relationship_reference = (
                            f'<a:blip xmlns:r="{format_paper_module._RELATIONSHIP_NAMESPACE}" '
                            'r:embed="rIdMissing"/>'
                            if case_name == "dangling-part-relationship"
                            and attribute == "dm"
                            else ""
                        )
                        payload = (
                            f'<dgm:{actual_root} xmlns:dgm="{diagram_namespace}" '
                            f'xmlns:a="{format_paper_module._THEME_NAMESPACE}">'
                            f"{relationship_reference}</dgm:{actual_root}>"
                        ).encode("utf-8")
                        part = format_paper_module.Part(
                            format_paper_module.PackURI(partname),
                            actual_content_type,
                            payload,
                            first.part.package,
                        )
                        relationship_ids[attribute] = first.part.relate_to(
                            part,
                            relationship_type,
                        )

                    attributes = [
                        f'r:{attribute}="{relationship_ids[attribute]}"'
                        for attribute, *_rest in specs
                        if not (
                            case_name == "missing-data-reference"
                            and attribute == "dm"
                        )
                    ]
                    paragraph = first.add_paragraph(case_name)
                    paragraph._element.append(
                        parse_xml(
                            f'<w:r {nsdecls("w", "dgm", "r")}><w:drawing>'
                            f'<dgm:relIds {" ".join(attributes)}/>'
                            "</w:drawing></w:r>"
                        )
                    )
                    first.save(first_path)

                    second = Document()
                    second.add_paragraph("BODY")
                    second.save(second_path)

                    with self.assertRaises(DocumentConcatError):
                        concatenate_documents(
                            first_path,
                            second_path,
                            output_path,
                            restart_body_page_number=False,
                        )

    def test_concatenate_documents_preserves_picture_bullet_numbering(self):
        first_image = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )
        second_image = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        )
        vml_namespace = "urn:schemas-microsoft-com:vml"
        relationship_namespace = (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        )

        def add_picture_bullet(
            document,
            label,
            image_payload,
            *,
            use_level_override=False,
        ):
            numbering_part = document.part.numbering_part
            image_part = document.part.package.get_or_add_image_part(
                BytesIO(image_payload)
            )
            relationship_id = numbering_part.relate_to(
                image_part,
                format_paper_module.RT.IMAGE,
            )
            numbering_root = numbering_part.element
            numbering_root.insert(
                0,
                parse_xml(
                    f'<w:numPicBullet {nsdecls("w")} '
                    f'xmlns:v="{vml_namespace}" '
                    f'xmlns:r="{relationship_namespace}" '
                    'w:numPicBulletId="42">'
                    '<w:pict><v:shape>'
                    f'<v:imagedata r:id="{relationship_id}"/>'
                    '</v:shape></w:pict></w:numPicBullet>'
                ),
            )
            picture_reference = (
                "" if use_level_override else '<w:lvlPicBulletId w:val="42"/>'
            )
            abstract_num = parse_xml(
                f'<w:abstractNum {nsdecls("w")} w:abstractNumId="1000">'
                '<w:multiLevelType w:val="singleLevel"/>'
                '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
                '<w:numFmt w:val="bullet"/><w:lvlText w:val=""/>'
                f'{picture_reference}'
                '</w:lvl></w:abstractNum>'
            )
            first_num = numbering_root.find(qn("w:num"))
            abstract_index = (
                numbering_root.index(first_num)
                if first_num is not None
                else len(numbering_root)
            )
            numbering_root.insert(abstract_index, abstract_num)
            level_override = (
                '<w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0">'
                '<w:start w:val="1"/><w:numFmt w:val="bullet"/>'
                '<w:lvlText w:val=""/><w:lvlPicBulletId w:val="42"/>'
                '</w:lvl></w:lvlOverride>'
                if use_level_override
                else ""
            )
            numbering_root.append(
                parse_xml(
                    f'<w:num {nsdecls("w")} w:numId="1000">'
                    '<w:abstractNumId w:val="1000"/>'
                    f'{level_override}</w:num>'
                )
            )

            paragraph = document.add_paragraph(label)
            paragraph._element.get_or_add_pPr().append(
                parse_xml(
                    f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/>'
                    '<w:numId w:val="1000"/></w:numPr>'
                )
            )

        def resolve_picture_bullet(document, paragraph):
            numbering_part = document.part.numbering_part
            numbering_root = numbering_part.element
            num_id = paragraph._element.pPr.numPr.numId.val
            num = next(
                item
                for item in numbering_root.findall(qn("w:num"))
                if int(item.get(qn("w:numId"))) == num_id
            )
            abstract_id = num.find(qn("w:abstractNumId")).get(qn("w:val"))
            abstract_num = next(
                item
                for item in numbering_root.findall(qn("w:abstractNum"))
                if item.get(qn("w:abstractNumId")) == abstract_id
            )
            picture_reference = num.find(".//" + qn("w:lvlPicBulletId"))
            if picture_reference is None:
                picture_reference = abstract_num.find(
                    ".//" + qn("w:lvlPicBulletId")
                )
            picture_bullet_id = picture_reference.get(qn("w:val"))
            definition = next(
                item
                for item in numbering_root.findall(qn("w:numPicBullet"))
                if item.get(qn("w:numPicBulletId")) == picture_bullet_id
            )
            image_data = next(
                definition.iter(f"{{{vml_namespace}}}imagedata")
            )
            relationship = numbering_part.rels[
                image_data.get(qn("r:id"))
            ]
            return picture_bullet_id, relationship

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "picture-bullet-first.docx"
            second_path = tmp / "picture-bullet-second.docx"
            output_path = tmp / "picture-bullet-merged.docx"

            first = Document()
            add_picture_bullet(
                first,
                "FIRST PICTURE BULLET",
                first_image,
                use_level_override=True,
            )
            first.save(first_path)

            second = Document()
            add_picture_bullet(second, "SECOND PICTURE BULLET", second_image)
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            resolved = {
                paragraph.text: resolve_picture_bullet(merged, paragraph)
                for paragraph in merged.paragraphs
                if paragraph.text.endswith("PICTURE BULLET")
            }
            self.assertEqual(
                set(resolved),
                {"FIRST PICTURE BULLET", "SECOND PICTURE BULLET"},
            )
            self.assertNotEqual(resolved["FIRST PICTURE BULLET"][0], "42")
            self.assertEqual(resolved["SECOND PICTURE BULLET"][0], "42")
            self.assertEqual(
                resolved["FIRST PICTURE BULLET"][1].target_part.blob,
                first_image,
            )
            self.assertEqual(
                resolved["SECOND PICTURE BULLET"][1].target_part.blob,
                second_image,
            )
            self.assertTrue(
                all(
                    relationship.reltype == format_paper_module.RT.IMAGE
                    for _picture_id, relationship in resolved.values()
                )
            )
            with ZipFile(output_path) as archive:
                member_names = archive.namelist()
                self.assertIsNone(archive.testzip())
            self.assertEqual(
                len(member_names),
                len({name.casefold() for name in member_names}),
            )

    def test_concatenate_documents_preserves_embedded_font_tables(self):
        font_content_type = (
            "application/vnd.openxmlformats-officedocument.obfuscatedFont"
        )
        first_font_payload = b"FIRST EMBEDDED ODTTF PAYLOAD"
        second_font_payload = b"SECOND EMBEDDED ODTTF PAYLOAD"
        shared_bold_payload = b"SHARED BOLD ODTTF PAYLOAD"
        shared_regular_payload = b"SHARED REGULAR ODTTF PAYLOAD"

        def add_embedded_font(
            document,
            font_name,
            font_key,
            payload,
            *,
            variant="w:embedRegular",
            partname="/word/fonts/font1.odttf",
        ):
            font_table_part = document.part.part_related_by(
                format_paper_module.RT.FONT_TABLE
            )
            font_part = format_paper_module.Part(
                format_paper_module.PackURI(partname),
                font_content_type,
                payload,
                document.part.package,
            )
            relationship_id = font_table_part.relate_to(
                font_part,
                format_paper_module.RT.FONT,
            )
            fonts_root = etree.fromstring(font_table_part.blob)
            font = etree.SubElement(fonts_root, qn("w:font"))
            font.set(qn("w:name"), font_name)
            embedded = etree.SubElement(font, qn(variant))
            embedded.set(qn("r:id"), relationship_id)
            embedded.set(qn("w:fontKey"), font_key)
            font_table_part._blob = etree.tostring(
                fonts_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

            run = document.add_paragraph(font_name).add_run(" sample")
            run.font.name = font_name

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "embedded-font-first.docx"
            second_path = tmp / "embedded-font-second.docx"
            output_path = tmp / "embedded-font-merged.docx"

            first = Document()
            add_embedded_font(
                first,
                "CoverEmbeddedFont",
                "{00112233-4455-6677-8899-AABBCCDDEEFF}",
                first_font_payload,
            )
            add_embedded_font(
                first,
                "SharedEmbeddedFont",
                "{01234567-89AB-CDEF-0123-456789ABCDEF}",
                shared_bold_payload,
                variant="w:embedBold",
                partname="/word/fonts/font2.odttf",
            )
            first.part.part_related_by(
                format_paper_module.RT.FONT_TABLE
            )._content_type = format_paper_module.CT.WML_FONT_TABLE.upper()
            first.save(first_path)

            second = Document()
            add_embedded_font(
                second,
                "BodyEmbeddedFont",
                "{FFEEDDCC-BBAA-9988-7766-554433221100}",
                second_font_payload,
            )
            add_embedded_font(
                second,
                "SharedEmbeddedFont",
                "{FEDCBA98-7654-3210-FEDC-BA9876543210}",
                shared_regular_payload,
                partname="/word/fonts/font2.odttf",
            )
            second.part.part_related_by(
                format_paper_module.RT.FONT_TABLE
            )._content_type = format_paper_module.CT.WML_FONT_TABLE.upper()
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            font_table_part = merged.part.part_related_by(
                format_paper_module.RT.FONT_TABLE
            )
            fonts_root = etree.fromstring(font_table_part.blob)
            resolved_fonts = {}
            for font in fonts_root.findall(qn("w:font")):
                name = font.get(qn("w:name"))
                if name not in {"CoverEmbeddedFont", "BodyEmbeddedFont"}:
                    continue
                embedded = font.find(qn("w:embedRegular"))
                relationship = font_table_part.rels[
                    embedded.get(qn("r:id"))
                ]
                resolved_fonts[name] = relationship

            self.assertEqual(
                set(resolved_fonts),
                {"CoverEmbeddedFont", "BodyEmbeddedFont"},
            )
            self.assertEqual(
                resolved_fonts["CoverEmbeddedFont"].target_part.blob,
                first_font_payload,
            )
            self.assertEqual(
                resolved_fonts["BodyEmbeddedFont"].target_part.blob,
                second_font_payload,
            )
            self.assertTrue(
                all(
                    relationship.reltype == format_paper_module.RT.FONT
                    and not relationship.is_external
                    for relationship in resolved_fonts.values()
                )
            )
            self.assertEqual(
                len(
                    {
                        relationship.target_part.partname
                        for relationship in resolved_fonts.values()
                    }
                ),
                2,
            )

            shared_font = next(
                font
                for font in fonts_root.findall(qn("w:font"))
                if font.get(qn("w:name")) == "SharedEmbeddedFont"
            )
            shared_variants = {}
            for variant in ("w:embedRegular", "w:embedBold"):
                embedded = shared_font.find(qn(variant))
                relationship = font_table_part.rels[
                    embedded.get(qn("r:id"))
                ]
                shared_variants[variant] = relationship.target_part.blob
            self.assertEqual(
                shared_variants,
                {
                    "w:embedRegular": shared_regular_payload,
                    "w:embedBold": shared_bold_payload,
                },
            )
            with ZipFile(output_path) as archive:
                member_names = archive.namelist()
                self.assertIsNone(archive.testzip())
            self.assertEqual(
                len(member_names),
                len({name.casefold() for name in member_names}),
            )

    def test_concatenate_documents_preserves_colliding_paragraph_and_character_styles(self):
        def add_formatted_style(document, name, style_type, size, color):
            style = document.styles.add_style(name, style_type)
            style.element.append(
                parse_xml(
                    f'<w:rPr {nsdecls("w")}>'
                    f'<w:sz w:val="{size}"/>'
                    f'<w:color w:val="{color}"/>'
                    '</w:rPr>'
                )
            )
            return style

        def style_format(style):
            run_properties = style.element.find(qn("w:rPr"))
            return (
                run_properties.find(qn("w:sz")).get(qn("w:val")),
                run_properties.find(qn("w:color")).get(qn("w:val")),
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "style-collision-first.docx"
            second_path = tmp / "style-collision-second.docx"
            output_path = tmp / "style-collision-merged.docx"

            first = Document()
            cover_paragraph_style = add_formatted_style(
                first,
                "SharedParagraphCollision",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
                "18",
                "E01020",
            )
            cover_character_style = add_formatted_style(
                first,
                "SharedCharacterCollision",
                format_paper_module.WD_STYLE_TYPE.CHARACTER,
                "20",
                "20A040",
            )
            cover = first.add_paragraph(style=cover_paragraph_style)
            cover.add_run("COVER STYLE COLLISION", style=cover_character_style)
            first.sections[0].header.paragraphs[0].text = (
                "COVER STYLE COLLISION HEADER"
            )
            first.sections[0].header.paragraphs[0].style = (
                cover_paragraph_style
            )
            first.save(first_path)

            second = Document()
            body_paragraph_style = add_formatted_style(
                second,
                "SharedParagraphCollision",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
                "48",
                "1020E0",
            )
            body_character_style = add_formatted_style(
                second,
                "SharedCharacterCollision",
                format_paper_module.WD_STYLE_TYPE.CHARACTER,
                "44",
                "A02080",
            )
            body = second.add_paragraph(style=body_paragraph_style)
            body.add_run("BODY STYLE COLLISION", style=body_character_style)
            second.sections[0].header.paragraphs[0].text = (
                "BODY STYLE COLLISION HEADER"
            )
            second.sections[0].header.paragraphs[0].style = (
                body_paragraph_style
            )
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "COVER STYLE COLLISION"
            )
            body = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "BODY STYLE COLLISION"
            )
            self.assertNotEqual(cover.style.style_id, body.style.style_id)
            self.assertNotEqual(
                cover.runs[0].style.style_id,
                body.runs[0].style.style_id,
            )
            self.assertNotEqual(
                cover.style.style_id,
                cover.runs[0].style.style_id,
            )
            self.assertEqual(
                cover.style.type,
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            self.assertEqual(
                cover.runs[0].style.type,
                format_paper_module.WD_STYLE_TYPE.CHARACTER,
            )
            self.assertEqual(style_format(cover.style), ("18", "E01020"))
            self.assertEqual(
                style_format(cover.runs[0].style),
                ("20", "20A040"),
            )
            self.assertEqual(style_format(body.style), ("48", "1020E0"))
            self.assertEqual(
                style_format(body.runs[0].style),
                ("44", "A02080"),
            )
            cover_header_part = next(
                relationship.target_part
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.HEADER
                and "COVER STYLE COLLISION HEADER"
                in "".join(relationship.target_part.element.itertext())
            )
            header_style_id = next(
                cover_header_part.element.iter(qn("w:pStyle"))
            ).get(qn("w:val"))
            self.assertEqual(header_style_id, cover.style.style_id)

    def test_concatenate_documents_compares_transitive_style_dependencies(self):
        def add_style_chain(document, size, color):
            base = document.styles.add_style(
                "BaseCollision",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            base.element.append(
                parse_xml(
                    f'<w:rPr {nsdecls("w")}>'
                    f'<w:sz w:val="{size}"/>'
                    f'<w:color w:val="{color}"/>'
                    '</w:rPr>'
                )
            )
            derived = document.styles.add_style(
                "DerivedCollision",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            derived.base_style = base
            return derived

        def base_format(style):
            run_properties = style.base_style.element.find(qn("w:rPr"))
            return (
                run_properties.find(qn("w:sz")).get(qn("w:val")),
                run_properties.find(qn("w:color")).get(qn("w:val")),
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "derived-style-first.docx"
            second_path = tmp / "derived-style-second.docx"
            output_path = tmp / "derived-style-merged.docx"

            first = Document()
            first.add_paragraph(
                "COVER DERIVED STYLE",
                style=add_style_chain(first, "18", "E01020"),
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph(
                "BODY DERIVED STYLE",
                style=add_style_chain(second, "48", "1020E0"),
            )
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "COVER DERIVED STYLE"
            )
            body = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "BODY DERIVED STYLE"
            )
            self.assertNotEqual(cover.style.style_id, body.style.style_id)
            self.assertNotEqual(
                cover.style.base_style.style_id,
                body.style.base_style.style_id,
            )
            self.assertEqual(base_format(cover.style), ("18", "E01020"))
            self.assertEqual(base_format(body.style), ("48", "1020E0"))

    def test_concatenate_documents_preserves_defaults_without_overriding_direct_formatting(self):
        def format_normal(document, font_name, size, color):
            normal = document.styles["Normal"]
            normal.font.name = font_name
            normal.font.size = format_paper_module.Pt(size)
            normal.font.color.rgb = format_paper_module.RGBColor.from_string(
                color
            )
            return normal

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "default-style-first.docx"
            second_path = tmp / "default-style-second.docx"
            output_path = tmp / "default-style-merged.docx"

            first = Document()
            cover_normal = format_normal(
                first,
                "SourceDefaultFont",
                12,
                "20A040",
            )
            cover_normal_paragraph = first.add_paragraph(
                "COVER EXPLICIT NORMAL",
                style=cover_normal,
            )
            cover_normal_paragraph._element.get_or_add_pPr().append(
                parse_xml(
                    f'<w:pStyle {nsdecls("w")} w:val="Normal"/>'
                )
            )
            direct_paragraph = first.add_paragraph("COVER DIRECT ")
            direct_run = direct_paragraph.add_run("FORMATTING")
            direct_run.font.name = "DirectFont"
            direct_run.font.size = format_paper_module.Pt(9)
            direct_run.font.color.rgb = (
                format_paper_module.RGBColor.from_string("E01020")
            )
            first.save(first_path)

            second = Document()
            body_normal = format_normal(
                second,
                "TargetDefaultFont",
                24,
                "1020E0",
            )
            body_normal_paragraph = second.add_paragraph(
                "BODY EXPLICIT NORMAL",
                style=body_normal,
            )
            body_normal_paragraph._element.get_or_add_pPr().append(
                parse_xml(
                    f'<w:pStyle {nsdecls("w")} w:val="Normal"/>'
                )
            )
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover_normal_paragraph = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "COVER EXPLICIT NORMAL"
            )
            body_normal_paragraph = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "BODY EXPLICIT NORMAL"
            )
            self.assertNotEqual(
                cover_normal_paragraph.style.style_id,
                body_normal_paragraph.style.style_id,
            )
            self.assertIsNone(
                cover_normal_paragraph.style.element.get(qn("w:default"))
            )
            cover_style_rpr = cover_normal_paragraph.style.element.find(
                qn("w:rPr")
            )
            self.assertEqual(
                cover_style_rpr.find(qn("w:rFonts")).get(qn("w:ascii")),
                "SourceDefaultFont",
            )
            self.assertEqual(
                cover_style_rpr.find(qn("w:sz")).get(qn("w:val")),
                "24",
            )

            paragraph_defaults = [
                style
                for style in merged.styles.element.findall(qn("w:style"))
                if style.get(qn("w:type")) == "paragraph"
                and style.get(qn("w:default")) in {"1", "true", "on"}
            ]
            self.assertEqual(len(paragraph_defaults), 1)
            self.assertEqual(
                paragraph_defaults[0].get(qn("w:styleId")),
                body_normal_paragraph.style.style_id,
            )

            direct_paragraph = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "COVER DIRECT FORMATTING"
            )
            direct_run = next(
                run for run in direct_paragraph.runs if run.text == "FORMATTING"
            )
            run_properties = direct_run._element.rPr
            self.assertEqual(len(run_properties.findall(qn("w:rFonts"))), 1)
            self.assertEqual(len(run_properties.findall(qn("w:sz"))), 1)
            self.assertEqual(len(run_properties.findall(qn("w:color"))), 1)
            self.assertEqual(
                run_properties.find(qn("w:rFonts")).get(qn("w:ascii")),
                "DirectFont",
            )
            self.assertEqual(
                run_properties.find(qn("w:sz")).get(qn("w:val")),
                "18",
            )
            self.assertEqual(
                run_properties.find(qn("w:color")).get(qn("w:val")),
                "E01020",
            )

    def test_concatenate_documents_materializes_default_slots_with_character_and_textbox_styles(self):
        def format_normal(document, font_name, east_asia, size):
            normal = document.styles["Normal"]
            normal.font.name = font_name
            normal.font.size = format_paper_module.Pt(size)
            normal.element.find(".//" + qn("w:rFonts")).set(
                qn("w:eastAsia"),
                east_asia,
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "default-slots-first.docx"
            second_path = tmp / "default-slots-second.docx"
            output_path = tmp / "default-slots-merged.docx"

            first = Document()
            format_normal(first, "SourceDefaultFont", "SourceEastAsia", 12)
            character_style = first.styles.add_style(
                "DefaultAwareCharacter",
                format_paper_module.WD_STYLE_TYPE.CHARACTER,
            )
            character_style.font.italic = True
            character_paragraph = first.add_paragraph()
            character_paragraph.add_run(
                "CHARACTER DEFAULTS",
                style=character_style,
            )

            partial_run = first.add_paragraph().add_run("PARTIAL RFONTS")
            partial_run._element.get_or_add_rPr().append(
                parse_xml(
                    f'<w:rFonts {nsdecls("w")} w:eastAsia="DirectEastAsia"/>'
                )
            )

            inner_style = first.styles.add_style(
                "InnerTextboxStyle",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            inner_style.font.size = format_paper_module.Pt(8)
            textbox_host = first.add_paragraph()
            textbox_host._element.append(
                parse_xml(
                    f'<w:r {nsdecls("w", "v")}><w:pict><v:shape>'
                    '<v:textbox><w:txbxContent><w:p>'
                    '<w:pPr><w:pStyle w:val="InnerTextboxStyle"/></w:pPr>'
                    '<w:r><w:t>INNER TEXTBOX STYLE</w:t></w:r>'
                    '</w:p></w:txbxContent></v:textbox>'
                    '</v:shape></w:pict></w:r>'
                )
            )
            first.save(first_path)

            second = Document()
            format_normal(second, "TargetDefaultFont", "TargetEastAsia", 24)
            second.add_paragraph("BODY")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            character_paragraph = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "CHARACTER DEFAULTS"
            )
            character_rpr = character_paragraph.runs[0]._element.rPr
            self.assertIsNotNone(character_rpr.find(qn("w:rStyle")))
            self.assertIsNone(character_rpr.find(qn("w:rFonts")))
            character_paragraph_style_rpr = (
                character_paragraph.style.element.find(qn("w:rPr"))
            )
            self.assertEqual(
                character_paragraph_style_rpr.find(qn("w:rFonts")).get(
                    qn("w:ascii")
                ),
                "SourceDefaultFont",
            )
            self.assertEqual(
                character_paragraph_style_rpr.find(qn("w:sz")).get(
                    qn("w:val")
                ),
                "24",
            )

            partial_paragraph = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "PARTIAL RFONTS"
            )
            partial_rfonts = partial_paragraph.runs[0]._element.rPr.find(
                qn("w:rFonts")
            )
            self.assertEqual(
                partial_rfonts.get(qn("w:eastAsia")),
                "DirectEastAsia",
            )
            self.assertIsNone(partial_rfonts.get(qn("w:ascii")))
            partial_style_rfonts = partial_paragraph.style.element.find(
                ".//" + qn("w:rFonts")
            )
            self.assertEqual(
                partial_style_rfonts.get(qn("w:ascii")),
                "SourceDefaultFont",
            )
            self.assertEqual(
                partial_style_rfonts.get(qn("w:hAnsi")),
                "SourceDefaultFont",
            )

            textbox_content = next(
                merged.element.body.iter(qn("w:txbxContent"))
            )
            inner_run_properties = next(
                textbox_content.iter(qn("w:rPr"))
            )
            self.assertIsNone(inner_run_properties.find(qn("w:sz")))
            inner_style_id = next(
                textbox_content.iter(qn("w:pStyle"))
            ).get(qn("w:val"))
            copied_inner_style = merged.styles.element.get_by_id(inner_style_id)
            self.assertEqual(
                copied_inner_style.find(".//" + qn("w:sz")).get(qn("w:val")),
                "16",
            )

    def test_concatenate_documents_isolates_empty_paragraph_and_table_defaults(self):
        def set_table_default_fill(document, fill):
            table_default = document.styles.default(
                format_paper_module.WD_STYLE_TYPE.TABLE
            )
            table_properties = table_default.element.find(qn("w:tblPr"))
            if table_properties is None:
                table_properties = parse_xml(
                    f'<w:tblPr {nsdecls("w")}/>'
                )
                table_default.element.append(table_properties)
            table_properties.append(
                parse_xml(
                    f'<w:shd {nsdecls("w")} w:val="clear" '
                    f'w:color="auto" w:fill="{fill}"/>'
                )
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "empty-default-first.docx"
            second_path = tmp / "empty-default-second.docx"
            output_path = tmp / "empty-default-merged.docx"

            first = Document()
            source_normal = first.styles["Normal"].element
            for tag_name in ("w:pPr", "w:rPr"):
                node = source_normal.find(qn(tag_name))
                if node is not None:
                    source_normal.remove(node)
            set_table_default_fill(first, "E01020")
            first.add_paragraph("COVER EMPTY DEFAULT")
            cover_table = first.add_table(rows=1, cols=1)
            cover_table.cell(0, 0).text = "COVER TABLE DEFAULT"
            first.save(first_path)

            second = Document()
            second.styles["Normal"].font.size = format_paper_module.Pt(24)
            second.styles["Normal"].font.bold = True
            set_table_default_fill(second, "1020E0")
            second.add_paragraph("BODY DEFAULT")
            body_table = second.add_table(rows=1, cols=1)
            body_table.cell(0, 0).text = "BODY TABLE DEFAULT"
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "COVER EMPTY DEFAULT"
            )
            cover_style_reference = cover._element.pPr.pStyle.val
            self.assertNotEqual(cover_style_reference, "Normal")
            copied_normal = merged.styles.element.get_by_id(
                cover_style_reference
            )
            self.assertIsNone(copied_normal.get(qn("w:default")))
            self.assertIsNone(copied_normal.find(qn("w:rPr")))

            cover_table = next(
                table
                for table in merged.tables
                if table.cell(0, 0).text == "COVER TABLE DEFAULT"
            )
            cover_table_style_id = cover_table._tbl.tblPr.tblStyle.val
            copied_table_default = merged.styles.element.get_by_id(
                cover_table_style_id
            )
            self.assertIsNone(copied_table_default.get(qn("w:default")))
            self.assertEqual(
                copied_table_default.find(".//" + qn("w:shd")).get(
                    qn("w:fill")
                ),
                "E01020",
            )

    def test_concatenate_documents_isolates_document_defaults(self):
        def set_document_defaults(
            document,
            *,
            size,
            bold,
            alignment,
            keep_next,
        ):
            defaults = document.styles.element.find(qn("w:docDefaults"))
            run_properties = defaults.find(qn("w:rPrDefault")).find(
                qn("w:rPr")
            )
            for tag_name in ("w:sz", "w:b"):
                existing = run_properties.find(qn(tag_name))
                if existing is not None:
                    run_properties.remove(existing)
            run_properties.append(
                parse_xml(
                    f'<w:sz {nsdecls("w")} w:val="{size}"/>'
                )
            )
            if bold:
                run_properties.append(
                    parse_xml(f'<w:b {nsdecls("w")}/>')
                )

            paragraph_properties = defaults.find(
                qn("w:pPrDefault")
            ).find(qn("w:pPr"))
            for tag_name in ("w:jc", "w:keepNext"):
                existing = paragraph_properties.find(qn(tag_name))
                if existing is not None:
                    paragraph_properties.remove(existing)
            paragraph_properties.append(
                parse_xml(
                    f'<w:jc {nsdecls("w")} w:val="{alignment}"/>'
                )
            )
            if keep_next:
                paragraph_properties.append(
                    parse_xml(f'<w:keepNext {nsdecls("w")}/>')
                )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "doc-default-first.docx"
            second_path = tmp / "doc-default-second.docx"
            output_path = tmp / "doc-default-merged.docx"

            first = Document()
            set_document_defaults(
                first,
                size="20",
                bold=False,
                alignment="center",
                keep_next=False,
            )
            first.add_paragraph("COVER DOCUMENT DEFAULT")
            derived = first.styles.add_style(
                "DocumentDefaultDerived",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            derived.base_style = first.styles["Normal"]
            derived_paragraph = first.add_paragraph(
                "COVER DERIVED DOCUMENT DEFAULT",
                style=derived,
            )
            derived_paragraph.runs[0].bold = True
            first.add_table(rows=1, cols=1).cell(0, 0).text = (
                "COVER PLAIN DEFAULT TABLE"
            )
            first.save(first_path)

            second = Document()
            set_document_defaults(
                second,
                size="48",
                bold=True,
                alignment="right",
                keep_next=True,
            )
            target_derived = second.styles.add_style(
                "DocumentDefaultDerived",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            target_derived.base_style = second.styles["Normal"]
            second.add_paragraph("BODY DOCUMENT DEFAULT")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "COVER DOCUMENT DEFAULT"
            )
            copied_normal_id = cover._element.pPr.pStyle.val
            self.assertNotEqual(copied_normal_id, "Normal")
            copied_normal = merged.styles.element.get_by_id(
                copied_normal_id
            )
            copied_run_defaults = copied_normal.find(qn("w:rPr"))
            self.assertEqual(
                copied_run_defaults.find(qn("w:sz")).get(qn("w:val")),
                "20",
            )
            self.assertEqual(
                copied_run_defaults.find(qn("w:b")).get(qn("w:val")),
                "0",
            )
            copied_paragraph_defaults = copied_normal.find(qn("w:pPr"))
            self.assertEqual(
                copied_paragraph_defaults.find(qn("w:jc")).get(
                    qn("w:val")
                ),
                "center",
            )
            self.assertEqual(
                copied_paragraph_defaults.find(qn("w:keepNext")).get(
                    qn("w:val")
                ),
                "0",
            )

            copied_derived_paragraph = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "COVER DERIVED DOCUMENT DEFAULT"
            )
            self.assertEqual(
                copied_derived_paragraph.style.base_style.style_id,
                copied_normal_id,
            )
            self.assertTrue(copied_derived_paragraph.runs[0].bold)
            copied_table_paragraph = next(
                table.cell(0, 0).paragraphs[0]
                for table in merged.tables
                if table.cell(0, 0).text == "COVER PLAIN DEFAULT TABLE"
            )
            self.assertEqual(
                copied_table_paragraph.style.style_id,
                copied_normal_id,
            )

            body = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "BODY DOCUMENT DEFAULT"
            )
            self.assertEqual(body.style.style_id, "Normal")
            output_run_defaults = (
                merged.styles.element.find(qn("w:docDefaults"))
                .find(qn("w:rPrDefault"))
                .find(qn("w:rPr"))
            )
            self.assertEqual(
                output_run_defaults.find(qn("w:sz")).get(qn("w:val")),
                "48",
            )
            self.assertIsNotNone(output_run_defaults.find(qn("w:b")))

    def test_concatenate_documents_keeps_direct_indent_and_spacing_groups_authoritative(self):
        def set_paragraph_defaults(document, start, first_line, before, after):
            paragraph_properties = (
                document.styles.element.find(qn("w:docDefaults"))
                .find(qn("w:pPrDefault"))
                .find(qn("w:pPr"))
            )
            for tag_name in ("w:ind", "w:spacing"):
                existing = paragraph_properties.find(qn(tag_name))
                if existing is not None:
                    paragraph_properties.remove(existing)
            paragraph_properties.extend(
                (
                    parse_xml(
                        f'<w:ind {nsdecls("w")} w:start="{start}" '
                        f'w:firstLine="{first_line}"/>'
                    ),
                    parse_xml(
                        f'<w:spacing {nsdecls("w")} w:before="{before}" '
                        f'w:after="{after}"/>'
                    ),
                )
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "grouped-default-first.docx"
            second_path = tmp / "grouped-default-second.docx"
            output_path = tmp / "grouped-default-merged.docx"

            first = Document()
            set_paragraph_defaults(first, "720", "360", "100", "200")
            normal_properties = first.styles["Normal"].element.get_or_add_pPr()
            normal_properties.append(
                parse_xml(
                    f'<w:ind {nsdecls("w")} w:left="120" w:hanging="240"/>'
                )
            )
            normal_properties.append(
                parse_xml(
                    f'<w:spacing {nsdecls("w")} '
                    'w:beforeLines="150" w:afterLines="250"/>'
                )
            )
            first.add_paragraph("COVER GROUPED DEFAULTS")
            first.save(first_path)

            second = Document()
            set_paragraph_defaults(second, "1440", "720", "300", "400")
            second.add_paragraph("BODY")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "COVER GROUPED DEFAULTS"
            )
            copied_style = merged.styles.element.get_by_id(
                cover._element.pPr.pStyle.val
            )
            indentation = copied_style.find(".//" + qn("w:ind"))
            self.assertEqual(indentation.get(qn("w:left")), "120")
            self.assertEqual(indentation.get(qn("w:hanging")), "240")
            self.assertIsNone(indentation.get(qn("w:start")))
            self.assertIsNone(indentation.get(qn("w:firstLine")))
            spacing = copied_style.find(".//" + qn("w:spacing"))
            self.assertEqual(spacing.get(qn("w:beforeLines")), "150")
            self.assertEqual(spacing.get(qn("w:afterLines")), "250")
            self.assertIsNone(spacing.get(qn("w:before")))
            self.assertIsNone(spacing.get(qn("w:after")))

    def test_concatenate_documents_rejects_unrepresentable_document_default_gap(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "missing-doc-default-first.docx"
            second_path = tmp / "missing-doc-default-second.docx"
            output_path = tmp / "missing-doc-default-merged.docx"

            first = Document()
            source_run_defaults = (
                first.styles.element.find(qn("w:docDefaults"))
                .find(qn("w:rPrDefault"))
                .find(qn("w:rPr"))
            )
            source_fonts = source_run_defaults.find(qn("w:rFonts"))
            source_run_defaults.remove(source_fonts)
            first.add_paragraph("COVER APPLICATION FONT DEFAULT")
            first.save(first_path)

            second = Document()
            second.add_paragraph("BODY")
            second.save(second_path)

            with self.assertRaises(DocumentConcatError):
                concatenate_documents(
                    first_path,
                    second_path,
                    output_path,
                    restart_body_page_number=False,
                )

    def test_concatenate_documents_rejects_partial_composite_default_gap(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "partial-font-default-first.docx"
            second_path = tmp / "partial-font-default-second.docx"
            output_path = tmp / "partial-font-default-merged.docx"

            first = Document()
            source_fonts = (
                first.styles.element.find(qn("w:docDefaults"))
                .find(qn("w:rPrDefault"))
                .find(qn("w:rPr"))
                .find(qn("w:rFonts"))
            )
            source_fonts.attrib.pop(qn("w:eastAsiaTheme"))
            first.add_paragraph("COVER PARTIAL FONT DEFAULT")
            first.save(first_path)

            second = Document()
            second.add_paragraph("BODY")
            second.save(second_path)

            with self.assertRaises(DocumentConcatError):
                concatenate_documents(
                    first_path,
                    second_path,
                    output_path,
                    restart_body_page_number=False,
                )

    def test_concatenate_documents_rejects_table_cascade_with_different_document_defaults(self):
        def set_default_size(document, value):
            run_properties = (
                document.styles.element.find(qn("w:docDefaults"))
                .find(qn("w:rPrDefault"))
                .find(qn("w:rPr"))
            )
            size = run_properties.find(qn("w:sz"))
            if size is None:
                size = parse_xml(f'<w:sz {nsdecls("w")}/>')
                run_properties.append(size)
            size.set(qn("w:val"), value)

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "table-default-first.docx"
            second_path = tmp / "table-default-second.docx"
            output_path = tmp / "table-default-merged.docx"

            first = Document()
            set_default_size(first, "20")
            conflicting_table_style = first.styles.add_style(
                "ConflictingDefaultTableStyle",
                format_paper_module.WD_STYLE_TYPE.TABLE,
            )
            conflicting_table_style.element.append(
                parse_xml(
                    f'<w:rPr {nsdecls("w")}><w:sz w:val="30"/></w:rPr>'
                )
            )
            cover_table = first.add_table(rows=1, cols=1)
            cover_table.style = conflicting_table_style
            cover_table.cell(0, 0).text = "COVER TABLE"
            first.save(first_path)

            second = Document()
            set_default_size(second, "48")
            second.add_paragraph("BODY")
            second.save(second_path)

            with self.assertRaises(DocumentConcatError):
                concatenate_documents(
                    first_path,
                    second_path,
                    output_path,
                    restart_body_page_number=False,
                )

    def test_concatenate_documents_materializes_only_different_document_defaults(self):
        def set_run_default(document, tag_name, value):
            run_properties = (
                document.styles.element.find(qn("w:docDefaults"))
                .find(qn("w:rPrDefault"))
                .find(qn("w:rPr"))
            )
            property_element = run_properties.find(qn(tag_name))
            if property_element is None:
                property_element = parse_xml(
                    f'<{tag_name} {nsdecls("w")}/>'
                )
                run_properties.append(property_element)
            property_element.set(qn("w:val"), value)

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "table-differing-size-first.docx"
            second_path = tmp / "table-differing-size-second.docx"
            output_path = tmp / "table-differing-size-merged.docx"

            first = Document()
            set_run_default(first, "w:sz", "20")
            set_run_default(first, "w:color", "000000")
            table_style = first.styles.add_style(
                "ColorOnlyTableStyle",
                format_paper_module.WD_STYLE_TYPE.TABLE,
            )
            table_style.element.append(
                parse_xml(
                    f'<w:rPr {nsdecls("w")}>\n'
                    '<w:color w:val="E01020"/>'
                    '</w:rPr>'
                )
            )
            table = first.add_table(rows=1, cols=1)
            table.style = table_style
            table.cell(0, 0).text = "COVER DIFFERENTIAL DEFAULTS"
            first.save(first_path)

            second = Document()
            set_run_default(second, "w:sz", "48")
            set_run_default(second, "w:color", "000000")
            second.add_paragraph("BODY")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            merged_table = next(
                candidate
                for candidate in merged.tables
                if candidate.cell(0, 0).text == "COVER DIFFERENTIAL DEFAULTS"
            )
            copied_table_style = merged.styles.element.get_by_id(
                merged_table._tbl.tblPr.tblStyle.val
            )
            self.assertEqual(
                copied_table_style.find(".//" + qn("w:color")).get(
                    qn("w:val")
                ),
                "E01020",
            )
            paragraph = merged_table.cell(0, 0).paragraphs[0]
            copied_paragraph_style = merged.styles.element.get_by_id(
                paragraph._element.pPr.pStyle.val
            )
            copied_run_properties = copied_paragraph_style.find(qn("w:rPr"))
            self.assertEqual(
                copied_run_properties.find(qn("w:sz")).get(qn("w:val")),
                "20",
            )
            self.assertIsNone(
                copied_run_properties.find(qn("w:color"))
            )

    def test_concatenate_documents_honors_active_table_style_conditions(self):
        def set_default_size(document, value):
            run_properties = (
                document.styles.element.find(qn("w:docDefaults"))
                .find(qn("w:rPrDefault"))
                .find(qn("w:rPr"))
            )
            size = run_properties.find(qn("w:sz"))
            if size is None:
                size = parse_xml(f'<w:sz {nsdecls("w")}/>')
                run_properties.append(size)
            size.set(qn("w:val"), value)

        cases = (
            ("first-row-mask-disabled", "firstRow", "0000", {}, True),
            ("first-row-mask-enabled", "firstRow", "0020", {}, False),
            (
                "first-row-named-disable",
                "firstRow",
                "0020",
                {"firstRow": "0"},
                True,
            ),
            (
                "first-row-named-enable",
                "firstRow",
                "0000",
                {"firstRow": "1"},
                False,
            ),
            ("corner-missing-column", "nwCell", "0020", {}, True),
            ("corner-enabled", "nwCell", "00A0", {}, False),
            ("horizontal-band-disabled", "band1Horz", "0200", {}, True),
            ("horizontal-band-enabled", "band1Horz", "0000", {}, False),
            ("whole-table-always-active", "wholeTable", "0000", {}, False),
        )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            for case_name, condition, mask, named, expected_success in cases:
                with self.subTest(case=case_name):
                    first_path = tmp / f"{case_name}-first.docx"
                    second_path = tmp / f"{case_name}-second.docx"
                    output_path = tmp / f"{case_name}-merged.docx"

                    first = Document()
                    set_default_size(first, "20")
                    table_style = first.styles.add_style(
                        f"ConditionalStyle{case_name}",
                        format_paper_module.WD_STYLE_TYPE.TABLE,
                    )
                    table_style.element.append(
                        parse_xml(
                            f'<w:tblStylePr {nsdecls("w")} '
                            f'w:type="{condition}">'
                            '<w:rPr><w:sz w:val="30"/></w:rPr>'
                            '</w:tblStylePr>'
                        )
                    )
                    table = first.add_table(rows=2, cols=2)
                    table.style = table_style
                    table.cell(0, 0).text = "CONDITIONAL TABLE"
                    table_look = table._tbl.tblPr.find(qn("w:tblLook"))
                    table_look.attrib.clear()
                    table_look.set(qn("w:val"), mask)
                    for attribute_name, value in named.items():
                        table_look.set(qn(f"w:{attribute_name}"), value)
                    first.save(first_path)

                    second = Document()
                    set_default_size(second, "48")
                    second.add_paragraph("BODY")
                    second.save(second_path)

                    if expected_success:
                        result = concatenate_documents(
                            first_path,
                            second_path,
                            output_path,
                            restart_body_page_number=False,
                        )
                        self.assertTrue(result["concatenated"])
                    else:
                        with self.assertRaises(DocumentConcatError):
                            concatenate_documents(
                                first_path,
                                second_path,
                                output_path,
                                restart_body_page_number=False,
                            )

    def test_concatenate_documents_keeps_generated_cover_table_with_different_theme_defaults(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "generated-cover-theme-first.docx"
            second_path = tmp / "generated-cover-theme-second.docx"
            output_path = tmp / "generated-cover-theme-merged.docx"

            first = Document()
            self.assertTrue(
                generate_cover_page(
                    first,
                    {
                        "title": "主题默认值兼容测试",
                        "student_name": "测试学生",
                        "student_id": "20260001",
                    },
                )
            )
            first.save(first_path)

            second = Document()
            theme_part = second.part.part_related_by(
                format_paper_module.RT.THEME
            )
            theme_root = etree.fromstring(theme_part.blob)
            theme_root.find(
                ".//a:fontScheme/a:minorFont/a:latin",
                namespaces={
                    "a": format_paper_module._THEME_NAMESPACE,
                },
            ).set("typeface", "BodyMinorThemeFont")
            theme_part._blob = etree.tostring(
                theme_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            second.add_paragraph("BODY WITH DIFFERENT THEME")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            self.assertTrue(
                any(
                    "测试学生" in cell.text
                    for table in merged.tables
                    for row in table.rows
                    for cell in row.cells
                )
            )
            cover_table = next(
                table
                for table in merged.tables
                if any(
                    "测试学生" in cell.text
                    for row in table.rows
                    for cell in row.cells
                )
            )
            copied_table_style_id = cover_table._tbl.tblPr.tblStyle.val
            copied_table_style = merged.styles.element.get_by_id(
                copied_table_style_id
            )
            self.assertFalse(
                any(
                    len(properties)
                    for tag_name in ("w:pPr", "w:rPr")
                    for properties in copied_table_style.iter(qn(tag_name))
                )
            )

    def test_concatenate_documents_remaps_internal_style_dependencies(self):
        def add_styles(document, base_size, base_color):
            base = document.styles.add_style(
                "InternalBaseCollision",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            base.element.append(
                parse_xml(
                    f'<w:rPr {nsdecls("w")}>'
                    f'<w:sz w:val="{base_size}"/>'
                    f'<w:color w:val="{base_color}"/>'
                    '</w:rPr>'
                )
            )
            root = document.styles.add_style(
                "InternalRootCollision",
                format_paper_module.WD_STYLE_TYPE.TABLE,
            )
            root.element.append(
                parse_xml(
                    f'<w:tblStylePr {nsdecls("w")} w:type="firstRow">'
                    '<w:pPr><w:pStyle w:val="InternalBaseCollision"/>'
                    '</w:pPr></w:tblStylePr>'
                )
            )
            return root

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "internal-style-first.docx"
            second_path = tmp / "internal-style-second.docx"
            output_path = tmp / "internal-style-merged.docx"

            first = Document()
            cover_style = add_styles(first, "18", "E01020")
            cover_table = first.add_table(rows=1, cols=1)
            cover_table.style = cover_style
            cover_table.cell(0, 0).text = "COVER INTERNAL STYLE"
            first.save(first_path)

            second = Document()
            body_style = add_styles(second, "48", "1020E0")
            body_table = second.add_table(rows=1, cols=1)
            body_table.style = body_style
            body_table.cell(0, 0).text = "BODY INTERNAL STYLE"
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover_table = next(
                table
                for table in merged.tables
                if table.cell(0, 0).text == "COVER INTERNAL STYLE"
            )
            copied_root = merged.styles.element.get_by_id(
                cover_table._tbl.tblPr.tblStyle.val
            )
            copied_base_id = copied_root.find(
                ".//" + qn("w:pStyle")
            ).get(qn("w:val"))
            self.assertNotEqual(copied_base_id, "InternalBaseCollision")
            copied_base = merged.styles.element.get_by_id(copied_base_id)
            self.assertEqual(
                copied_base.find(".//" + qn("w:sz")).get(qn("w:val")),
                "18",
            )
            self.assertEqual(
                copied_base.find(".//" + qn("w:color")).get(qn("w:val")),
                "E01020",
            )

    def test_concatenate_documents_rewrites_fields_for_renamed_styles(self):
        def add_collision_styles(document, size, color):
            shared = document.styles.add_style(
                "SharedFieldStyle",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            shared.font.size = format_paper_module.Pt(size)
            shared.font.color.rgb = (
                format_paper_module.RGBColor.from_string(color)
            )
            field_only = document.styles.add_style(
                "FieldOnlyStyle",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            field_only.font.size = format_paper_module.Pt(size + 1)
            field_only.font.color.rgb = (
                format_paper_module.RGBColor.from_string(color)
            )
            return shared

        def add_simple_field(paragraph, instruction, result):
            paragraph._element.append(
                parse_xml(
                    f'<w:fldSimple {nsdecls("w")} w:instr=\'{instruction}\'>'
                    f"<w:r><w:t>{result}</w:t></w:r></w:fldSimple>"
                )
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "style-field-first.docx"
            second_path = tmp / "style-field-second.docx"
            output_path = tmp / "style-field-merged.docx"

            first = Document()
            shared_source = add_collision_styles(first, 9, "E01020")
            first.add_paragraph(
                "COVER FIELD HEADING",
                style=shared_source,
            )
            simple_paragraph = first.add_paragraph("COVER SIMPLE FIELD ")
            add_simple_field(
                simple_paragraph,
                ' STYLEREF "SharedFieldStyle" \\p ',
                "COVER SIMPLE RESULT",
            )
            comma_toc = first.add_paragraph("COVER COMMA TOC ")
            add_simple_field(
                comma_toc,
                ' TOC \\t "SharedFieldStyle,1, FieldOnlyStyle,2" \\h ',
                "COVER COMMA RESULT",
            )
            semicolon_toc = first.add_paragraph("COVER SEMICOLON TOC ")
            add_simple_field(
                semicolon_toc,
                ' TOC \\t "SharedFieldStyle;1; FieldOnlyStyle;2" \\z ',
                "COVER SEMICOLON RESULT",
            )

            header = first.sections[0].header.paragraphs[0]
            header.add_run("COVER HEADER ")
            header_field = parse_xml(
                    f'<w:p {nsdecls("w")}><w:r><w:fldChar '
                    'w:fldCharType="begin"/></w:r>'
                    '<w:r><w:instrText xml:space="preserve"> STY'
                    '</w:instrText></w:r>'
                    '<w:r><w:instrText>LEREF "Shared'
                    '</w:instrText></w:r>'
                    '<w:r><w:instrText xml:space="preserve">'
                    'FieldStyle" \\n </w:instrText></w:r>'
                    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                    '<w:r><w:t>COVER HEADER RESULT</w:t></w:r>'
                    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
                    "</w:p>"
                )
            for field_child in list(header_field):
                header._element.append(field_child)
            first.save(first_path)

            second = Document()
            add_collision_styles(second, 24, "1020E0")
            target_field = second.add_paragraph("TARGET FIELD ")
            add_simple_field(
                target_field,
                ' STYLEREF "SharedFieldStyle" \\p ',
                "TARGET FIELD RESULT",
            )
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover_heading = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "COVER FIELD HEADING"
            )
            copied_shared_name = cover_heading.style.element.find(
                qn("w:name")
            ).get(qn("w:val"))
            self.assertNotEqual(copied_shared_name, "SharedFieldStyle")

            paragraph_style_names = {
                name.get(qn("w:val"))
                for style in merged.styles.element.findall(qn("w:style"))
                if style.get(qn("w:type")) == "paragraph"
                for name in style.findall(qn("w:name"))
            }
            copied_field_only_name = next(
                name
                for name in paragraph_style_names
                if name.startswith("FieldOnlyStyle")
                and name != "FieldOnlyStyle"
            )

            simple_fields = list(
                merged.element.body.iter(qn("w:fldSimple"))
            )
            instructions_by_result = {
                field.find(".//" + qn("w:t")).text: field.get(
                    qn("w:instr")
                )
                for field in simple_fields
            }
            cover_styleref = instructions_by_result["COVER SIMPLE RESULT"]
            self.assertIn(f'"{copied_shared_name}"', cover_styleref)
            self.assertIn("\\p", cover_styleref)

            comma_instruction = instructions_by_result["COVER COMMA RESULT"]
            self.assertIn(
                f'"{copied_shared_name},1, {copied_field_only_name},2"',
                comma_instruction,
            )
            self.assertIn("\\h", comma_instruction)
            semicolon_instruction = instructions_by_result[
                "COVER SEMICOLON RESULT"
            ]
            self.assertIn(
                f'"{copied_shared_name};1; {copied_field_only_name};2"',
                semicolon_instruction,
            )
            self.assertIn("\\z", semicolon_instruction)
            self.assertIn(
                '"SharedFieldStyle"',
                instructions_by_result["TARGET FIELD RESULT"],
            )

            cover_header_part = next(
                relationship.target_part
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.HEADER
                and "COVER HEADER RESULT"
                in "".join(relationship.target_part.element.itertext())
            )
            header_instruction = "".join(
                node.text or ""
                for node in cover_header_part.element.iter(qn("w:instrText"))
            )
            self.assertIn(f'"{copied_shared_name}"', header_instruction)
            self.assertIn("\\n", header_instruction)

    def test_concatenate_documents_rewrites_style_alias_fields_without_alias_collisions(self):
        def add_alias(style, value):
            style.element.insert(
                1,
                parse_xml(
                    f'<w:aliases {nsdecls("w")} w:val="{value}"/>'
                ),
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "style-alias-first.docx"
            second_path = tmp / "style-alias-second.docx"
            output_path = tmp / "style-alias-merged.docx"

            first = Document()
            source_style = first.styles.add_style(
                "AliasSourceStyle",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            source_style.font.size = format_paper_module.Pt(9)
            add_alias(source_style, "Chapter Alias,Localized Alias")
            first.add_paragraph("COVER ALIAS STYLE", style=source_style)
            field_paragraph = first.add_paragraph("ALIAS FIELD ")
            field_paragraph._element.append(
                parse_xml(
                    f'<w:fldSimple {nsdecls("w")} '
                    'w:instr=\' STYLEREF "Chapter Alias" \'>'
                    '<w:r><w:t>ALIAS RESULT</w:t></w:r>'
                    "</w:fldSimple>"
                )
            )
            first.save(first_path)

            second = Document()
            target_collision = second.styles.add_style(
                "AliasSourceStyle",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            target_collision.font.size = format_paper_module.Pt(24)
            target_alias_owner = second.styles.add_style(
                "TargetAliasOwner",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            add_alias(target_alias_owner, "Chapter Alias")
            second.add_paragraph("BODY")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "COVER ALIAS STYLE"
            )
            copied_style = cover.style.element
            copied_name = copied_style.find(qn("w:name")).get(qn("w:val"))
            self.assertNotEqual(copied_name, "AliasSourceStyle")
            copied_aliases = copied_style.find(qn("w:aliases"))
            self.assertEqual(
                copied_aliases.get(qn("w:val")),
                "Localized Alias",
            )
            alias_field = next(
                field
                for field in merged.element.body.iter(qn("w:fldSimple"))
                if field.find(".//" + qn("w:t")).text == "ALIAS RESULT"
            )
            self.assertIn(
                f'"{copied_name}"',
                alias_field.get(qn("w:instr")),
            )
            self.assertNotIn(
                '"Chapter Alias"',
                alias_field.get(qn("w:instr")),
            )

    def test_concatenate_documents_rejects_missing_field_style_collision(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "missing-field-style-first.docx"
            second_path = tmp / "missing-field-style-second.docx"
            output_path = tmp / "missing-field-style-merged.docx"

            first = Document()
            paragraph = first.add_paragraph("MISSING FIELD STYLE ")
            paragraph._element.append(
                parse_xml(
                    f'<w:fldSimple {nsdecls("w")} '
                    'w:instr=\' STYLEREF "GhostFieldStyle" \'>'
                    '<w:r><w:t>FIELD RESULT</w:t></w:r>'
                    "</w:fldSimple>"
                )
            )
            first.save(first_path)

            second = Document()
            second.styles.add_style(
                "GhostFieldStyle",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            second.add_paragraph("BODY")
            second.save(second_path)

            with self.assertRaises(DocumentConcatError):
                concatenate_documents(
                    first_path,
                    second_path,
                    output_path,
                    restart_body_page_number=False,
                )

    def test_concatenate_documents_remaps_cyclic_style_dependencies_uniquely(self):
        def dependency(style, tag_name, style_id):
            style.element.append(
                parse_xml(
                    f'<w:{tag_name} {nsdecls("w")} w:val="{style_id}"/>'
                )
            )

        def add_style_graph(document, *, cover):
            paragraph = document.styles.add_style(
                "CycleCollision",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            character = document.styles.add_style(
                "CycleCharacter",
                format_paper_module.WD_STYLE_TYPE.CHARACTER,
            )
            dependency(paragraph, "link", character.style_id)
            dependency(character, "link", paragraph.style_id)

            if cover:
                following = document.styles.add_style(
                    "CycleCollision_1",
                    format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
                )
                dependency(paragraph, "next", following.style_id)
                dependency(following, "next", paragraph.style_id)
                following.element.append(
                    parse_xml(
                        f'<w:rPr {nsdecls("w")}>'
                        '<w:sz w:val="22"/><w:color w:val="E01020"/>'
                        '</w:rPr>'
                    )
                )

            paragraph.element.append(
                parse_xml(
                    f'<w:rPr {nsdecls("w")}>'
                    f'<w:sz w:val="{18 if cover else 48}"/>'
                    f'<w:color w:val="{"E01020" if cover else "1020E0"}"/>'
                    '</w:rPr>'
                )
            )
            character.element.append(
                parse_xml(
                    f'<w:rPr {nsdecls("w")}>'
                    f'<w:sz w:val="{20 if cover else 44}"/>'
                    f'<w:color w:val="{"20A040" if cover else "A02080"}"/>'
                    '</w:rPr>'
                )
            )
            return paragraph, character

        def dependency_id(style, tag_name):
            style_element = getattr(style, "element", style)
            return style_element.find(qn(f"w:{tag_name}")).get(qn("w:val"))

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "cyclic-style-first.docx"
            second_path = tmp / "cyclic-style-second.docx"
            output_path = tmp / "cyclic-style-merged.docx"

            first = Document()
            cover_paragraph_style, cover_character_style = add_style_graph(
                first,
                cover=True,
            )
            cover = first.add_paragraph(style=cover_paragraph_style)
            cover.add_run("COVER CYCLIC STYLE", style=cover_character_style)
            first.save(first_path)

            second = Document()
            body_paragraph_style, body_character_style = add_style_graph(
                second,
                cover=False,
            )
            body = second.add_paragraph(style=body_paragraph_style)
            body.add_run("BODY CYCLIC STYLE", style=body_character_style)
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "COVER CYCLIC STYLE"
            )
            copied_paragraph = cover.style
            copied_character = cover.runs[0].style
            copied_following = merged.styles.element.get_by_id(
                dependency_id(copied_paragraph, "next")
            )
            self.assertIsNotNone(copied_following)
            self.assertEqual(
                dependency_id(copied_paragraph, "link"),
                copied_character.style_id,
            )
            self.assertEqual(
                dependency_id(copied_character, "link"),
                copied_paragraph.style_id,
            )
            self.assertEqual(
                dependency_id(copied_following, "next"),
                copied_paragraph.style_id,
            )

            styles = merged.styles.element.findall(qn("w:style"))
            style_ids = [style.get(qn("w:styleId")) for style in styles]
            style_names = [
                name.get(qn("w:val"))
                for style in styles
                if (name := style.find(qn("w:name"))) is not None
            ]
            self.assertEqual(len(style_ids), len(set(style_ids)))
            self.assertEqual(
                len(style_names),
                len({name.casefold() for name in style_names}),
            )

    def test_concatenate_documents_remaps_numbering_bound_to_colliding_style(self):
        def add_numbered_style(document, marker):
            numbering = document.part.numbering_part.element
            abstract_num = parse_xml(
                f'<w:abstractNum {nsdecls("w")} w:abstractNumId="3000">'
                '<w:multiLevelType w:val="singleLevel"/>'
                '<w:lvl w:ilvl="0">'
                '<w:pStyle w:val="NumberedStyleCollision"/>'
                '<w:start w:val="1"/>'
                '<w:numFmt w:val="decimal"/>'
                f'<w:lvlText w:val="{marker}-%1"/>'
                '</w:lvl></w:abstractNum>'
            )
            first_num = numbering.find(qn("w:num"))
            numbering.insert(
                numbering.index(first_num)
                if first_num is not None
                else len(numbering),
                abstract_num,
            )
            numbering.append(
                parse_xml(
                    f'<w:num {nsdecls("w")} w:numId="3000">'
                    '<w:abstractNumId w:val="3000"/></w:num>'
                )
            )
            style = document.styles.add_style(
                "NumberedStyleCollision",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            style.element.append(
                parse_xml(
                    f'<w:pPr {nsdecls("w")}><w:numPr>'
                    '<w:ilvl w:val="0"/><w:numId w:val="3000"/>'
                    '</w:numPr></w:pPr>'
                )
            )
            return style

        def style_numbering_marker(document, style):
            num_id = style.element.find(
                ".//" + qn("w:numId")
            ).get(qn("w:val"))
            num = next(
                node
                for node in document.part.numbering_part.element.findall(
                    qn("w:num")
                )
                if node.get(qn("w:numId")) == num_id
            )
            abstract_id = num.find(qn("w:abstractNumId")).get(qn("w:val"))
            abstract_num = next(
                node
                for node in document.part.numbering_part.element.findall(
                    qn("w:abstractNum")
                )
                if node.get(qn("w:abstractNumId")) == abstract_id
            )
            return (
                num_id,
                abstract_num.find(".//" + qn("w:lvlText")).get(qn("w:val")),
                abstract_num.find(".//" + qn("w:pStyle")).get(qn("w:val")),
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "style-numbering-first.docx"
            second_path = tmp / "style-numbering-second.docx"
            output_path = tmp / "style-numbering-merged.docx"

            first = Document()
            first.add_paragraph(
                "COVER STYLE NUMBERING",
                style=add_numbered_style(first, "C"),
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph(
                "BODY STYLE NUMBERING",
                style=add_numbered_style(second, "B"),
            )
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "COVER STYLE NUMBERING"
            )
            body = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "BODY STYLE NUMBERING"
            )
            cover_num_id, cover_marker, cover_numbering_style = style_numbering_marker(
                merged,
                cover.style,
            )
            body_num_id, body_marker, body_numbering_style = style_numbering_marker(
                merged,
                body.style,
            )
            self.assertNotEqual(cover.style.style_id, body.style.style_id)
            self.assertNotEqual(cover_num_id, body_num_id)
            self.assertEqual(cover_marker, "C-%1")
            self.assertEqual(body_marker, "B-%1")
            self.assertEqual(cover_numbering_style, cover.style.style_id)
            self.assertEqual(body_numbering_style, body.style.style_id)

    def test_concatenate_documents_allocates_each_style_numbering_reference(self):
        def add_numbering(document, num_id, marker):
            numbering = document.part.numbering_part.element
            abstract_num = parse_xml(
                f'<w:abstractNum {nsdecls("w")} '
                f'w:abstractNumId="{num_id}">'
                '<w:multiLevelType w:val="singleLevel"/>'
                '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
                '<w:numFmt w:val="decimal"/>'
                f'<w:lvlText w:val="{marker}-%1"/>'
                '</w:lvl></w:abstractNum>'
            )
            first_num = numbering.find(qn("w:num"))
            numbering.insert(
                numbering.index(first_num)
                if first_num is not None
                else len(numbering),
                abstract_num,
            )
            numbering.append(
                parse_xml(
                    f'<w:num {nsdecls("w")} w:numId="{num_id}">'
                    f'<w:abstractNumId w:val="{num_id}"/></w:num>'
                )
            )

        def add_multi_number_style(document, marker_prefix):
            style = document.styles.add_style(
                "MultiNumberCollision",
                format_paper_module.WD_STYLE_TYPE.TABLE,
            )
            for index, num_id in enumerate((3100, 3101)):
                add_numbering(document, num_id, f"{marker_prefix}{index}")
                style.element.append(
                    parse_xml(
                        f'<w:tblStylePr {nsdecls("w")} '
                        f'w:type="{"firstRow" if index == 0 else "lastRow"}">'
                        '<w:pPr><w:numPr><w:ilvl w:val="0"/>'
                        f'<w:numId w:val="{num_id}"/>'
                        '</w:numPr></w:pPr></w:tblStylePr>'
                    )
                )
            return style

        def numbering_marker(document, num_id):
            numbering = document.part.numbering_part.element
            num = next(
                node
                for node in numbering.findall(qn("w:num"))
                if node.get(qn("w:numId")) == num_id
            )
            abstract_id = num.find(qn("w:abstractNumId")).get(qn("w:val"))
            abstract_num = next(
                node
                for node in numbering.findall(qn("w:abstractNum"))
                if node.get(qn("w:abstractNumId")) == abstract_id
            )
            return abstract_num.find(".//" + qn("w:lvlText")).get(qn("w:val"))

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "multi-number-style-first.docx"
            second_path = tmp / "multi-number-style-second.docx"
            output_path = tmp / "multi-number-style-merged.docx"

            first = Document()
            cover_style = add_multi_number_style(first, "C")
            cover_table = first.add_table(rows=1, cols=1)
            cover_table.style = cover_style
            cover_table.cell(0, 0).text = "COVER MULTI NUMBER STYLE"
            first.save(first_path)

            second = Document()
            body_style = add_multi_number_style(second, "B")
            body_table = second.add_table(rows=1, cols=1)
            body_table.style = body_style
            body_table.cell(0, 0).text = "BODY MULTI NUMBER STYLE"
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover_table = next(
                table
                for table in merged.tables
                if table.cell(0, 0).text == "COVER MULTI NUMBER STYLE"
            )
            cover_style_id = cover_table._tbl.tblPr.tblStyle.val
            copied_style = merged.styles.element.get_by_id(cover_style_id)
            copied_num_ids = [
                node.get(qn("w:val"))
                for node in copied_style.iter(qn("w:numId"))
            ]
            self.assertEqual(len(copied_num_ids), 2)
            self.assertEqual(len(set(copied_num_ids)), 2)
            self.assertEqual(
                {
                    numbering_marker(merged, num_id)
                    for num_id in copied_num_ids
                },
                {"C0-%1", "C1-%1"},
            )

    def test_concatenate_documents_rejects_missing_style_and_numbering_collisions(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)

            with self.subTest(reference="style"):
                first_path = tmp / "missing-style-first.docx"
                second_path = tmp / "missing-style-second.docx"
                output_path = tmp / "missing-style-merged.docx"

                first = Document()
                paragraph = first.add_paragraph("MISSING COVER STYLE")
                paragraph._element.get_or_add_pPr().append(
                    parse_xml(
                        f'<w:pStyle {nsdecls("w")} w:val="GhostStyle"/>'
                    )
                )
                first.save(first_path)

                second = Document()
                second.styles.add_style(
                    "GhostStyle",
                    format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
                )
                second.add_paragraph("BODY")
                second.save(second_path)

                with self.assertRaises(DocumentConcatError):
                    concatenate_documents(
                        first_path,
                        second_path,
                        output_path,
                        restart_body_page_number=False,
                    )

            with self.subTest(reference="numbering"):
                first_path = tmp / "missing-number-first.docx"
                second_path = tmp / "missing-number-second.docx"
                output_path = tmp / "missing-number-merged.docx"

                first = Document()
                paragraph = first.add_paragraph("MISSING COVER NUMBERING")
                paragraph._element.get_or_add_pPr().append(
                    parse_xml(
                        f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/>'
                        '<w:numId w:val="3200"/></w:numPr>'
                    )
                )
                first.save(first_path)

                second = Document()
                numbering = second.part.numbering_part.element
                first_num = numbering.find(qn("w:num"))
                abstract_num = parse_xml(
                    f'<w:abstractNum {nsdecls("w")} '
                    'w:abstractNumId="3200"><w:lvl w:ilvl="0">'
                    '<w:start w:val="1"/><w:numFmt w:val="decimal"/>'
                    '<w:lvlText w:val="B-%1"/>'
                    '</w:lvl></w:abstractNum>'
                )
                numbering.insert(
                    numbering.index(first_num)
                    if first_num is not None
                    else len(numbering),
                    abstract_num,
                )
                numbering.append(
                    parse_xml(
                        f'<w:num {nsdecls("w")} w:numId="3200">'
                        '<w:abstractNumId w:val="3200"/></w:num>'
                    )
                )
                second.add_paragraph("BODY")
                second.save(second_path)

                with self.assertRaises(DocumentConcatError):
                    concatenate_documents(
                        first_path,
                        second_path,
                        output_path,
                        restart_body_page_number=False,
                    )

    def test_concatenate_documents_handles_style_chains_beyond_recursion_limit(self):
        def add_deep_style_chain(document, terminal_size, terminal_color):
            styles = document.styles.element
            chain_length = 1050
            for index in range(chain_length):
                dependency = (
                    f'<w:basedOn w:val="DeepStyle{index + 1}"/>'
                    if index + 1 < chain_length
                    else ""
                )
                formatting = (
                    f'<w:rPr><w:sz w:val="{terminal_size}"/>'
                    f'<w:color w:val="{terminal_color}"/></w:rPr>'
                    if index + 1 == chain_length
                    else ""
                )
                styles.append(
                    parse_xml(
                        f'<w:style {nsdecls("w")} w:type="paragraph" '
                        f'w:customStyle="1" w:styleId="DeepStyle{index}">'
                        f'<w:name w:val="DeepStyle{index}"/>'
                        f'{dependency}{formatting}</w:style>'
                    )
                )
            paragraph = document.add_paragraph("PLACEHOLDER")
            paragraph._element.get_or_add_pPr().append(
                parse_xml(
                    f'<w:pStyle {nsdecls("w")} w:val="DeepStyle0"/>'
                )
            )
            return paragraph

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "deep-style-first.docx"
            second_path = tmp / "deep-style-second.docx"
            output_path = tmp / "deep-style-merged.docx"

            first = Document()
            cover = add_deep_style_chain(first, "18", "E01020")
            cover.runs[0].text = "COVER DEEP STYLE"
            first.save(first_path)

            second = Document()
            body = add_deep_style_chain(second, "48", "1020E0")
            body.runs[0].text = "BODY DEEP STYLE"
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "COVER DEEP STYLE"
            )
            body = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "BODY DEEP STYLE"
            )
            self.assertNotEqual(cover.style.style_id, body.style.style_id)
            style_ids = [
                style.get(qn("w:styleId"))
                for style in merged.styles.element.findall(qn("w:style"))
            ]
            self.assertEqual(len(style_ids), len(set(style_ids)))

    def test_style_chain_merge_indexes_each_styles_part_once(self):
        chain_length = 64

        def add_style_chain(document, terminal_color):
            styles = document.styles.element
            for index in range(chain_length):
                dependency = (
                    f'<w:basedOn w:val="IndexedStyle{index + 1}"/>'
                    if index + 1 < chain_length
                    else ""
                )
                formatting = (
                    f'<w:rPr><w:color w:val="{terminal_color}"/></w:rPr>'
                    if index + 1 == chain_length
                    else ""
                )
                styles.append(
                    parse_xml(
                        f'<w:style {nsdecls("w")} w:type="paragraph" '
                        f'w:customStyle="1" w:styleId="IndexedStyle{index}">'
                        f'<w:name w:val="IndexedStyle{index}"/>'
                        f'{dependency}{formatting}</w:style>'
                    )
                )

        master = Document()
        add_style_chain(master, "E01020")
        source = Document()
        add_style_chain(source, "1020E0")
        composer = format_paper_module._create_structure_preserving_composer(
            master
        )
        composer._create_style_id_mapping(source)
        original_index_builder = composer._build_style_index
        original_name_builder = composer._build_target_style_names

        with (
            patch.object(
                composer,
                "_build_style_index",
                side_effect=original_index_builder,
            ) as build_index,
            patch.object(
                composer,
                "_build_target_style_names",
                side_effect=original_name_builder,
            ) as build_names,
        ):
            mapped_root = composer._ensure_inserted_style_mapping(
                source,
                "IndexedStyle0",
            )
            for target_id in composer._academic_style_id_mapping.values():
                self.assertIsNotNone(
                    composer._style_element_by_id(master, target_id)
                )

        self.assertEqual(build_index.call_count, 2)
        self.assertEqual(build_names.call_count, 1)
        mapping = composer._academic_style_id_mapping
        self.assertEqual(
            set(mapping),
            {f"IndexedStyle{index}" for index in range(chain_length)},
        )
        self.assertEqual(mapped_root, mapping["IndexedStyle0"])
        self.assertEqual(len(set(mapping.values())), chain_length)
        for index in range(chain_length - 1):
            copied_style = composer._style_element_by_id(
                master,
                mapping[f"IndexedStyle{index}"],
            )
            self.assertEqual(
                copied_style.find(qn("w:basedOn")).get(qn("w:val")),
                mapping[f"IndexedStyle{index + 1}"],
            )

    def test_concatenate_documents_preserves_chart_and_chartex_theme_semantics(self):
        drawing_namespace = format_paper_module._THEME_NAMESPACE
        cases = (
            {
                "name": "chart-generated",
                "extended": False,
                "override": "none",
                "expected_accent": "D01020",
                "sidecar": True,
                "part_collision": True,
            },
            {
                "name": "chartex-partial",
                "extended": True,
                "override": "partial",
                "expected_accent": "AA5500",
                "sidecar": True,
                "part_collision": False,
            },
            {
                "name": "chart-complete",
                "extended": False,
                "override": "complete",
                "expected_accent": "D01020",
                "sidecar": False,
                "part_collision": False,
            },
        )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            for case in cases:
                with self.subTest(case=case["name"]):
                    first_path = tmp / f'{case["name"]}-first.docx'
                    second_path = tmp / f'{case["name"]}-second.docx'
                    output_path = tmp / f'{case["name"]}-merged.docx'

                    first = Document()
                    self.customize_test_theme(
                        first,
                        accent1="D01020",
                        major_latin="CoverChartFont",
                        format_name="CoverChartFormat",
                        background_mapping="accent2",
                    )
                    override_specs = ()
                    original_override_blob = None
                    if case["override"] == "partial":
                        original_override_blob = self.make_test_theme_override_blob(
                            first,
                            included_schemes=("clrScheme",),
                            accent1="AA5500",
                        )
                        override_specs = (
                            (
                                original_override_blob,
                                format_paper_module.CT.OFC_THEME_OVERRIDE,
                                False,
                            ),
                        )
                    elif case["override"] == "complete":
                        original_override_blob = self.make_test_theme_override_blob(
                            first
                        )
                        override_specs = (
                            (
                                original_override_blob,
                                format_paper_module.CT.OFC_THEME_OVERRIDE,
                                False,
                            ),
                        )
                    self.add_test_chart(
                        first,
                        case["name"],
                        extended=case["extended"],
                        override_specs=override_specs,
                        add_sidecar=case["sidecar"],
                    )
                    first.save(first_path)

                    second = Document()
                    self.customize_test_theme(
                        second,
                        accent1="1020D0",
                        major_latin="BodyChartFont",
                        format_name="BodyChartFormat",
                        background_mapping="accent1",
                    )
                    if case["part_collision"]:
                        collision_part = format_paper_module.Part(
                            format_paper_module.PackURI(
                                "/word/theme/themeOverride1.xml"
                            ),
                            "application/octet-stream",
                            b"<collision/>",
                            second.part.package,
                        )
                        second.part.relate_to(
                            collision_part,
                            "urn:test:theme-override-name-collision",
                        )
                    second.add_paragraph("BODY")
                    second.save(second_path)

                    result = concatenate_documents(
                        first_path,
                        second_path,
                        output_path,
                        restart_body_page_number=False,
                    )

                    self.assertTrue(result["concatenated"])
                    merged = Document(output_path)
                    chart_relationship_type = (
                        format_paper_module._CHARTEX_RELATIONSHIP_TYPE
                        if case["extended"]
                        else format_paper_module.RT.CHART
                    )
                    chart_relationship = next(
                        relationship
                        for relationship in merged.part.rels.values()
                        if relationship.reltype == chart_relationship_type
                    )
                    chart_part = chart_relationship.target_part
                    override_relationships = [
                        relationship
                        for relationship in chart_part.rels.values()
                        if relationship.reltype
                        == format_paper_module.RT.THEME_OVERRIDE
                    ]
                    self.assertEqual(len(override_relationships), 1)
                    self.assertFalse(override_relationships[0].is_external)
                    override_part = override_relationships[0].target_part
                    override_root = etree.fromstring(override_part.blob)
                    self.assertEqual(
                        [
                            etree.QName(child).localname
                            for child in override_root
                        ],
                        ["clrScheme", "fontScheme", "fmtScheme"],
                    )
                    self.assertEqual(
                        override_root.find(
                            ".//a:clrScheme/a:accent1/*",
                            namespaces={"a": drawing_namespace},
                        ).get("val"),
                        case["expected_accent"],
                    )
                    self.assertEqual(
                        override_root.find(
                            ".//a:fontScheme/a:majorFont/a:latin",
                            namespaces={"a": drawing_namespace},
                        ).get("typeface"),
                        "CoverChartFont",
                    )
                    self.assertEqual(
                        override_root.find(
                            ".//a:fmtScheme",
                            namespaces={"a": drawing_namespace},
                        ).get("name"),
                        "CoverChartFormat",
                    )
                    if case["override"] == "complete":
                        self.assertEqual(
                            override_part.blob,
                            original_override_blob,
                        )

                    chart_root = etree.fromstring(chart_part.blob)
                    mapping_tag = (
                        format_paper_module._CHARTEX_COLOR_MAPPING_TAG
                        if case["extended"]
                        else format_paper_module._CHART_COLOR_MAPPING_TAG
                    )
                    mapping = chart_root.find(mapping_tag)
                    self.assertIsNotNone(mapping)
                    self.assertEqual(mapping.get("bg1"), "accent2")
                    self.assertEqual(mapping.get("tx1"), "dk1")

                    if case["sidecar"]:
                        sidecars = {
                            relationship.reltype: relationship.target_part
                            for relationship in chart_part.rels.values()
                            if relationship.reltype
                            in {
                                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                                format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                            }
                        }
                        expected_sidecars = {
                            format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE: (
                                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                                self.make_test_chart_style_blob(),
                            ),
                            format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE: (
                                format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                                self.make_test_chart_color_style_blob(),
                            ),
                        }
                        self.assertEqual(
                            set(sidecars),
                            set(expected_sidecars),
                        )
                        for relationship_type, sidecar in sidecars.items():
                            expected_content_type, expected_blob = (
                                expected_sidecars[relationship_type]
                            )
                            self.assertEqual(
                                sidecar.content_type,
                                expected_content_type,
                            )
                            self.assertEqual(sidecar.blob, expected_blob)
                    if case["part_collision"]:
                        self.assertNotEqual(
                            str(override_part.partname).casefold(),
                            "/word/theme/themeoverride1.xml".casefold(),
                        )

                    theme_part = merged.part.part_related_by(
                        format_paper_module.RT.THEME
                    )
                    merged_theme = etree.fromstring(theme_part.blob)
                    self.assertEqual(
                        merged_theme.find(
                            ".//a:clrScheme/a:accent1/*",
                            namespaces={"a": drawing_namespace},
                        ).get("val"),
                        "1020D0",
                    )
                    with ZipFile(output_path) as archive:
                        names = archive.namelist()
                        self.assertIsNone(archive.testzip())
                    self.assertEqual(
                        len(names),
                        len({name.casefold() for name in names}),
                    )

    def test_concatenate_documents_preserves_valid_chart_style_extensions(self):
        chart_style_namespace = format_paper_module._CHART_STYLE_NAMESPACE
        drawing_namespace = format_paper_module._THEME_NAMESPACE
        drawing_2010_namespace = (
            format_paper_module._DRAWING_2010_NAMESPACE
        )

        def serialize(root):
            return etree.tostring(
                root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

        parsed_style_root = etree.fromstring(
            self.make_test_chart_style_blob()
        )
        mc_namespace = (
            format_paper_module._MARKUP_COMPATIBILITY_NAMESPACE
        )
        vendor_namespace = "urn:test:chart-style-vendor"
        drawing_2012_namespace = (
            "http://schemas.microsoft.com/office/drawing/2012/main"
        )
        style_root = etree.Element(
            parsed_style_root.tag,
            nsmap={
                **parsed_style_root.nsmap,
                "mc": mc_namespace,
                "v": vendor_namespace,
                "a15": drawing_2012_namespace,
            },
        )
        style_root.attrib.update(parsed_style_root.attrib)
        style_root.extend(list(parsed_style_root))
        style_root.set(f"{{{mc_namespace}}}Ignorable", "v")
        style_root.set(
            f"{{{mc_namespace}}}PreserveAttributes",
            "v:rootAttribute",
        )
        style_root.set(
            f"{{{mc_namespace}}}PreserveElements",
            "v:ignored",
        )
        style_root.set(f"{{{vendor_namespace}}}rootAttribute", "preserve")
        style_root.set(
            "{http://www.w3.org/XML/1998/namespace}space",
            "preserve",
        )
        style_root.insert(1, etree.Comment("schema misc node"))
        ignored_extension = etree.Element(
            f"{{{vendor_namespace}}}ignored",
        )
        etree.SubElement(
            ignored_extension,
            f"{{{chart_style_namespace}}}bogus",
        )
        style_root.insert(2, ignored_extension)

        category_axis = style_root.find(
            f".//{{{chart_style_namespace}}}categoryAxis"
        )
        category_axis_index = style_root.index(category_axis)
        style_root.remove(category_axis)
        process_wrapper = etree.Element(
            f"{{{vendor_namespace}}}processWrapper",
        )
        process_wrapper.set(
            f"{{{mc_namespace}}}ProcessContent",
            "v:processWrapper",
        )
        process_wrapper.set(
            "{http://www.w3.org/XML/1998/namespace}id",
            "process-wrapper",
        )
        process_wrapper.append(category_axis)
        style_root.insert(category_axis_index, process_wrapper)

        chart_area = style_root.find(
            f".//{{{chart_style_namespace}}}chartArea"
        )
        chart_area_index = style_root.index(chart_area)
        style_root.remove(chart_area)
        supported_alternate = etree.Element(
            f"{{{mc_namespace}}}AlternateContent"
        )
        supported_alternate.set(
            f"{{{vendor_namespace}}}qualifiedAttribute",
            "preserve",
        )
        supported_choice = etree.SubElement(
            supported_alternate,
            f"{{{mc_namespace}}}Choice",
            Requires="a15",
        )
        supported_choice.append(chart_area)
        etree.SubElement(
            supported_alternate,
            f"{{{vendor_namespace}}}ignorableBetweenChoices",
        )
        supported_fallback = etree.SubElement(
            supported_alternate,
            f"{{{mc_namespace}}}Fallback",
        )
        etree.SubElement(
            supported_fallback,
            f"{{{chart_style_namespace}}}bogusFallback",
        )
        style_root.insert(chart_area_index, supported_alternate)

        data_label = style_root.find(
            f".//{{{chart_style_namespace}}}dataLabel"
        )
        data_label_index = style_root.index(data_label)
        style_root.remove(data_label)
        fallback_alternate = etree.Element(
            f"{{{mc_namespace}}}AlternateContent"
        )
        unsupported_choice = etree.SubElement(
            fallback_alternate,
            f"{{{mc_namespace}}}Choice",
            Requires="v",
        )
        unsupported_choice.set(
            f"{{{mc_namespace}}}MustUnderstand",
            "v",
        )
        etree.SubElement(
            unsupported_choice,
            f"{{{chart_style_namespace}}}bogusChoice",
        )
        selected_fallback = etree.SubElement(
            fallback_alternate,
            f"{{{mc_namespace}}}Fallback",
        )
        selected_fallback.append(data_label)
        style_root.insert(data_label_index, fallback_alternate)

        axis_title = style_root.find(
            f"{{{chart_style_namespace}}}axisTitle"
        )
        axis_title.insert(
            1,
            etree.ProcessingInstruction("chart-style", "preserve"),
        )
        references = {
            name: axis_title.find(f"{{{chart_style_namespace}}}{name}")
            for name in ("lnRef", "fillRef", "effectRef")
        }
        custom_style_color = etree.SubElement(
            references["lnRef"],
            f"{{{chart_style_namespace}}}styleClr",
            val="vendorColorToken",
        )
        custom_style_color.append(etree.Comment("inside transform list"))
        etree.SubElement(
            custom_style_color,
            f"{{{drawing_namespace}}}lumMod",
            val="80000",
        )
        etree.SubElement(
            references["fillRef"],
            f"{{{chart_style_namespace}}}styleClr",
            val="",
        )
        legacy_color = etree.SubElement(
            references["effectRef"],
            f"{{{drawing_namespace}}}srgbClr",
            val="A1B2C3",
        )
        legacy_color.set(
            f"{{{drawing_2010_namespace}}}legacySpreadsheetColorIndex",
            "80",
        )
        font_reference = axis_title.find(
            f"{{{chart_style_namespace}}}fontRef"
        )
        shape_properties = etree.Element(
            f"{{{chart_style_namespace}}}spPr",
            bwMode="auto",
        )
        transform_2d = etree.SubElement(
            shape_properties,
            f"{{{drawing_namespace}}}xfrm",
            rot="-2147483648",
            flipH="1",
            flipV="false",
        )
        etree.SubElement(
            transform_2d,
            f"{{{drawing_namespace}}}off",
            x="-27273042329600",
            y="27273042316900",
        )
        etree.SubElement(
            transform_2d,
            f"{{{drawing_namespace}}}ext",
            cx="0",
            cy="2147483647",
        )
        preset_geometry = etree.SubElement(
            shape_properties,
            f"{{{drawing_namespace}}}prstGeom",
            prst="roundRect",
        )
        preset_adjustments = etree.SubElement(
            preset_geometry,
            f"{{{drawing_namespace}}}avLst",
        )
        etree.SubElement(
            preset_adjustments,
            f"{{{drawing_namespace}}}gd",
            name="adj",
            fmla="val 50000",
        )
        etree.SubElement(
            preset_adjustments,
            f"{{{drawing_namespace}}}gd",
            name="",
            fmla="",
        )
        solid_fill = etree.SubElement(
            shape_properties,
            f"{{{drawing_namespace}}}solidFill",
        )
        etree.SubElement(
            solid_fill,
            f"{{{drawing_namespace}}}schemeClr",
            val="accent1",
        )
        outline = etree.SubElement(
            shape_properties,
            f"{{{drawing_namespace}}}ln",
            w="20116800",
            cap="rnd",
            cmpd="sng",
            algn="ctr",
        )
        etree.SubElement(
            outline,
            f"{{{drawing_namespace}}}prstDash",
            val="lgDashDotDot",
        )
        effect_list = etree.SubElement(
            shape_properties,
            f"{{{drawing_namespace}}}effectLst",
        )
        outer_shadow = etree.SubElement(
            effect_list,
            f"{{{drawing_namespace}}}outerShdw",
            blurRad="50800",
            dist="38100",
            dir="5400000",
            algn="ctr",
            rotWithShape="0",
        )
        shadow_color = etree.SubElement(
            outer_shadow,
            f"{{{drawing_namespace}}}srgbClr",
            val="000000",
        )
        etree.SubElement(
            shadow_color,
            f"{{{drawing_namespace}}}alpha",
            val="40000",
        )
        scene = etree.SubElement(
            shape_properties,
            f"{{{drawing_namespace}}}scene3d",
        )
        camera = etree.SubElement(
            scene,
            f"{{{drawing_namespace}}}camera",
            prst="orthographicFront",
            fov="10800000",
            zoom="0",
        )
        etree.SubElement(
            camera,
            f"{{{drawing_namespace}}}rot",
            lat="0",
            lon="21599999",
            rev="1",
        )
        light_rig = etree.SubElement(
            scene,
            f"{{{drawing_namespace}}}lightRig",
            rig="balanced",
            dir="tr",
        )
        etree.SubElement(
            light_rig,
            f"{{{drawing_namespace}}}rot",
            lat="1",
            lon="2",
            rev="3",
        )
        backdrop = etree.SubElement(
            scene,
            f"{{{drawing_namespace}}}backdrop",
        )
        etree.SubElement(
            backdrop,
            f"{{{drawing_namespace}}}anchor",
            x="-27273042329600",
            y="0",
            z="27273042316900",
        )
        etree.SubElement(
            backdrop,
            f"{{{drawing_namespace}}}norm",
            dx="0",
            dy="1",
            dz="2",
        )
        etree.SubElement(
            backdrop,
            f"{{{drawing_namespace}}}up",
            dx="2",
            dy="1",
            dz="0",
        )
        shape_3d = etree.SubElement(
            shape_properties,
            f"{{{drawing_namespace}}}sp3d",
            z="-27273042329600",
            extrusionH="2147483647",
            contourW="0",
            prstMaterial="softmetal",
        )
        etree.SubElement(
            shape_3d,
            f"{{{drawing_namespace}}}bevelT",
            w="0",
            h="2147483647",
            prst="artDeco",
        )
        extrusion_color = etree.SubElement(
            shape_3d,
            f"{{{drawing_namespace}}}extrusionClr",
        )
        etree.SubElement(
            extrusion_color,
            f"{{{drawing_namespace}}}schemeClr",
            val="accent2",
        )
        contour_color = etree.SubElement(
            shape_3d,
            f"{{{drawing_namespace}}}contourClr",
        )
        etree.SubElement(
            contour_color,
            f"{{{drawing_namespace}}}srgbClr",
            val="123456",
        )
        font_reference.addnext(shape_properties)
        character_properties = etree.Element(
            f"{{{chart_style_namespace}}}defRPr",
            b="1",
            sz="100",
            kern="400000",
            spc="-400000",
            baseline="-2147483648",
            smtId="4294967295",
            u="wavyDbl",
            strike="dblStrike",
            cap="all",
        )
        highlight = etree.SubElement(
            character_properties,
            f"{{{drawing_namespace}}}highlight",
        )
        etree.SubElement(
            highlight,
            f"{{{drawing_namespace}}}schemeClr",
            val="accent4",
        )
        underline_line = etree.SubElement(
            character_properties,
            f"{{{drawing_namespace}}}uLn",
            w="1",
        )
        etree.SubElement(
            underline_line,
            f"{{{drawing_namespace}}}prstDash",
            val="solid",
        )
        underline_fill = etree.SubElement(
            character_properties,
            f"{{{drawing_namespace}}}uFill",
        )
        underline_solid_fill = etree.SubElement(
            underline_fill,
            f"{{{drawing_namespace}}}solidFill",
        )
        etree.SubElement(
            underline_solid_fill,
            f"{{{drawing_namespace}}}srgbClr",
            val="ABCDEF",
        )
        etree.SubElement(
            character_properties,
            f"{{{drawing_namespace}}}latin",
            typeface="Test Latin",
            panose="00112233445566778899",
            pitchFamily="-128",
            charset="127",
        )
        etree.SubElement(
            character_properties,
            f"{{{drawing_namespace}}}ea",
            typeface="",
        )
        etree.SubElement(
            character_properties,
            f"{{{drawing_namespace}}}cs",
        )
        etree.SubElement(
            character_properties,
            f"{{{drawing_namespace}}}sym",
        )
        click_hyperlink = etree.SubElement(
            character_properties,
            f"{{{drawing_namespace}}}hlinkClick",
            invalidUrl="not a URI",
            history="1",
        )
        click_hyperlink.set(
            f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}id",
            "",
        )
        mouse_over_hyperlink = etree.SubElement(
            character_properties,
            f"{{{drawing_namespace}}}hlinkMouseOver",
            tooltip="tooltip",
            endSnd="false",
        )
        mouse_over_hyperlink.set(
            f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}id",
            "",
        )
        etree.SubElement(
            character_properties,
            f"{{{drawing_namespace}}}rtl",
            val="true",
        )
        shape_properties.addnext(character_properties)
        body_properties = etree.Element(
            f"{{{chart_style_namespace}}}bodyPr",
            rot="2147483647",
            numCol="16",
            spcCol="2147483647",
            wrap="square",
            vert="wordArtVertRtl",
            anchor="ctr",
            upright="true",
        )
        etree.SubElement(
            body_properties,
            f"{{{drawing_namespace}}}noAutofit",
        )
        character_properties.addnext(body_properties)
        entry_extension_list = etree.SubElement(
            axis_title,
            f"{{{chart_style_namespace}}}extLst",
        )
        entry_extension = etree.SubElement(
            entry_extension_list,
            f"{{{drawing_namespace}}}ext",
            uri="",
        )
        entry_extension_payload = etree.SubElement(
            entry_extension,
            "{urn:test:chart-style}payload",
        )
        entry_extension_payload.set(
            f"{{{mc_namespace}}}Ignorable",
            "private-prefix-is-opaque",
        )

        # Exercise the DrawingML fill/line particles that occur in real
        # Office chart-style sidecars, while retaining the exact source bytes.
        category_axis = style_root.find(
            f".//{{{chart_style_namespace}}}categoryAxis"
        )
        category_shape = etree.SubElement(
            category_axis,
            f"{{{chart_style_namespace}}}spPr",
        )
        category_gradient = etree.SubElement(
            category_shape,
            f"{{{drawing_namespace}}}gradFill",
            flip="none",
            rotWithShape="1",
        )
        category_stops = etree.SubElement(
            category_gradient,
            f"{{{drawing_namespace}}}gsLst",
        )
        for position, luminance in (("0", "35000"), ("100000", "65000")):
            stop = etree.SubElement(
                category_stops,
                f"{{{drawing_namespace}}}gs",
                pos=position,
            )
            color = etree.SubElement(
                stop,
                f"{{{drawing_namespace}}}schemeClr",
                val="accent1",
            )
            etree.SubElement(
                color,
                f"{{{drawing_namespace}}}lumMod",
                val=luminance,
            )
        path_shade = etree.SubElement(
            category_gradient,
            f"{{{drawing_namespace}}}path",
            path="circle",
        )
        etree.SubElement(
            path_shade,
            f"{{{drawing_namespace}}}fillToRect",
            l="50000",
            t="50000",
            r="50000",
            b="50000",
        )
        etree.SubElement(
            category_gradient,
            f"{{{drawing_namespace}}}tileRect",
        )

        chart_area = style_root.find(
            f".//{{{chart_style_namespace}}}chartArea"
        )
        chart_area_shape = etree.SubElement(
            chart_area,
            f"{{{chart_style_namespace}}}spPr",
        )
        linear_gradient = etree.SubElement(
            chart_area_shape,
            f"{{{drawing_namespace}}}gradFill",
            rotWithShape="0",
        )
        linear_stops = etree.SubElement(
            linear_gradient,
            f"{{{drawing_namespace}}}gsLst",
        )
        for position, color_value in (("0", "112233"), ("100000", "AABBCC")):
            stop = etree.SubElement(
                linear_stops,
                f"{{{drawing_namespace}}}gs",
                pos=position,
            )
            etree.SubElement(
                stop,
                f"{{{drawing_namespace}}}srgbClr",
                val=color_value,
            )
        etree.SubElement(
            linear_gradient,
            f"{{{drawing_namespace}}}lin",
            ang="21599999",
            scaled="0",
        )

        data_label = style_root.find(
            f".//{{{chart_style_namespace}}}dataLabel"
        )
        pattern_shape = etree.SubElement(
            data_label,
            f"{{{chart_style_namespace}}}spPr",
        )
        pattern_fill = etree.SubElement(
            pattern_shape,
            f"{{{drawing_namespace}}}pattFill",
            prst="zigZag",
        )
        foreground = etree.SubElement(
            pattern_fill,
            f"{{{drawing_namespace}}}fgClr",
        )
        etree.SubElement(
            foreground,
            f"{{{drawing_namespace}}}schemeClr",
            val="phClr",
        )
        background = etree.SubElement(
            pattern_fill,
            f"{{{drawing_namespace}}}bgClr",
        )
        etree.SubElement(
            background,
            f"{{{drawing_namespace}}}srgbClr",
            val="FFFFFF",
        )

        data_point = style_root.find(
            f".//{{{chart_style_namespace}}}dataPoint"
        )
        custom_line_shape = etree.SubElement(
            data_point,
            f"{{{chart_style_namespace}}}spPr",
        )
        custom_line = etree.SubElement(
            custom_line_shape,
            f"{{{drawing_namespace}}}ln",
            w="1",
        )
        custom_dash = etree.SubElement(
            custom_line,
            f"{{{drawing_namespace}}}custDash",
        )
        etree.SubElement(
            custom_dash,
            f"{{{drawing_namespace}}}ds",
            d="1",
            sp="2147483647",
        )
        etree.SubElement(
            custom_line,
            f"{{{drawing_namespace}}}miter",
            lim="0",
        )
        etree.SubElement(
            custom_line,
            f"{{{drawing_namespace}}}headEnd",
            type="triangle",
            w="sm",
            len="lg",
        )
        etree.SubElement(
            custom_line,
            f"{{{drawing_namespace}}}tailEnd",
            type="none",
            w="med",
            len="med",
        )

        data_point_3d = style_root.find(
            f".//{{{chart_style_namespace}}}dataPoint3D"
        )
        blip_shape = etree.SubElement(
            data_point_3d,
            f"{{{chart_style_namespace}}}spPr",
        )
        blip_fill = etree.SubElement(
            blip_shape,
            f"{{{drawing_namespace}}}blipFill",
            dpi="4294967295",
            rotWithShape="true",
        )
        blip = etree.SubElement(
            blip_fill,
            f"{{{drawing_namespace}}}blip",
            cstate="screen",
        )
        etree.SubElement(
            blip,
            f"{{{drawing_namespace}}}alphaCeiling",
        )
        etree.SubElement(
            blip_fill,
            f"{{{drawing_namespace}}}srcRect",
            l="-2147483648",
            r="2147483647",
        )
        stretch = etree.SubElement(
            blip_fill,
            f"{{{drawing_namespace}}}stretch",
        )
        etree.SubElement(
            stretch,
            f"{{{drawing_namespace}}}fillRect",
            t="1",
            b="2",
        )
        wireframe = style_root.find(
            f".//{{{chart_style_namespace}}}dataPointWireframe"
        )
        group_fill_shape = etree.SubElement(
            wireframe,
            f"{{{chart_style_namespace}}}spPr",
        )
        etree.SubElement(
            group_fill_shape,
            f"{{{drawing_namespace}}}grpFill",
        )
        plot_area_3d = style_root.find(
            f".//{{{chart_style_namespace}}}plotArea3D"
        )
        custom_geometry_shape = etree.SubElement(
            plot_area_3d,
            f"{{{chart_style_namespace}}}spPr",
        )
        custom_geometry = etree.SubElement(
            custom_geometry_shape,
            f"{{{drawing_namespace}}}custGeom",
        )
        adjustment_values = etree.SubElement(
            custom_geometry,
            f"{{{drawing_namespace}}}avLst",
        )
        etree.SubElement(
            adjustment_values,
            f"{{{drawing_namespace}}}gd",
            name="",
            fmla="",
        )
        guides = etree.SubElement(
            custom_geometry,
            f"{{{drawing_namespace}}}gdLst",
        )
        etree.SubElement(
            guides,
            f"{{{drawing_namespace}}}gd",
            name="g1",
            fmla="val 1",
        )
        handles = etree.SubElement(
            custom_geometry,
            f"{{{drawing_namespace}}}ahLst",
        )
        xy_handle = etree.SubElement(
            handles,
            f"{{{drawing_namespace}}}ahXY",
            gdRefX="g1",
            minX="-27273042329600",
            maxX="formulaX",
        )
        etree.SubElement(
            xy_handle,
            f"{{{drawing_namespace}}}pos",
            x="0",
            y="formulaY",
        )
        polar_handle = etree.SubElement(
            handles,
            f"{{{drawing_namespace}}}ahPolar",
            gdRefR="",
            minAng="angleFormula",
        )
        etree.SubElement(
            polar_handle,
            f"{{{drawing_namespace}}}pos",
            x="1",
            y="2",
        )
        connection_sites = etree.SubElement(
            custom_geometry,
            f"{{{drawing_namespace}}}cxnLst",
        )
        connection = etree.SubElement(
            connection_sites,
            f"{{{drawing_namespace}}}cxn",
            ang="angleFormula",
        )
        etree.SubElement(
            connection,
            f"{{{drawing_namespace}}}pos",
            x="3",
            y="4",
        )
        etree.SubElement(
            custom_geometry,
            f"{{{drawing_namespace}}}rect",
            l="lFormula",
            t="0",
            r="2147483647",
            b="bFormula",
        )
        path_list = etree.SubElement(
            custom_geometry,
            f"{{{drawing_namespace}}}pathLst",
        )
        custom_path = etree.SubElement(
            path_list,
            f"{{{drawing_namespace}}}path",
            w="2147483647",
            h="0",
            fill="norm",
            stroke="1",
            extrusionOk="0",
        )
        move_to = etree.SubElement(
            custom_path,
            f"{{{drawing_namespace}}}moveTo",
        )
        etree.SubElement(
            move_to,
            f"{{{drawing_namespace}}}pt",
            x="0",
            y="0",
        )
        line_to = etree.SubElement(
            custom_path,
            f"{{{drawing_namespace}}}lnTo",
        )
        etree.SubElement(
            line_to,
            f"{{{drawing_namespace}}}pt",
            x="xFormula",
            y="yFormula",
        )
        etree.SubElement(
            custom_path,
            f"{{{drawing_namespace}}}arcTo",
            wR="wFormula",
            hR="hFormula",
            stAng="startFormula",
            swAng="swingFormula",
        )
        etree.SubElement(
            custom_path,
            f"{{{drawing_namespace}}}close",
        )
        leader_line = style_root.find(
            f".//{{{chart_style_namespace}}}leaderLine"
        )
        effect_dag_shape = etree.SubElement(
            leader_line,
            f"{{{chart_style_namespace}}}spPr",
        )
        effect_dag = etree.SubElement(
            effect_dag_shape,
            f"{{{drawing_namespace}}}effectDag",
        )
        effect_container = etree.SubElement(
            effect_dag,
            f"{{{drawing_namespace}}}cont",
            type="sib",
            name="",
        )
        etree.SubElement(
            effect_container,
            f"{{{drawing_namespace}}}alphaBiLevel",
            thresh="0",
        )
        alpha_mod = etree.SubElement(
            effect_container,
            f"{{{drawing_namespace}}}alphaMod",
        )
        etree.SubElement(
            alpha_mod,
            f"{{{drawing_namespace}}}cont",
        )
        etree.SubElement(
            effect_container,
            f"{{{drawing_namespace}}}effect",
            ref="",
        )
        dag_shadow = etree.SubElement(
            effect_container,
            f"{{{drawing_namespace}}}outerShdw",
            dir="0",
        )
        etree.SubElement(
            dag_shadow,
            f"{{{drawing_namespace}}}schemeClr",
            val="accent3",
        )
        dag_fill_overlay = etree.SubElement(
            effect_container,
            f"{{{drawing_namespace}}}fillOverlay",
            blend="over",
        )
        etree.SubElement(
            dag_fill_overlay,
            f"{{{drawing_namespace}}}noFill",
        )

        for entry_name, value in (
            ("axisTitle", "+1.25e-3"),
            ("categoryAxis", "INF"),
            ("chartArea", "-INF"),
            ("dataLabel", "NaN"),
        ):
            entry = style_root.find(
                f".//{{{chart_style_namespace}}}{entry_name}"
            )
            line_reference = entry.find(
                f"{{{chart_style_namespace}}}lnRef"
            )
            line_width_scale = etree.Element(
                f"{{{chart_style_namespace}}}lineWidthScale"
            )
            line_width_scale.text = value
            line_reference.addnext(line_width_scale)

        root_extension_list = etree.SubElement(
            style_root,
            f"{{{chart_style_namespace}}}extLst",
        )
        root_extension = etree.SubElement(
            root_extension_list,
            f"{{{drawing_namespace}}}ext",
            uri="urn:test:chart-style",
        )
        root_extension_payload = etree.SubElement(
            root_extension,
            "{urn:test:chart-style}payload",
        )
        root_extension_payload.set(
            f"{{{mc_namespace}}}Ignorable",
            "private-prefix-is-opaque",
        )
        style_blob = serialize(style_root)

        color_root = etree.fromstring(
            self.make_test_chart_color_style_blob()
        )
        color_root.set("meth", "vendorMethod")
        color_root.insert(0, etree.Comment("before color sequence"))
        etree.SubElement(
            color_root,
            f"{{{drawing_namespace}}}scrgbClr",
            r="-2147483648",
            g="+2147483647",
            b="000000000000000000000",
        )
        etree.SubElement(
            color_root,
            f"{{{drawing_namespace}}}hslClr",
            hue="21599999",
            sat="-2147483648",
            lum="2147483647",
        )
        etree.SubElement(
            color_root,
            f"{{{drawing_namespace}}}sysClr",
            val="menuBar",
            lastClr="00A0FF",
        )
        etree.SubElement(
            color_root,
            f"{{{drawing_namespace}}}prstClr",
            val="darkGrey",
        )
        color_blob = serialize(color_root)

        empty_method_root = etree.fromstring(
            self.make_test_chart_color_style_blob()
        )
        empty_method_root.set("meth", "")
        empty_method_blob = serialize(empty_method_root)

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "valid-chart-style-extensions-first.docx"
            second_path = tmp / "valid-chart-style-extensions-second.docx"
            output_path = tmp / "valid-chart-style-extensions-merged.docx"

            first = Document()
            self.add_test_chart(
                first,
                "VALID CHART STYLE EXTENSIONS",
                sidecar_specs=(
                    (
                        format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                        style_blob,
                        format_paper_module._CHART_STYLE_CONTENT_TYPE,
                        False,
                    ),
                    (
                        format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                        color_blob,
                        format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                        False,
                    ),
                    (
                        format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                        empty_method_blob,
                        format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                        False,
                    ),
                ),
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph("BODY")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            chart_part = next(
                relationship.target_part
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.CHART
            )
            copied_style_blobs = [
                relationship.target_part.blob
                for relationship in chart_part.rels.values()
                if relationship.reltype
                == format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE
            ]
            copied_color_blobs = [
                relationship.target_part.blob
                for relationship in chart_part.rels.values()
                if relationship.reltype
                == format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE
            ]
            self.assertCountEqual(copied_style_blobs, [style_blob])
            self.assertCountEqual(
                copied_color_blobs,
                [color_blob, empty_method_blob],
            )

    def test_concatenate_documents_rejects_invalid_chart_style_sidecars(self):
        style_blob = self.make_test_chart_style_blob()
        color_blob = self.make_test_chart_color_style_blob()
        chart_style_namespace = format_paper_module._CHART_STYLE_NAMESPACE
        drawing_namespace = format_paper_module._THEME_NAMESPACE

        wrong_style_root = etree.fromstring(style_blob)
        wrong_style_root.tag = (
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}notChartStyle"
        )
        wrong_style_root_blob = etree.tostring(
            wrong_style_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        incomplete_style = etree.fromstring(style_blob)
        incomplete_style.remove(
            incomplete_style.find(
                f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}dataPoint"
            )
        )
        incomplete_style_blob = etree.tostring(
            incomplete_style,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        missing_color_method = etree.fromstring(color_blob)
        missing_color_method.attrib.pop("meth")
        missing_color_method_blob = etree.tostring(
            missing_color_method,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        def color_blob_with_model(model_name, **attributes):
            root = etree.fromstring(color_blob)
            root.remove(root[0])
            root.insert(
                0,
                etree.Element(
                    f"{{{format_paper_module._THEME_NAMESPACE}}}{model_name}",
                    **attributes,
                ),
            )
            return etree.tostring(
                root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

        invalid_scrgb_blob = color_blob_with_model(
            "scrgbClr",
            r="2147483648",
            g="0",
            b="0",
        )
        invalid_hsl_blob = color_blob_with_model(
            "hslClr",
            hue="21600000",
            sat="0",
            lum="0",
        )
        invalid_system_color_blob = color_blob_with_model(
            "sysClr",
            val="definitelyNotASystemColor",
        )
        invalid_scheme_color_blob = color_blob_with_model(
            "schemeClr",
            val="definitelyNotASchemeColor",
        )
        invalid_preset_color_blob = color_blob_with_model(
            "prstClr",
            val="definitelyNotAPresetColor",
        )
        invalid_legacy_color = etree.fromstring(color_blob)
        invalid_legacy_color_model = invalid_legacy_color[0]
        invalid_legacy_color_model.tag = (
            f"{{{format_paper_module._THEME_NAMESPACE}}}srgbClr"
        )
        invalid_legacy_color_model.set("val", "A1B2C3")
        invalid_legacy_color_model.set(
            f"{{{format_paper_module._DRAWING_2010_NAMESPACE}}}"
            "legacySpreadsheetColorIndex",
            "81",
        )
        invalid_legacy_color_blob = etree.tostring(
            invalid_legacy_color,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        def style_blob_with_line_width(value):
            root = etree.fromstring(style_blob)
            entry = root.find(
                f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}axisTitle"
            )
            line_reference = entry.find(
                f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}lnRef"
            )
            line_width_scale = etree.Element(
                f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}"
                "lineWidthScale"
            )
            line_width_scale.text = value
            line_reference.addnext(line_width_scale)
            return etree.tostring(
                root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

        invalid_line_width_blobs = {
            "python-digit-separator": style_blob_with_line_width("1_0"),
            "python-infinity": style_blob_with_line_width("Infinity"),
            "lowercase-infinity": style_blob_with_line_width("inf"),
        }

        invalid_extension_list = etree.fromstring(style_blob)
        extension_list = etree.SubElement(
            invalid_extension_list,
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}extLst",
        )
        etree.SubElement(
            extension_list,
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}bogus",
        )
        invalid_extension_list_blob = etree.tostring(
            invalid_extension_list,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        mc_namespace = (
            format_paper_module._MARKUP_COMPATIBILITY_NAMESPACE
        )
        vendor_namespace = "urn:test:invalid-chart-style-vendor"

        def style_root_with_mc_namespaces(*, vendor=False):
            parsed = etree.fromstring(style_blob)
            nsmap = {**parsed.nsmap, "mc": mc_namespace}
            if vendor:
                nsmap["v"] = vendor_namespace
            root = etree.Element(parsed.tag, nsmap=nsmap)
            root.attrib.update(parsed.attrib)
            root.extend(list(parsed))
            return root

        undefined_ignorable_prefix = style_root_with_mc_namespaces()
        undefined_ignorable_prefix.set(
            f"{{{mc_namespace}}}Ignorable",
            "missing",
        )
        undefined_ignorable_prefix_blob = etree.tostring(
            undefined_ignorable_prefix,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        process_content_with_xml_attribute = style_root_with_mc_namespaces(
            vendor=True
        )
        process_content_with_xml_attribute.set(
            f"{{{mc_namespace}}}Ignorable",
            "v",
        )
        category_axis = process_content_with_xml_attribute.find(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}categoryAxis"
        )
        category_axis_index = process_content_with_xml_attribute.index(
            category_axis
        )
        process_content_with_xml_attribute.remove(category_axis)
        invalid_process_wrapper = etree.Element(
            f"{{{vendor_namespace}}}processWrapper"
        )
        invalid_process_wrapper.set(
            f"{{{mc_namespace}}}ProcessContent",
            "v:processWrapper",
        )
        invalid_process_wrapper.set(
            "{http://www.w3.org/XML/1998/namespace}space",
            "preserve",
        )
        invalid_process_wrapper.append(category_axis)
        process_content_with_xml_attribute.insert(
            category_axis_index,
            invalid_process_wrapper,
        )
        process_content_with_xml_attribute_blob = etree.tostring(
            process_content_with_xml_attribute,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        unknown_mc_attribute = style_root_with_mc_namespaces()
        unknown_mc_attribute.set(f"{{{mc_namespace}}}Bogus", "v")
        unknown_mc_attribute_blob = etree.tostring(
            unknown_mc_attribute,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        mc_namespace_as_ignorable = style_root_with_mc_namespaces()
        mc_namespace_as_ignorable.set(
            f"{{{mc_namespace}}}Ignorable",
            "mc",
        )
        mc_namespace_as_ignorable_blob = etree.tostring(
            mc_namespace_as_ignorable,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        invalid_process_qname = style_root_with_mc_namespaces(vendor=True)
        invalid_process_qname.set(
            f"{{{mc_namespace}}}Ignorable",
            "v",
        )
        invalid_process_qname.set(
            f"{{{mc_namespace}}}ProcessContent",
            "v:1bad",
        )
        invalid_process_qname_blob = etree.tostring(
            invalid_process_qname,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        non_xml_whitespace_prefix = style_root_with_mc_namespaces(vendor=True)
        non_xml_whitespace_prefix.set(
            f"{{{mc_namespace}}}Ignorable",
            "v\N{NO-BREAK SPACE}",
        )
        non_xml_whitespace_prefix_blob = etree.tostring(
            non_xml_whitespace_prefix,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        understood_ignorable_attribute = style_root_with_mc_namespaces()
        understood_ignorable_attribute.set(
            f"{{{mc_namespace}}}Ignorable",
            "a",
        )
        understood_ignorable_attribute.set(
            f"{{{format_paper_module._THEME_NAMESPACE}}}bogus",
            "1",
        )
        understood_ignorable_attribute_blob = etree.tostring(
            understood_ignorable_attribute,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        unsupported_must_understand = style_root_with_mc_namespaces(
            vendor=True
        )
        unsupported_must_understand.set(
            f"{{{mc_namespace}}}MustUnderstand",
            "v",
        )
        unsupported_must_understand_blob = etree.tostring(
            unsupported_must_understand,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        def style_blob_with_alternate_content(alternate_content):
            root = style_root_with_mc_namespaces(vendor=True)
            category = root.find(
                f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}categoryAxis"
            )
            category_index = root.index(category)
            root.remove(category)
            alternate_content[0].append(category)
            root.insert(category_index, alternate_content)
            return etree.tostring(
                root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

        missing_requires_alternate = etree.Element(
            f"{{{mc_namespace}}}AlternateContent"
        )
        etree.SubElement(
            missing_requires_alternate,
            f"{{{mc_namespace}}}Choice",
        )
        missing_requires_alternate_blob = style_blob_with_alternate_content(
            missing_requires_alternate
        )

        fallback_only_alternate = etree.Element(
            f"{{{mc_namespace}}}AlternateContent"
        )
        etree.SubElement(
            fallback_only_alternate,
            f"{{{mc_namespace}}}Fallback",
        )
        fallback_only_alternate_blob = style_blob_with_alternate_content(
            fallback_only_alternate
        )

        invalid_unselected_choice = etree.Element(
            f"{{{mc_namespace}}}AlternateContent"
        )
        etree.SubElement(
            invalid_unselected_choice,
            f"{{{mc_namespace}}}Choice",
            Requires="cs",
        )
        unselected_choice = etree.SubElement(
            invalid_unselected_choice,
            f"{{{mc_namespace}}}Choice",
            Requires="v",
        )
        unselected_choice.set(
            "{http://www.w3.org/XML/1998/namespace}space",
            "preserve",
        )
        invalid_unselected_choice_blob = style_blob_with_alternate_content(
            invalid_unselected_choice
        )

        empty_requires_alternate = etree.Element(
            f"{{{mc_namespace}}}AlternateContent"
        )
        etree.SubElement(
            empty_requires_alternate,
            f"{{{mc_namespace}}}Choice",
            Requires="",
        )
        empty_requires_alternate_blob = style_blob_with_alternate_content(
            empty_requires_alternate
        )

        mc_requires_alternate = etree.Element(
            f"{{{mc_namespace}}}AlternateContent"
        )
        etree.SubElement(
            mc_requires_alternate,
            f"{{{mc_namespace}}}Choice",
            Requires="mc",
        )
        mc_requires_alternate_blob = style_blob_with_alternate_content(
            mc_requires_alternate
        )

        unignorable_qualified_attribute = etree.Element(
            f"{{{mc_namespace}}}AlternateContent",
            nsmap={"x": "urn:test:not-ignorable"},
        )
        unignorable_qualified_attribute.set(
            "{urn:test:not-ignorable}attribute",
            "1",
        )
        etree.SubElement(
            unignorable_qualified_attribute,
            f"{{{mc_namespace}}}Choice",
            Requires="cs",
        )
        unignorable_qualified_attribute_blob = (
            style_blob_with_alternate_content(
                unignorable_qualified_attribute
            )
        )

        def style_blob_with_optional_entry_node(node):
            root = etree.fromstring(style_blob)
            entry = root.find(
                f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}axisTitle"
            )
            entry.append(node)
            return etree.tostring(
                root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

        invalid_shape_child = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}spPr"
        )
        etree.SubElement(
            invalid_shape_child,
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}bogus",
        )
        invalid_shape_child_blob = style_blob_with_optional_entry_node(
            invalid_shape_child
        )

        duplicate_shape_fill = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}spPr"
        )
        etree.SubElement(
            duplicate_shape_fill,
            f"{{{format_paper_module._THEME_NAMESPACE}}}noFill",
        )
        etree.SubElement(
            duplicate_shape_fill,
            f"{{{format_paper_module._THEME_NAMESPACE}}}solidFill",
        )
        duplicate_shape_fill_blob = style_blob_with_optional_entry_node(
            duplicate_shape_fill
        )

        reordered_shape_children = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}spPr"
        )
        etree.SubElement(
            reordered_shape_children,
            f"{{{format_paper_module._THEME_NAMESPACE}}}ln",
        )
        etree.SubElement(
            reordered_shape_children,
            f"{{{format_paper_module._THEME_NAMESPACE}}}noFill",
        )
        reordered_shape_children_blob = style_blob_with_optional_entry_node(
            reordered_shape_children
        )

        invalid_shape_attribute = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}spPr",
            bogus="1",
        )
        invalid_shape_attribute_blob = style_blob_with_optional_entry_node(
            invalid_shape_attribute
        )

        invalid_shape_value = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}spPr",
            bwMode="garbage",
        )
        invalid_shape_value_blob = style_blob_with_optional_entry_node(
            invalid_shape_value
        )

        invalid_nested_scheme = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}spPr"
        )
        invalid_solid_fill = etree.SubElement(
            invalid_nested_scheme,
            f"{{{format_paper_module._THEME_NAMESPACE}}}solidFill",
        )
        etree.SubElement(
            invalid_solid_fill,
            f"{{{format_paper_module._THEME_NAMESPACE}}}schemeClr",
            val="notAColor",
        )
        invalid_nested_scheme_blob = style_blob_with_optional_entry_node(
            invalid_nested_scheme
        )

        invalid_nested_dash = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}spPr"
        )
        invalid_dash_line = etree.SubElement(
            invalid_nested_dash,
            f"{{{format_paper_module._THEME_NAMESPACE}}}ln",
        )
        etree.SubElement(
            invalid_dash_line,
            f"{{{format_paper_module._THEME_NAMESPACE}}}prstDash",
            val="bogusDash",
        )
        invalid_nested_dash_blob = style_blob_with_optional_entry_node(
            invalid_nested_dash
        )

        invalid_nested_no_fill = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}spPr"
        )
        etree.SubElement(
            invalid_nested_no_fill,
            f"{{{format_paper_module._THEME_NAMESPACE}}}noFill",
            bogus="1",
        )
        invalid_nested_no_fill_blob = style_blob_with_optional_entry_node(
            invalid_nested_no_fill
        )

        def style_blob_with_shape_children(*children):
            shape = etree.Element(
                f"{{{chart_style_namespace}}}spPr",
            )
            shape.extend(children)
            return style_blob_with_optional_entry_node(shape)

        invalid_group_fill = etree.Element(
            f"{{{drawing_namespace}}}grpFill",
            bogus="1",
        )
        invalid_group_fill_blob = style_blob_with_shape_children(
            invalid_group_fill
        )

        invalid_transform_2d = etree.Element(
            f"{{{drawing_namespace}}}xfrm",
        )
        etree.SubElement(
            invalid_transform_2d,
            f"{{{drawing_namespace}}}off",
            x="27273042316901",
            y="0",
        )
        invalid_transform_2d_blob = style_blob_with_shape_children(
            invalid_transform_2d
        )

        reordered_transform_2d = etree.Element(
            f"{{{drawing_namespace}}}xfrm",
        )
        etree.SubElement(
            reordered_transform_2d,
            f"{{{drawing_namespace}}}ext",
            cx="1",
            cy="1",
        )
        etree.SubElement(
            reordered_transform_2d,
            f"{{{drawing_namespace}}}off",
            x="0",
            y="0",
        )
        reordered_transform_2d_blob = style_blob_with_shape_children(
            reordered_transform_2d
        )

        invalid_preset_geometry = etree.Element(
            f"{{{drawing_namespace}}}prstGeom",
        )
        invalid_preset_geometry_blob = style_blob_with_shape_children(
            invalid_preset_geometry
        )

        invalid_preset_geometry_value = etree.Element(
            f"{{{drawing_namespace}}}prstGeom",
            prst="notAShape",
        )
        invalid_preset_geometry_value_blob = style_blob_with_shape_children(
            invalid_preset_geometry_value
        )

        duplicate_preset_geometry_adjustments = etree.Element(
            f"{{{drawing_namespace}}}prstGeom",
            prst="rect",
        )
        etree.SubElement(
            duplicate_preset_geometry_adjustments,
            f"{{{drawing_namespace}}}avLst",
        )
        etree.SubElement(
            duplicate_preset_geometry_adjustments,
            f"{{{drawing_namespace}}}avLst",
        )
        duplicate_preset_geometry_adjustments_blob = (
            style_blob_with_shape_children(
                duplicate_preset_geometry_adjustments
            )
        )

        invalid_custom_geometry_missing_paths = etree.Element(
            f"{{{drawing_namespace}}}custGeom",
        )
        invalid_custom_geometry_missing_paths_blob = (
            style_blob_with_shape_children(invalid_custom_geometry_missing_paths)
        )

        invalid_custom_geometry_attribute = etree.Element(
            f"{{{drawing_namespace}}}custGeom",
            bogus="1",
        )
        invalid_custom_geometry_attribute_blob = (
            style_blob_with_shape_children(invalid_custom_geometry_attribute)
        )

        duplicate_custom_geometry_paths = etree.Element(
            f"{{{drawing_namespace}}}custGeom",
        )
        etree.SubElement(
            duplicate_custom_geometry_paths,
            f"{{{drawing_namespace}}}pathLst",
        )
        etree.SubElement(
            duplicate_custom_geometry_paths,
            f"{{{drawing_namespace}}}pathLst",
        )
        duplicate_custom_geometry_paths_blob = style_blob_with_shape_children(
            duplicate_custom_geometry_paths
        )

        invalid_custom_geometry_point_count = etree.Element(
            f"{{{drawing_namespace}}}custGeom",
        )
        invalid_point_paths = etree.SubElement(
            invalid_custom_geometry_point_count,
            f"{{{drawing_namespace}}}pathLst",
        )
        invalid_point_path = etree.SubElement(
            invalid_point_paths,
            f"{{{drawing_namespace}}}path",
        )
        invalid_move = etree.SubElement(
            invalid_point_path,
            f"{{{drawing_namespace}}}moveTo",
        )
        etree.SubElement(
            invalid_move,
            f"{{{drawing_namespace}}}pt",
            x="0",
            y="0",
        )
        etree.SubElement(
            invalid_move,
            f"{{{drawing_namespace}}}pt",
            x="1",
            y="1",
        )
        invalid_custom_geometry_point_count_blob = (
            style_blob_with_shape_children(invalid_custom_geometry_point_count)
        )

        invalid_custom_geometry_rect = etree.Element(
            f"{{{drawing_namespace}}}custGeom",
        )
        etree.SubElement(
            invalid_custom_geometry_rect,
            f"{{{drawing_namespace}}}rect",
            l="0",
            t="0",
            r="0",
        )
        etree.SubElement(
            invalid_custom_geometry_rect,
            f"{{{drawing_namespace}}}pathLst",
        )
        invalid_custom_geometry_rect_blob = style_blob_with_shape_children(
            invalid_custom_geometry_rect
        )

        invalid_gradient_stop_count = etree.Element(
            f"{{{drawing_namespace}}}gradFill",
        )
        invalid_gradient_stops = etree.SubElement(
            invalid_gradient_stop_count,
            f"{{{drawing_namespace}}}gsLst",
        )
        invalid_gradient_stop = etree.SubElement(
            invalid_gradient_stops,
            f"{{{drawing_namespace}}}gs",
            pos="0",
        )
        etree.SubElement(
            invalid_gradient_stop,
            f"{{{drawing_namespace}}}schemeClr",
            val="accent1",
        )
        invalid_gradient_stop_count_blob = style_blob_with_shape_children(
            invalid_gradient_stop_count
        )

        invalid_gradient_order = etree.Element(
            f"{{{drawing_namespace}}}gradFill",
        )
        etree.SubElement(
            invalid_gradient_order,
            f"{{{drawing_namespace}}}path",
            path="circle",
        )
        etree.SubElement(
            invalid_gradient_order,
            f"{{{drawing_namespace}}}gsLst",
        )
        invalid_gradient_order_blob = style_blob_with_shape_children(
            invalid_gradient_order
        )

        invalid_gradient_angle = etree.Element(
            f"{{{drawing_namespace}}}gradFill",
        )
        etree.SubElement(
            invalid_gradient_angle,
            f"{{{drawing_namespace}}}lin",
            ang="21600000",
        )
        invalid_gradient_angle_blob = style_blob_with_shape_children(
            invalid_gradient_angle
        )

        invalid_pattern_preset = etree.Element(
            f"{{{drawing_namespace}}}pattFill",
            prst="notAPattern",
        )
        invalid_pattern_preset_blob = style_blob_with_shape_children(
            invalid_pattern_preset
        )

        invalid_pattern_order = etree.Element(
            f"{{{drawing_namespace}}}pattFill",
        )
        invalid_pattern_background = etree.SubElement(
            invalid_pattern_order,
            f"{{{drawing_namespace}}}bgClr",
        )
        etree.SubElement(
            invalid_pattern_background,
            f"{{{drawing_namespace}}}srgbClr",
            val="FFFFFF",
        )
        invalid_pattern_foreground = etree.SubElement(
            invalid_pattern_order,
            f"{{{drawing_namespace}}}fgClr",
        )
        etree.SubElement(
            invalid_pattern_foreground,
            f"{{{drawing_namespace}}}schemeClr",
            val="accent1",
        )
        invalid_pattern_order_blob = style_blob_with_shape_children(
            invalid_pattern_order
        )

        invalid_pattern_color_wrapper = etree.Element(
            f"{{{drawing_namespace}}}pattFill",
        )
        etree.SubElement(
            invalid_pattern_color_wrapper,
            f"{{{drawing_namespace}}}fgClr",
        )
        invalid_pattern_color_wrapper_blob = style_blob_with_shape_children(
            invalid_pattern_color_wrapper
        )

        invalid_blip_dpi = etree.Element(
            f"{{{drawing_namespace}}}blipFill",
            dpi="4294967296",
        )
        invalid_blip_dpi_blob = style_blob_with_shape_children(
            invalid_blip_dpi
        )

        invalid_blip_order = etree.Element(
            f"{{{drawing_namespace}}}blipFill",
        )
        etree.SubElement(
            invalid_blip_order,
            f"{{{drawing_namespace}}}tile",
        )
        etree.SubElement(
            invalid_blip_order,
            f"{{{drawing_namespace}}}srcRect",
        )
        invalid_blip_order_blob = style_blob_with_shape_children(
            invalid_blip_order
        )

        invalid_blip_layout_choice = etree.Element(
            f"{{{drawing_namespace}}}blipFill",
        )
        etree.SubElement(
            invalid_blip_layout_choice,
            f"{{{drawing_namespace}}}tile",
        )
        etree.SubElement(
            invalid_blip_layout_choice,
            f"{{{drawing_namespace}}}stretch",
        )
        invalid_blip_layout_choice_blob = style_blob_with_shape_children(
            invalid_blip_layout_choice
        )

        invalid_blip_tile = etree.Element(
            f"{{{drawing_namespace}}}blipFill",
        )
        etree.SubElement(
            invalid_blip_tile,
            f"{{{drawing_namespace}}}tile",
            tx="27273042316901",
        )
        invalid_blip_tile_blob = style_blob_with_shape_children(
            invalid_blip_tile
        )

        invalid_blip_effect = etree.Element(
            f"{{{drawing_namespace}}}blipFill",
        )
        invalid_blip = etree.SubElement(
            invalid_blip_effect,
            f"{{{drawing_namespace}}}blip",
        )
        etree.SubElement(
            invalid_blip,
            f"{{{drawing_namespace}}}notAnEffect",
        )
        invalid_blip_effect_blob = style_blob_with_shape_children(
            invalid_blip_effect
        )

        invalid_blip_alpha = etree.Element(
            f"{{{drawing_namespace}}}blipFill",
        )
        invalid_alpha_blip = etree.SubElement(
            invalid_blip_alpha,
            f"{{{drawing_namespace}}}blip",
        )
        etree.SubElement(
            invalid_alpha_blip,
            f"{{{drawing_namespace}}}alphaBiLevel",
        )
        invalid_blip_alpha_blob = style_blob_with_shape_children(
            invalid_blip_alpha
        )

        invalid_blip_alpha_mod = etree.Element(
            f"{{{drawing_namespace}}}blipFill",
        )
        invalid_alpha_mod_blip = etree.SubElement(
            invalid_blip_alpha_mod,
            f"{{{drawing_namespace}}}blip",
        )
        etree.SubElement(
            invalid_alpha_mod_blip,
            f"{{{drawing_namespace}}}alphaMod",
        )
        invalid_blip_alpha_mod_blob = style_blob_with_shape_children(
            invalid_blip_alpha_mod
        )

        invalid_blip_grayscl = etree.Element(
            f"{{{drawing_namespace}}}blipFill",
        )
        invalid_grayscl_blip = etree.SubElement(
            invalid_blip_grayscl,
            f"{{{drawing_namespace}}}blip",
        )
        etree.SubElement(
            invalid_grayscl_blip,
            f"{{{drawing_namespace}}}grayscl",
            bogus="1",
        )
        invalid_blip_grayscl_blob = style_blob_with_shape_children(
            invalid_blip_grayscl
        )

        invalid_effect_dag_type = etree.Element(
            f"{{{drawing_namespace}}}effectDag",
        )
        etree.SubElement(
            invalid_effect_dag_type,
            f"{{{drawing_namespace}}}cont",
            type="bogus",
        )
        invalid_effect_dag_type_blob = style_blob_with_shape_children(
            invalid_effect_dag_type
        )

        invalid_effect_dag_child = etree.Element(
            f"{{{drawing_namespace}}}effectDag",
        )
        invalid_effect_dag_container = etree.SubElement(
            invalid_effect_dag_child,
            f"{{{drawing_namespace}}}cont",
        )
        etree.SubElement(
            invalid_effect_dag_container,
            f"{{{drawing_namespace}}}notAnEffect",
        )
        invalid_effect_dag_child_blob = style_blob_with_shape_children(
            invalid_effect_dag_child
        )

        invalid_effect_dag_namespace = etree.Element(
            f"{{{drawing_namespace}}}effectDag",
        )
        invalid_namespace_container = etree.SubElement(
            invalid_effect_dag_namespace,
            f"{{{drawing_namespace}}}cont",
        )
        etree.SubElement(
            invalid_namespace_container,
            "{urn:test:foreign-effect}blur",
        )
        invalid_effect_dag_namespace_blob = style_blob_with_shape_children(
            invalid_effect_dag_namespace
        )

        invalid_custom_dash_value = etree.Element(
            f"{{{drawing_namespace}}}ln",
        )
        invalid_custom_dash = etree.SubElement(
            invalid_custom_dash_value,
            f"{{{drawing_namespace}}}custDash",
        )
        etree.SubElement(
            invalid_custom_dash,
            f"{{{drawing_namespace}}}ds",
            d="0",
            sp="1",
        )
        invalid_custom_dash_value_blob = style_blob_with_shape_children(
            invalid_custom_dash_value
        )

        invalid_miter_limit = etree.Element(
            f"{{{drawing_namespace}}}ln",
        )
        etree.SubElement(
            invalid_miter_limit,
            f"{{{drawing_namespace}}}miter",
            lim="-1",
        )
        invalid_miter_limit_blob = style_blob_with_shape_children(
            invalid_miter_limit
        )

        invalid_line_end_type = etree.Element(
            f"{{{drawing_namespace}}}ln",
        )
        etree.SubElement(
            invalid_line_end_type,
            f"{{{drawing_namespace}}}headEnd",
            type="arrowhead",
        )
        invalid_line_end_type_blob = style_blob_with_shape_children(
            invalid_line_end_type
        )

        invalid_effect_list_child = etree.Element(
            f"{{{drawing_namespace}}}effectLst",
        )
        etree.SubElement(
            invalid_effect_list_child,
            f"{{{drawing_namespace}}}unknownEffect",
        )
        invalid_effect_list_child_blob = style_blob_with_shape_children(
            invalid_effect_list_child
        )

        duplicate_effect_list_shadow = etree.Element(
            f"{{{drawing_namespace}}}effectLst",
        )
        for _ in range(2):
            shadow = etree.SubElement(
                duplicate_effect_list_shadow,
                f"{{{drawing_namespace}}}outerShdw",
            )
            etree.SubElement(
                shadow,
                f"{{{drawing_namespace}}}srgbClr",
                val="000000",
            )
        duplicate_effect_list_shadow_blob = style_blob_with_shape_children(
            duplicate_effect_list_shadow
        )

        reordered_effect_list = etree.Element(
            f"{{{drawing_namespace}}}effectLst",
        )
        shadow = etree.SubElement(
            reordered_effect_list,
            f"{{{drawing_namespace}}}outerShdw",
        )
        etree.SubElement(
            shadow,
            f"{{{drawing_namespace}}}srgbClr",
            val="000000",
        )
        etree.SubElement(
            reordered_effect_list,
            f"{{{drawing_namespace}}}blur",
            rad="1",
        )
        reordered_effect_list_blob = style_blob_with_shape_children(
            reordered_effect_list
        )

        invalid_outer_shadow_color = etree.Element(
            f"{{{drawing_namespace}}}effectLst",
        )
        etree.SubElement(
            invalid_outer_shadow_color,
            f"{{{drawing_namespace}}}outerShdw",
            dir="21600000",
        )
        invalid_outer_shadow_color_blob = style_blob_with_shape_children(
            invalid_outer_shadow_color
        )

        invalid_effect_blur = etree.Element(
            f"{{{drawing_namespace}}}effectLst",
        )
        etree.SubElement(
            invalid_effect_blur,
            f"{{{drawing_namespace}}}blur",
            rad="2147483648",
        )
        invalid_effect_blur_blob = style_blob_with_shape_children(
            invalid_effect_blur
        )

        def make_minimal_scene(*, camera_preset="orthographicFront"):
            scene = etree.Element(f"{{{drawing_namespace}}}scene3d")
            etree.SubElement(
                scene,
                f"{{{drawing_namespace}}}camera",
                prst=camera_preset,
            )
            etree.SubElement(
                scene,
                f"{{{drawing_namespace}}}lightRig",
                rig="balanced",
                dir="t",
            )
            return scene

        invalid_scene_missing_light = etree.Element(
            f"{{{drawing_namespace}}}scene3d"
        )
        etree.SubElement(
            invalid_scene_missing_light,
            f"{{{drawing_namespace}}}camera",
            prst="orthographicFront",
        )
        invalid_scene_missing_light_blob = style_blob_with_shape_children(
            invalid_scene_missing_light
        )

        invalid_scene_camera = make_minimal_scene(
            camera_preset="notACamera"
        )
        invalid_scene_camera_blob = style_blob_with_shape_children(
            invalid_scene_camera
        )

        invalid_scene_rotation = make_minimal_scene()
        etree.SubElement(
            invalid_scene_rotation[0],
            f"{{{drawing_namespace}}}rot",
            lat="21600000",
            lon="0",
            rev="0",
        )
        invalid_scene_rotation_blob = style_blob_with_shape_children(
            invalid_scene_rotation
        )

        invalid_scene_backdrop = make_minimal_scene()
        invalid_backdrop = etree.SubElement(
            invalid_scene_backdrop,
            f"{{{drawing_namespace}}}backdrop",
        )
        etree.SubElement(
            invalid_backdrop,
            f"{{{drawing_namespace}}}anchor",
            x="0",
            y="0",
            z="0",
        )
        etree.SubElement(
            invalid_backdrop,
            f"{{{drawing_namespace}}}norm",
            dx="0",
            dy="0",
            dz="0",
        )
        invalid_scene_backdrop_blob = style_blob_with_shape_children(
            invalid_scene_backdrop
        )

        invalid_shape3d_material = etree.Element(
            f"{{{drawing_namespace}}}sp3d",
            prstMaterial="notAMaterial",
        )
        invalid_shape3d_material_blob = style_blob_with_shape_children(
            invalid_shape3d_material
        )

        invalid_shape3d_bevel = etree.Element(
            f"{{{drawing_namespace}}}sp3d",
        )
        etree.SubElement(
            invalid_shape3d_bevel,
            f"{{{drawing_namespace}}}bevelT",
            prst="notABevel",
        )
        invalid_shape3d_bevel_blob = style_blob_with_shape_children(
            invalid_shape3d_bevel
        )

        invalid_shape3d_color = etree.Element(
            f"{{{drawing_namespace}}}sp3d",
        )
        etree.SubElement(
            invalid_shape3d_color,
            f"{{{drawing_namespace}}}extrusionClr",
        )
        invalid_shape3d_color_blob = style_blob_with_shape_children(
            invalid_shape3d_color
        )

        invalid_character_child = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}defRPr"
        )
        etree.SubElement(
            invalid_character_child,
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}bogus",
        )
        invalid_character_child_blob = style_blob_with_optional_entry_node(
            invalid_character_child
        )

        invalid_character_boolean = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}defRPr",
            b="notBoolean",
        )
        invalid_character_boolean_blob = style_blob_with_optional_entry_node(
            invalid_character_boolean
        )

        invalid_character_size = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}defRPr",
            sz="99",
        )
        invalid_character_size_blob = style_blob_with_optional_entry_node(
            invalid_character_size
        )

        invalid_character_highlight = etree.Element(
            f"{{{chart_style_namespace}}}defRPr"
        )
        etree.SubElement(
            invalid_character_highlight,
            f"{{{drawing_namespace}}}highlight",
        )
        invalid_character_highlight_blob = style_blob_with_optional_entry_node(
            invalid_character_highlight
        )

        invalid_character_underline_fill = etree.Element(
            f"{{{chart_style_namespace}}}defRPr"
        )
        etree.SubElement(
            invalid_character_underline_fill,
            f"{{{drawing_namespace}}}uFill",
        )
        invalid_character_underline_fill_blob = (
            style_blob_with_optional_entry_node(invalid_character_underline_fill)
        )

        invalid_character_font = etree.Element(
            f"{{{chart_style_namespace}}}defRPr"
        )
        etree.SubElement(
            invalid_character_font,
            f"{{{drawing_namespace}}}latin",
            panose="0011",
        )
        invalid_character_font_blob = style_blob_with_optional_entry_node(
            invalid_character_font
        )

        invalid_character_rtl = etree.Element(
            f"{{{chart_style_namespace}}}defRPr"
        )
        etree.SubElement(
            invalid_character_rtl,
            f"{{{drawing_namespace}}}rtl",
            val="yes",
        )
        invalid_character_rtl_blob = style_blob_with_optional_entry_node(
            invalid_character_rtl
        )

        invalid_character_hyperlink = etree.Element(
            f"{{{chart_style_namespace}}}defRPr"
        )
        invalid_hyperlink = etree.SubElement(
            invalid_character_hyperlink,
            f"{{{drawing_namespace}}}hlinkClick",
        )
        invalid_hyperlink.set(
            f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}id",
            "rIdMalformedHyperlink",
        )
        etree.SubElement(
            invalid_hyperlink,
            f"{{{drawing_namespace}}}snd",
        )
        invalid_character_hyperlink_blob = style_blob_with_optional_entry_node(
            invalid_character_hyperlink
        )

        duplicate_body_autofit = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}bodyPr"
        )
        etree.SubElement(
            duplicate_body_autofit,
            f"{{{format_paper_module._THEME_NAMESPACE}}}noAutofit",
        )
        etree.SubElement(
            duplicate_body_autofit,
            f"{{{format_paper_module._THEME_NAMESPACE}}}normAutofit",
        )
        duplicate_body_autofit_blob = style_blob_with_optional_entry_node(
            duplicate_body_autofit
        )

        invalid_body_column_count = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}bodyPr",
            numCol="999",
        )
        invalid_body_column_count_blob = style_blob_with_optional_entry_node(
            invalid_body_column_count
        )

        invalid_body_wrap = etree.Element(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}bodyPr",
            wrap="garbage",
        )
        invalid_body_wrap_blob = style_blob_with_optional_entry_node(
            invalid_body_wrap
        )

        invalid_body_no_autofit = etree.Element(
            f"{{{chart_style_namespace}}}bodyPr"
        )
        etree.SubElement(
            invalid_body_no_autofit,
            f"{{{drawing_namespace}}}noAutofit",
            bogus="1",
        )
        invalid_body_no_autofit_blob = style_blob_with_optional_entry_node(
            invalid_body_no_autofit
        )

        invalid_body_normal_autofit = etree.Element(
            f"{{{chart_style_namespace}}}bodyPr"
        )
        etree.SubElement(
            invalid_body_normal_autofit,
            f"{{{drawing_namespace}}}normAutofit",
            fontScale="999",
        )
        invalid_body_normal_autofit_blob = (
            style_blob_with_optional_entry_node(invalid_body_normal_autofit)
        )

        invalid_body_flat_text = etree.Element(
            f"{{{chart_style_namespace}}}bodyPr"
        )
        etree.SubElement(
            invalid_body_flat_text,
            f"{{{drawing_namespace}}}flatTx",
            z="27273042317000",
        )
        invalid_body_flat_text_blob = style_blob_with_optional_entry_node(
            invalid_body_flat_text
        )

        invalid_text_warp_preset = etree.Element(
            f"{{{chart_style_namespace}}}bodyPr"
        )
        etree.SubElement(
            invalid_text_warp_preset,
            f"{{{drawing_namespace}}}prstTxWarp",
            prst="notATextShape",
        )
        invalid_text_warp_preset_blob = style_blob_with_optional_entry_node(
            invalid_text_warp_preset
        )

        invalid_text_warp_guide = etree.Element(
            f"{{{chart_style_namespace}}}bodyPr"
        )
        invalid_warp = etree.SubElement(
            invalid_text_warp_guide,
            f"{{{drawing_namespace}}}prstTxWarp",
            prst="textPlain",
        )
        invalid_guides = etree.SubElement(
            invalid_warp,
            f"{{{drawing_namespace}}}avLst",
        )
        etree.SubElement(
            invalid_guides,
            f"{{{drawing_namespace}}}gd",
            name="adj",
        )
        invalid_text_warp_guide_blob = style_blob_with_optional_entry_node(
            invalid_text_warp_guide
        )

        invalid_text_warp_guide_token = etree.Element(
            f"{{{chart_style_namespace}}}bodyPr"
        )
        invalid_token_warp = etree.SubElement(
            invalid_text_warp_guide_token,
            f"{{{drawing_namespace}}}prstTxWarp",
            prst="textPlain",
        )
        invalid_token_guides = etree.SubElement(
            invalid_token_warp,
            f"{{{drawing_namespace}}}avLst",
        )
        etree.SubElement(
            invalid_token_guides,
            f"{{{drawing_namespace}}}gd",
            name="a  b",
            fmla="val 1",
        )
        invalid_text_warp_guide_token_blob = (
            style_blob_with_optional_entry_node(
                invalid_text_warp_guide_token
            )
        )

        empty_entry_modifiers = etree.fromstring(style_blob)
        empty_entry_modifiers.find(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}axisTitle"
        ).set("mods", "")
        empty_entry_modifiers_blob = etree.tostring(
            empty_entry_modifiers,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        empty_reference_modifiers = etree.fromstring(style_blob)
        empty_reference_modifiers.find(
            f".//{{{format_paper_module._CHART_STYLE_NAMESPACE}}}lnRef"
        ).set("mods", " \t\r\n")
        empty_reference_modifiers_blob = etree.tostring(
            empty_reference_modifiers,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        whitespace_font_index = etree.fromstring(style_blob)
        whitespace_font_index.find(
            f".//{{{format_paper_module._CHART_STYLE_NAMESPACE}}}fontRef"
        ).set("idx", " minor ")
        whitespace_font_index_blob = etree.tostring(
            whitespace_font_index,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        whitespace_marker_symbol = etree.fromstring(style_blob)
        marker_anchor = whitespace_marker_symbol.find(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}dataPointMarker"
        )
        marker_anchor.addnext(
            etree.Element(
                f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}"
                "dataPointMarkerLayout",
                symbol=" circle ",
            )
        )
        whitespace_marker_symbol_blob = etree.tostring(
            whitespace_marker_symbol,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        missing_style_reference = etree.fromstring(style_blob)
        axis_title = missing_style_reference.find(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}axisTitle"
        )
        axis_title.remove(
            axis_title.find(
                f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}fontRef"
            )
        )
        missing_style_reference_blob = etree.tostring(
            missing_style_reference,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        reordered_style_reference = etree.fromstring(style_blob)
        axis_title = reordered_style_reference.find(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}axisTitle"
        )
        fill_reference = axis_title.find(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}fillRef"
        )
        effect_reference = axis_title.find(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}effectRef"
        )
        fill_reference.addprevious(effect_reference)
        reordered_style_reference_blob = etree.tostring(
            reordered_style_reference,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        invalid_marker_layout = etree.fromstring(style_blob)
        marker_anchor = invalid_marker_layout.find(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}dataPointMarker"
        )
        marker_anchor.addnext(
            etree.Element(
                f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}dataPointMarkerLayout",
                symbol="hexagon",
                size="1",
            )
        )
        invalid_marker_layout_blob = etree.tostring(
            invalid_marker_layout,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        invalid_color_variation = etree.fromstring(color_blob)
        variation = etree.SubElement(
            invalid_color_variation,
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}variation",
        )
        etree.SubElement(
            variation,
            f"{{{format_paper_module._THEME_NAMESPACE}}}notATransform",
        )
        invalid_color_variation_blob = etree.tostring(
            invalid_color_variation,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        out_of_range_color_transform = etree.fromstring(color_blob)
        etree.SubElement(
            out_of_range_color_transform[0],
            f"{{{format_paper_module._THEME_NAMESPACE}}}tint",
            val="100001",
        )
        out_of_range_color_transform_blob = etree.tostring(
            out_of_range_color_transform,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        cases = (
            (
                "wrong-style-root",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                wrong_style_root_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "wrong-style-content-type",
                True,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                style_blob,
                "application/octet-stream",
                False,
            ),
            (
                "missing-required-style-entry",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                incomplete_style_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "missing-color-method",
                True,
                format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                missing_color_method_blob,
                format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-scrgb",
                False,
                format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                invalid_scrgb_blob,
                format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-hsl",
                False,
                format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                invalid_hsl_blob,
                format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-system-color",
                False,
                format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                invalid_system_color_blob,
                format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-scheme-color",
                False,
                format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                invalid_scheme_color_blob,
                format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-preset-color",
                False,
                format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                invalid_preset_color_blob,
                format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-legacy-color-index",
                False,
                format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                invalid_legacy_color_blob,
                format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                False,
            ),
            *(
                (
                    f"invalid-line-width-{case_name}",
                    False,
                    format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                    blob,
                    format_paper_module._CHART_STYLE_CONTENT_TYPE,
                    False,
                )
                for case_name, blob in invalid_line_width_blobs.items()
            ),
            (
                "invalid-extension-list",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_extension_list_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "undefined-ignorable-prefix",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                undefined_ignorable_prefix_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "process-content-with-xml-attribute",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                process_content_with_xml_attribute_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "unknown-mc-attribute",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                unknown_mc_attribute_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "mc-namespace-as-ignorable",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                mc_namespace_as_ignorable_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-process-qname",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_process_qname_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "non-xml-whitespace-prefix",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                non_xml_whitespace_prefix_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "understood-ignorable-attribute",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                understood_ignorable_attribute_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "unsupported-must-understand",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                unsupported_must_understand_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "alternate-content-missing-requires",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                missing_requires_alternate_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "alternate-content-without-choice",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                fallback_only_alternate_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "alternate-content-unselected-choice-xml-space",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_unselected_choice_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "alternate-content-empty-requires",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                empty_requires_alternate_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "alternate-content-mc-requires",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                mc_requires_alternate_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "alternate-content-unignorable-qualified-attribute",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                unignorable_qualified_attribute_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-shape-child",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_shape_child_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "duplicate-shape-fill",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                duplicate_shape_fill_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "reordered-shape-children",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                reordered_shape_children_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-shape-attribute",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_shape_attribute_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-shape-value",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_shape_value_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-nested-scheme-color",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_nested_scheme_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-nested-preset-dash",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_nested_dash_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-nested-no-fill",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_nested_no_fill_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-group-fill",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_group_fill_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-transform-2d",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_transform_2d_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "reordered-transform-2d",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                reordered_transform_2d_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-preset-geometry",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_preset_geometry_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-preset-geometry-value",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_preset_geometry_value_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "duplicate-preset-geometry-adjustments",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                duplicate_preset_geometry_adjustments_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-custom-geometry-missing-paths",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_custom_geometry_missing_paths_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-custom-geometry-attribute",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_custom_geometry_attribute_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "duplicate-custom-geometry-paths",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                duplicate_custom_geometry_paths_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-custom-geometry-point-count",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_custom_geometry_point_count_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-custom-geometry-rect",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_custom_geometry_rect_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-gradient-stop-count",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_gradient_stop_count_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-gradient-order",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_gradient_order_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-gradient-angle",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_gradient_angle_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-pattern-preset",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_pattern_preset_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-pattern-order",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_pattern_order_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-pattern-color-wrapper",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_pattern_color_wrapper_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-blip-dpi",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_blip_dpi_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-blip-order",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_blip_order_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-blip-layout-choice",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_blip_layout_choice_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-blip-tile",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_blip_tile_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-blip-effect",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_blip_effect_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-blip-alpha",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_blip_alpha_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-blip-alpha-mod",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_blip_alpha_mod_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-blip-grayscl",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_blip_grayscl_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-effect-dag-type",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_effect_dag_type_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-effect-dag-child",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_effect_dag_child_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-effect-dag-namespace",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_effect_dag_namespace_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-custom-dash-value",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_custom_dash_value_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-miter-limit",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_miter_limit_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-line-end-type",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_line_end_type_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-effect-list-child",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_effect_list_child_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "duplicate-effect-list-shadow",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                duplicate_effect_list_shadow_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "reordered-effect-list",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                reordered_effect_list_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-outer-shadow-color",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_outer_shadow_color_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-effect-blur",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_effect_blur_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-scene-missing-light",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_scene_missing_light_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-scene-camera",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_scene_camera_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-scene-rotation",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_scene_rotation_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-scene-backdrop",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_scene_backdrop_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-shape3d-material",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_shape3d_material_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-shape3d-bevel",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_shape3d_bevel_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-shape3d-color",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_shape3d_color_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-character-child",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_character_child_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-character-boolean",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_character_boolean_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-character-size",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_character_size_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-character-highlight",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_character_highlight_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-character-underline-fill",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_character_underline_fill_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-character-font",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_character_font_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-character-rtl",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_character_rtl_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-character-hyperlink",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_character_hyperlink_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "duplicate-body-autofit",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                duplicate_body_autofit_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-body-column-count",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_body_column_count_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-body-wrap",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_body_wrap_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-body-no-autofit",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_body_no_autofit_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-body-normal-autofit",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_body_normal_autofit_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-body-flat-text",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_body_flat_text_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-text-warp-preset",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_text_warp_preset_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-text-warp-guide",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_text_warp_guide_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-text-warp-guide-token",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_text_warp_guide_token_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "empty-entry-modifiers",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                empty_entry_modifiers_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "empty-reference-modifiers",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                empty_reference_modifiers_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "whitespace-font-index",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                whitespace_font_index_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "whitespace-marker-symbol",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                whitespace_marker_symbol_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "missing-style-reference",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                missing_style_reference_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "reordered-style-reference",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                reordered_style_reference_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-marker-layout",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                invalid_marker_layout_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "invalid-color-variation",
                True,
                format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                invalid_color_variation_blob,
                format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "out-of-range-color-transform",
                False,
                format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                out_of_range_color_transform_blob,
                format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                False,
            ),
            (
                "external-chart-style",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                style_blob,
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                True,
            ),
            (
                "external-chartex-color-style",
                True,
                format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                color_blob,
                format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                True,
            ),
            (
                "relationship-content-type-mismatch",
                False,
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                color_blob,
                format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                False,
            ),
        )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            for (
                case_name,
                extended,
                relationship_type,
                blob,
                content_type,
                external,
            ) in cases:
                with self.subTest(case=case_name):
                    first_path = tmp / f"{case_name}-first.docx"
                    second_path = tmp / f"{case_name}-second.docx"
                    output_path = tmp / f"{case_name}-merged.docx"

                    first = Document()
                    self.add_test_chart(
                        first,
                        case_name,
                        extended=extended,
                        add_sidecar=True,
                        sidecar_specs=(
                            (
                                relationship_type,
                                blob,
                                content_type,
                                external,
                            ),
                        ),
                    )
                    first.save(first_path)

                    second = Document()
                    second.add_paragraph("BODY")
                    second.save(second_path)

                    with self.assertRaises(DocumentConcatError):
                        concatenate_documents(
                            first_path,
                            second_path,
                            output_path,
                            restart_body_page_number=False,
                        )

    def test_chart_style_hyperlinks_require_relationship_ids(self):
        chart_style_namespace = format_paper_module._CHART_STYLE_NAMESPACE
        drawing_namespace = format_paper_module._THEME_NAMESPACE
        for hyperlink_name in ("hlinkClick", "hlinkMouseOver"):
            with self.subTest(hyperlink=hyperlink_name):
                style_root = etree.fromstring(
                    self.make_test_chart_style_blob()
                )
                axis_title = style_root.find(
                    f"{{{chart_style_namespace}}}axisTitle"
                )
                character_properties = etree.SubElement(
                    axis_title,
                    f"{{{chart_style_namespace}}}defRPr",
                )
                etree.SubElement(
                    character_properties,
                    f"{{{drawing_namespace}}}{hyperlink_name}",
                )

                with self.assertRaisesRegex(
                    ValueError,
                    "missing relationship ID",
                ):
                    format_paper_module._validate_chart_style_root(
                        style_root,
                        "style",
                    )

    def test_drawing_hyperlinks_require_relationship_id_attributes(self):
        document = Document()
        part = format_paper_module.Part(
            format_paper_module.PackURI(
                "/word/charts/hyperlink-validation.xml"
            ),
            "application/xml",
            b"<root/>",
            document.part.package,
        )
        drawing_namespace = format_paper_module._THEME_NAMESPACE
        relationship_id = (
            f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}id"
        )
        for hyperlink_name in (
            "hlinkClick",
            "hlinkHover",
            "hlinkMouseOver",
        ):
            with self.subTest(hyperlink=hyperlink_name):
                hyperlink = etree.Element(
                    f"{{{drawing_namespace}}}{hyperlink_name}"
                )
                with self.assertRaisesRegex(
                    ValueError,
                    "hyperlink.*missing relationship ID",
                ):
                    format_paper_module._relationship_references(
                        hyperlink,
                        part,
                        "DrawingML",
                    )
                hyperlink.set(relationship_id, "")
                self.assertEqual(
                    format_paper_module._relationship_references(
                        hyperlink,
                        part,
                        "DrawingML",
                    ),
                    set(),
                )

    def test_drawing_media_and_attribution_require_relationship_attributes(self):
        drawing_namespace = format_paper_module._THEME_NAMESPACE
        required_tags = (
            f"{{{drawing_namespace}}}snd",
            f"{{{drawing_namespace}}}wavAudioFile",
            f"{{{drawing_namespace}}}audioFile",
            f"{{{drawing_namespace}}}videoFile",
            f"{{{drawing_namespace}}}quickTimeFile",
            (
                f"{{{format_paper_module._DRAWING_2016_11_NAMESPACE}}}"
                "picAttrSrcUrl"
            ),
            (
                f"{{{format_paper_module._DRAWING_2017_MODEL3D_NAMESPACE}}}"
                "attrSrcUrl"
            ),
        )
        document = Document()
        part = format_paper_module.Part(
            format_paper_module.PackURI(
                "/word/charts/required-relationship-attributes.xml"
            ),
            "application/xml",
            b"<root/>",
            document.part.package,
        )
        for tag in required_tags:
            with self.subTest(tag=tag):
                with self.assertRaisesRegex(
                    ValueError,
                    "required relationship attribute",
                ):
                    format_paper_module._relationship_references(
                        etree.Element(tag),
                        part,
                        "DrawingML",
                    )

        composer = format_paper_module._create_structure_preserving_composer(
            Document()
        )
        with self.assertRaisesRegex(
            ValueError,
            "required relationship attribute",
        ):
            composer.add_images(
                document,
                etree.Element(f"{{{drawing_namespace}}}audioFile"),
            )

        wrong_sound = etree.Element(f"{{{drawing_namespace}}}snd")
        wrong_sound.set(
            format_paper_module._RELATIONSHIP_LINK_ATTRIBUTE,
            "rIdWrongSound",
        )
        part.rels.add_relationship(
            format_paper_module.RT.AUDIO,
            "https://example.test/wrong-sound-link",
            "rIdWrongSound",
            is_external=True,
        )
        with self.assertRaisesRegex(
            ValueError,
            "required relationship attribute",
        ):
            format_paper_module._relationship_references(
                wrong_sound,
                part,
                "DrawingML",
            )

    def test_relationship_references_validate_image_layer_and_model_content(self):
        cases = (
            (
                "valid-image-layer",
                f"{{{format_paper_module._DRAWING_2010_NAMESPACE}}}imgLayer",
                format_paper_module.RT.IMAGE,
                "image/png",
                True,
            ),
            (
                "wrong-image-layer-type",
                f"{{{format_paper_module._DRAWING_2010_NAMESPACE}}}imgLayer",
                format_paper_module.RT.PACKAGE,
                "application/octet-stream",
                False,
            ),
            (
                "valid-model",
                (
                    f"{{{format_paper_module._DRAWING_2017_MODEL3D_NAMESPACE}}}"
                    "model3d"
                ),
                format_paper_module._MODEL3D_RELATIONSHIP_TYPE,
                format_paper_module._MODEL3D_CONTENT_TYPE,
                True,
            ),
            (
                "wrong-model-content-type",
                (
                    f"{{{format_paper_module._DRAWING_2017_MODEL3D_NAMESPACE}}}"
                    "model3d"
                ),
                format_paper_module._MODEL3D_RELATIONSHIP_TYPE,
                "application/octet-stream",
                False,
            ),
        )
        for case_name, tag, relationship_type, content_type, valid in cases:
            with self.subTest(case=case_name):
                document = Document()
                source_part = format_paper_module.Part(
                    format_paper_module.PackURI(
                        f"/word/charts/{case_name}-source.xml"
                    ),
                    "application/xml",
                    b"<root/>",
                    document.part.package,
                )
                target_part = format_paper_module.Part(
                    format_paper_module.PackURI(
                        f"/word/media/{case_name}.bin"
                    ),
                    content_type,
                    b"relationship target",
                    document.part.package,
                )
                source_part.rels.add_relationship(
                    relationship_type,
                    target_part,
                    "rIdTarget",
                )
                node = etree.Element(tag)
                node.set(
                    format_paper_module._RELATIONSHIP_EMBED_ATTRIBUTE,
                    "rIdTarget",
                )
                if valid:
                    self.assertEqual(
                        format_paper_module._relationship_references(
                            node,
                            source_part,
                            "DrawingML",
                        ),
                        {"rIdTarget"},
                    )
                else:
                    with self.assertRaisesRegex(ValueError, "wrong target"):
                        format_paper_module._relationship_references(
                            node,
                            source_part,
                            "DrawingML",
                        )

        source_document = Document()
        wrong_model_part = format_paper_module.Part(
            format_paper_module.PackURI(
                "/word/media/wrong-body-model.bin"
            ),
            "application/octet-stream",
            b"not a glTF binary model",
            source_document.part.package,
        )
        source_document.part.rels.add_relationship(
            format_paper_module._MODEL3D_RELATIONSHIP_TYPE,
            wrong_model_part,
            "rIdWrongBodyModel",
        )
        model = etree.Element(
            f"{{{format_paper_module._DRAWING_2017_MODEL3D_NAMESPACE}}}model3d"
        )
        model.set(
            format_paper_module._RELATIONSHIP_EMBED_ATTRIBUTE,
            "rIdWrongBodyModel",
        )
        composer = format_paper_module._create_structure_preserving_composer(
            Document()
        )
        with self.assertRaisesRegex(ValueError, "3D model.*invalid relationship"):
            composer.add_images(source_document, model)

    def test_chart_relationship_roles_validate_typed_targets(self):
        chart_namespace = format_paper_module._CHART_NAMESPACE
        relationship_id = (
            f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}id"
        )
        for content_type, valid in (
            (format_paper_module.CT.DML_CHARTSHAPES, True),
            ("application/octet-stream", False),
        ):
            case_name = "valid" if valid else "wrong-content-type"
            with self.subTest(user_shapes=case_name):
                document = Document()
                source_part = format_paper_module.Part(
                    format_paper_module.PackURI(
                        f"/word/charts/user-shapes-{case_name}.xml"
                    ),
                    format_paper_module.CT.DML_CHART,
                    b"<root/>",
                    document.part.package,
                )
                target_part = format_paper_module.Part(
                    format_paper_module.PackURI(
                        f"/word/drawings/user-shapes-{case_name}.xml"
                    ),
                    content_type,
                    b"user shapes",
                    document.part.package,
                )
                source_part.rels.add_relationship(
                    format_paper_module.RT.CHART_USER_SHAPES,
                    target_part,
                    "rIdUserShapes",
                )
                root = etree.Element(f"{{{chart_namespace}}}chartSpace")
                etree.SubElement(
                    root,
                    f"{{{chart_namespace}}}userShapes",
                ).set(relationship_id, "rIdUserShapes")
                if valid:
                    format_paper_module._validate_chart_relationship_roles(
                        root,
                        source_part,
                        format_paper_module.RT.CHART,
                    )
                else:
                    with self.assertRaisesRegex(
                        ValueError,
                        "user shapes.*invalid relationship",
                    ):
                        format_paper_module._validate_chart_relationship_roles(
                            root,
                            source_part,
                            format_paper_module.RT.CHART,
                        )

        external_data_cases = (
            ("internal-package", format_paper_module.RT.PACKAGE, False, True),
            ("external-package", format_paper_module.RT.PACKAGE, True, False),
            (
                "external-link",
                format_paper_module.RT.EXTERNAL_LINK,
                True,
                False,
            ),
        )
        for case_name, relationship_type, external, valid in external_data_cases:
            with self.subTest(external_data=case_name):
                document = Document()
                source_part = format_paper_module.Part(
                    format_paper_module.PackURI(
                        f"/word/charts/external-data-{case_name}.xml"
                    ),
                    format_paper_module.CT.DML_CHART,
                    b"<root/>",
                    document.part.package,
                )
                target = (
                    "https://example.test/chart-data"
                    if external
                    else format_paper_module.Part(
                        format_paper_module.PackURI(
                            "/word/embeddings/chart-data.xlsx"
                        ),
                        (
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        b"embedded workbook",
                        document.part.package,
                    )
                )
                source_part.rels.add_relationship(
                    relationship_type,
                    target,
                    "rIdExternalData",
                    is_external=external,
                )
                root = etree.Element(f"{{{chart_namespace}}}chartSpace")
                etree.SubElement(
                    root,
                    f"{{{chart_namespace}}}externalData",
                ).set(relationship_id, "rIdExternalData")
                if valid:
                    format_paper_module._validate_chart_relationship_roles(
                        root,
                        source_part,
                        format_paper_module.RT.CHART,
                    )
                else:
                    with self.assertRaisesRegex(
                        ValueError,
                        "external data.*wrong relationship type",
                    ):
                        format_paper_module._validate_chart_relationship_roles(
                            root,
                            source_part,
                            format_paper_module.RT.CHART,
                        )

    def test_chart_style_luminance_transform_boundaries(self):
        drawing_namespace = format_paper_module._THEME_NAMESPACE
        for value, valid in (
            ("-2147483648", True),
            ("100000", True),
            ("100001", False),
        ):
            with self.subTest(value=value):
                color_root = etree.fromstring(
                    self.make_test_chart_color_style_blob()
                )
                etree.SubElement(
                    color_root[0],
                    f"{{{drawing_namespace}}}lum",
                    val=value,
                )
                if valid:
                    format_paper_module._validate_chart_style_root(
                        color_root,
                        "color",
                    )
                else:
                    with self.assertRaisesRegex(
                        ValueError,
                        "out-of-range color transform",
                    ):
                        format_paper_module._validate_chart_style_root(
                            color_root,
                            "color",
                        )

    def test_chart_style_mouse_over_rejects_wrong_relationship_type(self):
        style_root = etree.fromstring(self.make_test_chart_style_blob())
        chart_style_namespace = format_paper_module._CHART_STYLE_NAMESPACE
        drawing_namespace = format_paper_module._THEME_NAMESPACE
        relationship_namespace = format_paper_module._RELATIONSHIP_NAMESPACE
        axis_title = style_root.find(
            f"{{{chart_style_namespace}}}axisTitle"
        )
        character_properties = etree.SubElement(
            axis_title,
            f"{{{chart_style_namespace}}}defRPr",
        )
        mouse_over = etree.SubElement(
            character_properties,
            f"{{{drawing_namespace}}}hlinkMouseOver",
        )
        mouse_over.set(
            f"{{{relationship_namespace}}}id",
            "rIdWrongHyperlink",
        )

        document = Document()
        style_part = format_paper_module.Part(
            format_paper_module.PackURI(
                "/word/charts/chartStyleMouseOver.xml"
            ),
            format_paper_module._CHART_STYLE_CONTENT_TYPE,
            etree.tostring(
                style_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            ),
            document.part.package,
        )
        style_part.rels.add_relationship(
            format_paper_module.RT.IMAGE,
            "https://example.test/not-a-hyperlink",
            "rIdWrongHyperlink",
            is_external=True,
        )

        parsed_style = etree.fromstring(style_part.blob)
        format_paper_module._validate_chart_style_root(parsed_style, "style")
        with self.assertRaisesRegex(ValueError, "hyperlink.*wrong target"):
            format_paper_module._relationship_references(
                parsed_style,
                style_part,
                "chart style Part",
            )

    def test_concatenate_documents_preserves_multiple_chart_style_sidecars(self):
        def blob_with_id(blob, identifier):
            root = etree.fromstring(blob)
            root.set("id", str(identifier))
            return etree.tostring(
                root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

        style_with_marker = etree.fromstring(
            blob_with_id(self.make_test_chart_style_blob(), 202)
        )
        style_with_marker.find(
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}dataPointMarker"
        ).addnext(
            etree.Element(
                f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}dataPointMarkerLayout",
                symbol="circle",
                size="7",
            )
        )
        style_blobs = (
            blob_with_id(self.make_test_chart_style_blob(), 201),
            etree.tostring(
                style_with_marker,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            ),
        )

        color_with_variation = etree.fromstring(
            blob_with_id(self.make_test_chart_color_style_blob(), 11)
        )
        variation = etree.SubElement(
            color_with_variation,
            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}variation",
        )
        etree.SubElement(
            variation,
            f"{{{format_paper_module._THEME_NAMESPACE}}}lumMod",
            val="80000",
        )
        etree.SubElement(
            variation,
            f"{{{format_paper_module._THEME_NAMESPACE}}}lumOff",
            val="20000",
        )
        color_blobs = (
            blob_with_id(self.make_test_chart_color_style_blob(), 10),
            etree.tostring(
                color_with_variation,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            ),
        )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "multiple-sidecars-first.docx"
            second_path = tmp / "multiple-sidecars-second.docx"
            output_path = tmp / "multiple-sidecars-merged.docx"

            first = Document()
            self.add_test_chart(
                first,
                "MULTIPLE SIDECARS",
                sidecar_specs=tuple(
                    (
                        format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                        blob,
                        format_paper_module._CHART_STYLE_CONTENT_TYPE,
                        False,
                    )
                    for blob in style_blobs
                )
                + tuple(
                    (
                        format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
                        blob,
                        format_paper_module._CHART_COLOR_STYLE_CONTENT_TYPE,
                        False,
                    )
                    for blob in color_blobs
                ),
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph("BODY")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            chart_part = next(
                relationship.target_part
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.CHART
            )
            copied_style_blobs = sorted(
                relationship.target_part.blob
                for relationship in chart_part.rels.values()
                if relationship.reltype
                == format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE
            )
            copied_color_blobs = sorted(
                relationship.target_part.blob
                for relationship in chart_part.rels.values()
                if relationship.reltype
                == format_paper_module._CHART_COLOR_STYLE_RELATIONSHIP_TYPE
            )
            self.assertEqual(copied_style_blobs, sorted(style_blobs))
            self.assertEqual(copied_color_blobs, sorted(color_blobs))

    def test_concatenate_documents_copies_only_referenced_chart_style_relationships(self):
        tiny_png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC"
            "AAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            for referenced in (True, False):
                with self.subTest(referenced=referenced):
                    case_name = "referenced" if referenced else "unreferenced"
                    first_path = tmp / f"{case_name}-sidecar-rel-first.docx"
                    second_path = tmp / f"{case_name}-sidecar-rel-second.docx"
                    output_path = tmp / f"{case_name}-sidecar-rel-merged.docx"

                    first = Document()
                    chart_part = self.add_test_chart(
                        first,
                        "SIDECAR RELATIONSHIP",
                        add_sidecar=True,
                    )
                    style_part = next(
                        relationship.target_part
                        for relationship in chart_part.rels.values()
                        if relationship.reltype
                        == format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE
                    )
                    if referenced:
                        style_root = etree.fromstring(style_part.blob)
                        style_entry = style_root.find(
                            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}axisTitle"
                        )
                        shape_properties = etree.SubElement(
                            style_entry,
                            f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}spPr",
                        )
                        blip_fill = etree.SubElement(
                            shape_properties,
                            f"{{{format_paper_module._THEME_NAMESPACE}}}blipFill",
                        )
                        blip = etree.SubElement(
                            blip_fill,
                            f"{{{format_paper_module._THEME_NAMESPACE}}}blip",
                        )
                        blip.set(
                            f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}embed",
                            "rIdStyleImage",
                        )
                        stretch = etree.SubElement(
                            blip_fill,
                            f"{{{format_paper_module._THEME_NAMESPACE}}}stretch",
                        )
                        etree.SubElement(
                            stretch,
                            f"{{{format_paper_module._THEME_NAMESPACE}}}fillRect",
                        )
                        style_part._blob = etree.tostring(
                            style_root,
                            xml_declaration=True,
                            encoding="UTF-8",
                            standalone=True,
                        )

                    image_part = format_paper_module.Part(
                        format_paper_module.PackURI(
                            f"/word/media/{case_name}-chart-style.png"
                        ),
                        "image/png",
                        tiny_png,
                        first.part.package,
                    )
                    style_part.rels.add_relationship(
                        format_paper_module.RT.IMAGE,
                        image_part,
                        "rIdStyleImage",
                    )
                    first.save(first_path)

                    second = Document()
                    second.add_paragraph("BODY")
                    second.save(second_path)

                    if referenced:
                        result = concatenate_documents(
                            first_path,
                            second_path,
                            output_path,
                            restart_body_page_number=False,
                        )
                        self.assertTrue(result["concatenated"])
                        merged = Document(output_path)
                        copied_chart = next(
                            relationship.target_part
                            for relationship in merged.part.rels.values()
                            if relationship.reltype
                            == format_paper_module.RT.CHART
                        )
                        copied_style = next(
                            relationship.target_part
                            for relationship in copied_chart.rels.values()
                            if relationship.reltype
                            == format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE
                        )
                        copied_image_relationship = copied_style.rels[
                            "rIdStyleImage"
                        ]
                        self.assertFalse(copied_image_relationship.is_external)
                        self.assertEqual(
                            copied_image_relationship.reltype,
                            format_paper_module.RT.IMAGE,
                        )
                        self.assertEqual(
                            copied_image_relationship.target_part.blob,
                            tiny_png,
                        )
                    else:
                        with self.assertRaises(DocumentConcatError):
                            concatenate_documents(
                                first_path,
                                second_path,
                                output_path,
                                restart_body_page_number=False,
                            )

    def test_concatenate_documents_preserves_chart_style_external_hyperlink(self):
        hyperlink_url = "https://example.test/chart-style-link"
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "style-hyperlink-first.docx"
            second_path = tmp / "style-hyperlink-second.docx"
            output_path = tmp / "style-hyperlink-merged.docx"

            first = Document()
            chart_part = self.add_test_chart(
                first,
                "SIDECAR HYPERLINK",
                add_sidecar=True,
            )
            style_part = next(
                relationship.target_part
                for relationship in chart_part.rels.values()
                if relationship.reltype
                == format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE
            )
            style_root = etree.fromstring(style_part.blob)
            axis_title = style_root.find(
                f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}axisTitle"
            )
            character_properties = etree.SubElement(
                axis_title,
                f"{{{format_paper_module._CHART_STYLE_NAMESPACE}}}defRPr",
            )
            hyperlink = etree.SubElement(
                character_properties,
                f"{{{format_paper_module._THEME_NAMESPACE}}}hlinkClick",
            )
            hyperlink.set(
                format_paper_module._RELATIONSHIP_ID_ATTRIBUTE,
                "rIdStyleHyperlink",
            )
            style_part._blob = etree.tostring(
                style_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            style_part.rels.add_relationship(
                format_paper_module.RT.HYPERLINK,
                hyperlink_url,
                "rIdStyleHyperlink",
                is_external=True,
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph("BODY")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            copied_chart = next(
                relationship.target_part
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.CHART
            )
            copied_style = next(
                relationship.target_part
                for relationship in copied_chart.rels.values()
                if relationship.reltype
                == format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE
            )
            copied_hyperlink = copied_style.rels["rIdStyleHyperlink"]
            self.assertTrue(copied_hyperlink.is_external)
            self.assertEqual(
                copied_hyperlink.reltype,
                format_paper_module.RT.HYPERLINK,
            )
            self.assertEqual(copied_hyperlink.target_ref, hyperlink_url)

    def test_concatenate_documents_rejects_chart_style_from_non_chart_parent(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            for external in (False, True):
                with self.subTest(external=external):
                    case_name = "external" if external else "internal"
                    first_path = tmp / f"{case_name}-parent-first.docx"
                    second_path = tmp / f"{case_name}-parent-second.docx"
                    output_path = tmp / f"{case_name}-parent-merged.docx"

                    first = Document()
                    if external:
                        first.part.rels.add_relationship(
                            format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                            "https://example.test/not-a-chart-parent.xml",
                            "rIdWrongParent",
                            is_external=True,
                        )
                    else:
                        sidecar_part = format_paper_module.Part(
                            format_paper_module.PackURI(
                                "/word/charts/wrong-parent-style.xml"
                            ),
                            format_paper_module._CHART_STYLE_CONTENT_TYPE,
                            self.make_test_chart_style_blob(),
                            first.part.package,
                        )
                        first.part.rels.add_relationship(
                            format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                            sidecar_part,
                            "rIdWrongParent",
                        )
                    alt_chunk = OxmlElement("w:altChunk")
                    alt_chunk.set(qn("r:id"), "rIdWrongParent")
                    first.element.body.insert(
                        len(first.element.body) - 1,
                        alt_chunk,
                    )
                    first.save(first_path)

                    second = Document()
                    second.add_paragraph("BODY")
                    second.save(second_path)

                    with self.assertRaises(DocumentConcatError):
                        concatenate_documents(
                            first_path,
                            second_path,
                            output_path,
                            restart_body_page_number=False,
                        )

    def test_composer_revalidates_special_parts_in_target_package(self):
        master = Document()
        composer = format_paper_module._create_structure_preserving_composer(
            master
        )
        cases = (
            (
                "chart-wrong-relationship",
                format_paper_module.CT.DML_CHART,
                b"not parsed because the relationship role is rejected first",
                "urn:test:not-a-chart-relationship",
                False,
            ),
            (
                "style-wrong-parent",
                format_paper_module._CHART_STYLE_CONTENT_TYPE,
                self.make_test_chart_style_blob(),
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                False,
            ),
            (
                "style-wrong-content-type",
                "application/octet-stream",
                self.make_test_chart_style_blob(),
                format_paper_module._CHART_STYLE_RELATIONSHIP_TYPE,
                True,
            ),
            (
                "diagram-wrong-relationship",
                format_paper_module.CT.DML_DIAGRAM_DATA,
                b"not parsed because the relationship role is rejected first",
                "urn:test:not-a-diagram-relationship",
                False,
            ),
        )
        for (
            case_name,
            content_type,
            blob,
            relationship_type,
            allow_chart_style,
        ) in cases:
            with self.subTest(case=case_name):
                part = format_paper_module.Part(
                    format_paper_module.PackURI(
                        f"/word/charts/{case_name}.xml"
                    ),
                    content_type,
                    blob,
                    master.part.package,
                )
                with self.assertRaises(ValueError):
                    composer._copy_relationship_target_part(
                        part,
                        relationship_type,
                        allow_chart_style=allow_chart_style,
                    )

    def test_concatenate_documents_rejects_invalid_chart_theme_overrides(self):
        drawing_namespace = format_paper_module._THEME_NAMESPACE

        def override_with_root_attribute(document):
            root = etree.fromstring(
                self.make_test_theme_override_blob(document)
            )
            root.set("bad", "1")
            return (
                (
                    etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    ),
                    format_paper_module.CT.OFC_THEME_OVERRIDE,
                    False,
                ),
            )

        def override_with_invalid_rgb(document):
            root = etree.fromstring(
                self.make_test_theme_override_blob(document)
            )
            color = root.find(
                ".//a:clrScheme/a:accent1/*",
                namespaces={"a": drawing_namespace},
            )
            color.tag = f"{{{drawing_namespace}}}srgbClr"
            color.attrib.clear()
            color.set("val", "NOT-RGB")
            return (
                (
                    etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    ),
                    format_paper_module.CT.OFC_THEME_OVERRIDE,
                    False,
                ),
            )

        cases = {
            "external": lambda document: (
                (b"", format_paper_module.CT.OFC_THEME_OVERRIDE, True),
            ),
            "wrong-content-type": lambda document: (
                (
                    self.make_test_theme_override_blob(document),
                    "text/plain",
                    False,
                ),
            ),
            "wrong-root": lambda document: (
                (
                    (
                        f'<a:theme xmlns:a="{drawing_namespace}"/>'
                    ).encode("utf-8"),
                    format_paper_module.CT.OFC_THEME_OVERRIDE,
                    False,
                ),
            ),
            "malformed": lambda document: (
                (
                    b"not XML",
                    format_paper_module.CT.OFC_THEME_OVERRIDE,
                    False,
                ),
            ),
            "root-attribute": override_with_root_attribute,
            "invalid-rgb": override_with_invalid_rgb,
            "empty-schemes": lambda document: (
                (
                    (
                        f'<a:themeOverride xmlns:a="{drawing_namespace}">'
                        '<a:clrScheme/><a:fontScheme/><a:fmtScheme/>'
                        '</a:themeOverride>'
                    ).encode("utf-8"),
                    format_paper_module.CT.OFC_THEME_OVERRIDE,
                    False,
                ),
            ),
            "duplicate": lambda document: (
                (
                    self.make_test_theme_override_blob(document),
                    format_paper_module.CT.OFC_THEME_OVERRIDE,
                    False,
                ),
                (
                    self.make_test_theme_override_blob(document),
                    format_paper_module.CT.OFC_THEME_OVERRIDE,
                    False,
                ),
            ),
        }

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            for case_name, override_factory in cases.items():
                with self.subTest(case=case_name):
                    first_path = tmp / f"invalid-{case_name}-first.docx"
                    second_path = tmp / f"invalid-{case_name}-second.docx"
                    output_path = tmp / f"invalid-{case_name}-merged.docx"

                    first = Document()
                    self.customize_test_theme(
                        first,
                        accent1="D01020",
                        major_latin="CoverChartFont",
                        format_name="CoverChartFormat",
                        background_mapping="accent2",
                    )
                    self.add_test_chart(
                        first,
                        case_name,
                        override_specs=override_factory(first),
                    )
                    first.save(first_path)

                    second = Document()
                    self.customize_test_theme(
                        second,
                        accent1="1020D0",
                        major_latin="BodyChartFont",
                        format_name="BodyChartFormat",
                        background_mapping="accent1",
                    )
                    second.add_paragraph("BODY")
                    second.save(second_path)

                    with self.assertRaises(DocumentConcatError):
                        concatenate_documents(
                            first_path,
                            second_path,
                            output_path,
                            restart_body_page_number=False,
                        )

    def test_concatenate_documents_rejects_unreferenced_theme_override_relationship(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "unreferenced-override-rel-first.docx"
            second_path = tmp / "unreferenced-override-rel-second.docx"
            output_path = tmp / "unreferenced-override-rel-merged.docx"

            first = Document()
            chart_part = self.add_test_chart(
                first,
                "UNREFERENCED OVERRIDE RELATIONSHIP",
                override_specs=(
                    (
                        self.make_test_theme_override_blob(first),
                        format_paper_module.CT.OFC_THEME_OVERRIDE,
                        False,
                    ),
                ),
            )
            override_part = next(
                relationship.target_part
                for relationship in chart_part.rels.values()
                if relationship.reltype
                == format_paper_module.RT.THEME_OVERRIDE
            )
            unused_part = format_paper_module.Part(
                format_paper_module.PackURI(
                    "/word/theme/unused-override-sidecar.bin"
                ),
                "application/octet-stream",
                b"unused",
                first.part.package,
            )
            override_part.rels.add_relationship(
                "urn:test:unused-theme-override-relationship",
                unused_part,
                "rIdUnused",
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph("BODY")
            second.save(second_path)

            with self.assertRaises(DocumentConcatError):
                concatenate_documents(
                    first_path,
                    second_path,
                    output_path,
                    restart_body_page_number=False,
                )

    def test_concatenate_documents_rejects_unresolved_chart_relationship_references(self):
        cases = (
            ("chart-missing-plot-area", False, "missing-plot-area"),
            ("chartex-empty-chart-data", True, "empty-chart-data"),
            ("chart-external-data", False, "relationship-attribute"),
            (
                "chart-external-data-wrong-type",
                False,
                "wrong-external-data-type",
            ),
            ("chart-external-embed", False, "external-embed"),
            ("chart-wrong-embed-mime", False, "wrong-embed-mime"),
            ("chart-wrong-link-type", False, "wrong-link-type"),
            ("chart-wrong-link-mime", False, "wrong-link-mime"),
            ("chart-wrong-hyperlink-type", False, "wrong-hyperlink-type"),
            ("chart-wrong-media-type", False, "wrong-media-type"),
            ("chart-wrong-model-type", False, "wrong-model-type"),
            ("chartex-missing-fallback", True, "missing-fallback"),
            ("chartex-wrong-fallback", True, "wrong-fallback"),
            ("chartex-wrong-fallback-mime", True, "wrong-fallback-mime"),
        )
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            for case_name, extended, mutation in cases:
                with self.subTest(case=case_name):
                    first_path = tmp / f"{case_name}-first.docx"
                    second_path = tmp / f"{case_name}-second.docx"
                    output_path = tmp / f"{case_name}-merged.docx"

                    first = Document()
                    chart_part = self.add_test_chart(
                        first,
                        case_name,
                        extended=extended,
                    )
                    chart_root = etree.fromstring(chart_part.blob)
                    if mutation == "missing-plot-area":
                        chart = chart_root.find(
                            f"{{{format_paper_module._CHART_NAMESPACE}}}chart"
                        )
                        chart.remove(
                            chart.find(
                                f"{{{format_paper_module._CHART_NAMESPACE}}}plotArea"
                            )
                        )
                    elif mutation == "empty-chart-data":
                        chart_data = chart_root.find(
                            f"{{{format_paper_module._CHARTEX_NAMESPACE}}}chartData"
                        )
                        chart_data.clear()
                    elif mutation == "relationship-attribute":
                        external_data = etree.SubElement(
                            chart_root,
                            f"{{{format_paper_module._CHART_NAMESPACE}}}externalData",
                        )
                        external_data.set(
                            f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}id",
                            "rIdMissing",
                        )
                    elif mutation == "wrong-external-data-type":
                        external_data = etree.SubElement(
                            chart_root,
                            f"{{{format_paper_module._CHART_NAMESPACE}}}externalData",
                        )
                        external_data.set(
                            f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}id",
                            "rIdExternalData",
                        )
                        wrong_target = format_paper_module.Part(
                            format_paper_module.PackURI(
                                "/word/media/not-chart-data.png"
                            ),
                            "image/png",
                            base64.b64decode(
                                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC"
                                "AAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
                            ),
                            first.part.package,
                        )
                        chart_part.rels.add_relationship(
                            format_paper_module.RT.IMAGE,
                            wrong_target,
                            "rIdExternalData",
                        )
                    elif mutation in {
                        "external-embed",
                        "wrong-embed-mime",
                        "wrong-link-type",
                        "wrong-link-mime",
                    }:
                        shape_properties = chart_root.find(
                            f"{{{format_paper_module._CHART_NAMESPACE}}}spPr"
                        )
                        blip_fill = etree.SubElement(
                            shape_properties,
                            f"{{{format_paper_module._THEME_NAMESPACE}}}blipFill",
                        )
                        blip = etree.SubElement(
                            blip_fill,
                            f"{{{format_paper_module._THEME_NAMESPACE}}}blip",
                        )
                        blip.set(
                            f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}{'link' if mutation.startswith('wrong-link') else 'embed'}",
                            "rIdImage",
                        )
                        if mutation == "external-embed":
                            chart_part.rels.add_relationship(
                                format_paper_module.RT.IMAGE,
                                "https://example.test/chart-image.png",
                                "rIdImage",
                                is_external=True,
                            )
                        elif mutation == "wrong-link-type":
                            chart_part.rels.add_relationship(
                                format_paper_module.RT.HYPERLINK,
                                "https://example.test/not-an-image",
                                "rIdImage",
                                is_external=True,
                            )
                        else:
                            embedded_part = format_paper_module.Part(
                                format_paper_module.PackURI(
                                    "/word/charts/not-an-embedded-image.bin"
                                ),
                                "application/octet-stream",
                                b"not an image",
                                first.part.package,
                            )
                            chart_part.rels.add_relationship(
                                format_paper_module.RT.IMAGE,
                                embedded_part,
                                "rIdImage",
                            )
                    elif mutation in {
                        "wrong-hyperlink-type",
                        "wrong-media-type",
                        "wrong-model-type",
                    }:
                        shape_properties = chart_root.find(
                            f"{{{format_paper_module._CHART_NAMESPACE}}}spPr"
                        )
                        if mutation == "wrong-hyperlink-type":
                            relationship_node = etree.SubElement(
                                shape_properties,
                                f"{{{format_paper_module._THEME_NAMESPACE}}}hlinkClick",
                            )
                            relationship_node.set(
                                f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}id",
                                "rIdTypedTarget",
                            )
                            expected_wrong_type = format_paper_module.RT.IMAGE
                        elif mutation == "wrong-media-type":
                            relationship_node = etree.SubElement(
                                shape_properties,
                                f"{{{format_paper_module._THEME_NAMESPACE}}}audioFile",
                            )
                            relationship_node.set(
                                f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}link",
                                "rIdTypedTarget",
                            )
                            expected_wrong_type = format_paper_module.RT.VIDEO
                        else:
                            relationship_node = etree.SubElement(
                                shape_properties,
                                f"{{{format_paper_module._DRAWING_2017_MODEL3D_NAMESPACE}}}model3d",
                            )
                            relationship_node.set(
                                f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}embed",
                                "rIdTypedTarget",
                            )
                            expected_wrong_type = format_paper_module.RT.IMAGE
                        wrong_target = format_paper_module.Part(
                            format_paper_module.PackURI(
                                f"/word/charts/{mutation}.bin"
                            ),
                            "application/octet-stream",
                            b"wrong typed relationship payload",
                            first.part.package,
                        )
                        chart_part.rels.add_relationship(
                            expected_wrong_type,
                            wrong_target,
                            "rIdTypedTarget",
                        )
                    else:
                        chart_root.set("fallbackImg", "rIdFallback")
                        if mutation in {
                            "wrong-fallback",
                            "wrong-fallback-mime",
                        }:
                            fallback_part = format_paper_module.Part(
                                format_paper_module.PackURI(
                                    "/word/charts/not-an-image.bin"
                                ),
                                "application/octet-stream",
                                b"not an image",
                                first.part.package,
                            )
                            chart_part.rels.add_relationship(
                                (
                                    format_paper_module.RT.IMAGE
                                    if mutation == "wrong-fallback-mime"
                                    else "urn:test:not-an-image"
                                ),
                                fallback_part,
                                "rIdFallback",
                            )
                    chart_part._blob = etree.tostring(
                        chart_root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                    first.save(first_path)

                    second = Document()
                    second.add_paragraph("BODY")
                    second.save(second_path)

                    with self.assertRaises(DocumentConcatError):
                        concatenate_documents(
                            first_path,
                            second_path,
                            output_path,
                            restart_body_page_number=False,
                        )

    def test_concatenate_documents_accepts_extended_override_and_chart_fallback_image(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "extended-override-first.docx"
            second_path = tmp / "extended-override-second.docx"
            output_path = tmp / "extended-override-merged.docx"

            first = Document()
            override_root = etree.fromstring(
                self.make_test_theme_override_blob(first)
            )
            format_scheme = override_root.find(
                f"{{{format_paper_module._THEME_NAMESPACE}}}fmtScheme"
            )
            format_scheme.attrib.pop("name", None)
            fill_style_list = format_scheme.find(
                f"{{{format_paper_module._THEME_NAMESPACE}}}fillStyleLst"
            )
            fill_style_list.append(
                etree.fromstring(etree.tostring(fill_style_list[0]))
            )
            override_blob = etree.tostring(
                override_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            chart_part = self.add_test_chart(
                first,
                "EXTENDED OVERRIDE AND FALLBACK",
                extended=True,
                override_specs=(
                    (
                        override_blob,
                        format_paper_module.CT.OFC_THEME_OVERRIDE,
                        False,
                    ),
                ),
            )
            chart_root = etree.fromstring(chart_part.blob)
            chart_root.set("fallbackImg", "rIdFallback")
            shape_properties = chart_root.find(
                f"{{{format_paper_module._CHARTEX_NAMESPACE}}}spPr"
            )
            existing_fill = shape_properties[0]
            blip_fill = etree.Element(
                f"{{{format_paper_module._THEME_NAMESPACE}}}blipFill"
            )
            linked_blip = etree.SubElement(
                blip_fill,
                f"{{{format_paper_module._THEME_NAMESPACE}}}blip",
            )
            linked_blip.set(
                f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}link",
                "rIdLinkedImage",
            )
            stretch = etree.SubElement(
                blip_fill,
                f"{{{format_paper_module._THEME_NAMESPACE}}}stretch",
            )
            etree.SubElement(
                stretch,
                f"{{{format_paper_module._THEME_NAMESPACE}}}fillRect",
            )
            shape_properties.replace(existing_fill, blip_fill)
            chart_part._blob = etree.tostring(
                chart_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            fallback_part = format_paper_module.Part(
                format_paper_module.PackURI(
                    "/word/media/chart-fallback.png"
                ),
                "image/png",
                base64.b64decode(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC"
                    "AAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
                ),
                first.part.package,
            )
            chart_part.rels.add_relationship(
                format_paper_module.RT.IMAGE,
                fallback_part,
                "rIdFallback",
            )
            chart_part.rels.add_relationship(
                format_paper_module.RT.IMAGE,
                "https://example.test/chart-linked-image.png",
                "rIdLinkedImage",
                is_external=True,
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph("BODY")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            copied_chart = next(
                relationship.target_part
                for relationship in merged.part.rels.values()
                if relationship.reltype
                == format_paper_module._CHARTEX_RELATIONSHIP_TYPE
            )
            copied_fallback = copied_chart.rels["rIdFallback"]
            self.assertEqual(copied_fallback.reltype, format_paper_module.RT.IMAGE)
            self.assertEqual(
                copied_fallback.target_part.content_type,
                "image/png",
            )
            copied_link = copied_chart.rels["rIdLinkedImage"]
            self.assertTrue(copied_link.is_external)
            self.assertEqual(
                copied_link.reltype,
                format_paper_module.RT.IMAGE,
            )
            self.assertEqual(
                copied_link.target_ref,
                "https://example.test/chart-linked-image.png",
            )
            copied_override = next(
                relationship.target_part
                for relationship in copied_chart.rels.values()
                if relationship.reltype
                == format_paper_module.RT.THEME_OVERRIDE
            )
            self.assertEqual(copied_override.blob, override_blob)

    @staticmethod
    def make_theme_relationship_graph_builder():
        class CountingRelationships(dict):
            def __init__(self):
                super().__init__()
                self.values_call_count = 0

            def values(self):
                self.values_call_count += 1
                return super().values()

        class TestPart:
            def __init__(self, label):
                self.content_type = f"application/x-test-{label}"
                self.blob = label.encode()
                self.rels = CountingRelationships()

        class TestRelationship:
            def __init__(self, relationship_id, target):
                self.rId = relationship_id
                self.reltype = "urn:test:theme-relationship"
                self.is_external = False
                self.target_part = target

        def make_part(label):
            return TestPart(label)

        def connect(source, target, relationship_id):
            source.rels[relationship_id] = TestRelationship(
                relationship_id,
                target,
            )

        return make_part, connect

    def test_theme_relationship_signature_memoizes_shared_dag_subgraphs(self):
        make_part, connect = self.make_theme_relationship_graph_builder()
        root = make_part("root")
        left = make_part("left")
        right = make_part("right")
        shared = make_part("shared")
        leaf = make_part("leaf")
        connect(root, left, "rIdLeft")
        connect(root, right, "rIdRight")
        connect(left, shared, "rIdShared")
        connect(right, shared, "rIdShared")
        connect(shared, leaf, "rIdLeaf")

        signature = format_paper_module._theme_relationship_signature(root)

        self.assertEqual(len(signature), 2)
        self.assertEqual(root.rels.values_call_count, 1)
        self.assertEqual(left.rels.values_call_count, 1)
        self.assertEqual(right.rels.values_call_count, 1)
        self.assertEqual(shared.rels.values_call_count, 1)
        self.assertEqual(leaf.rels.values_call_count, 1)

    def test_theme_relationship_signature_preserves_cycles_depth_and_budgets(self):
        make_part, connect = self.make_theme_relationship_graph_builder()
        first = make_part("cycle-first")
        second = make_part("cycle-second")
        connect(first, second, "rIdSecond")
        connect(second, first, "rIdFirst")

        cycle_signature = format_paper_module._theme_relationship_signature(
            first
        )
        self.assertEqual(cycle_signature[0][5][0][5], ("cycle", 2))

        boundary_cycle = [
            make_part(f"boundary-cycle-{index}") for index in range(64)
        ]
        for index in range(len(boundary_cycle) - 1):
            connect(
                boundary_cycle[index],
                boundary_cycle[index + 1],
                f"rId{index}",
            )
        connect(boundary_cycle[-1], boundary_cycle[0], "rIdCycle")
        boundary_signature = (
            format_paper_module._theme_relationship_signature(
                boundary_cycle[0]
            )
        )
        for _ in boundary_cycle:
            boundary_signature = boundary_signature[0][5]
        self.assertEqual(boundary_signature, ("cycle", 64))

        depth_root = make_part("depth-root")
        cached_leaf = make_part("cached-leaf")
        depth_chain = [
            make_part(f"depth-{index}") for index in range(63)
        ]
        connect(depth_root, cached_leaf, "rIdDirect")
        connect(depth_root, depth_chain[0], "rIdDeep")
        for index in range(len(depth_chain) - 1):
            connect(
                depth_chain[index],
                depth_chain[index + 1],
                f"rId{index}",
            )
        connect(depth_chain[-1], cached_leaf, "rIdLeaf")
        with self.assertRaisesRegex(ValueError, "graph is too deep"):
            format_paper_module._theme_relationship_signature(depth_root)

        node_chain = [make_part(f"node-{index}") for index in range(3)]
        connect(node_chain[0], node_chain[1], "rId1")
        connect(node_chain[1], node_chain[2], "rId2")
        with patch.object(
            format_paper_module,
            "MAX_THEME_RELATIONSHIP_GRAPH_NODES",
            2,
        ):
            with self.assertRaisesRegex(ValueError, "too many parts"):
                format_paper_module._theme_relationship_signature(
                    node_chain[0]
                )

        edge_root = make_part("edge-root")
        connect(edge_root, make_part("edge-left"), "rIdLeft")
        connect(edge_root, make_part("edge-right"), "rIdRight")
        with patch.object(
            format_paper_module,
            "MAX_THEME_RELATIONSHIP_GRAPH_EDGES",
            1,
        ):
            with self.assertRaisesRegex(ValueError, "too many relationships"):
                format_paper_module._theme_relationship_signature(edge_root)

        expansion_root = make_part("expansion-root")
        expansion_first = make_part("expansion-first")
        expansion_second = make_part("expansion-second")
        connect(expansion_root, expansion_first, "rIdFirstVisit")
        connect(expansion_root, expansion_first, "rIdSecondVisit")
        connect(expansion_first, expansion_second, "rIdSecond")
        connect(expansion_second, expansion_first, "rIdFirst")
        with patch.object(
            format_paper_module,
            "MAX_THEME_RELATIONSHIP_GRAPH_EDGES",
            4,
        ):
            format_paper_module._theme_relationship_signature(expansion_root)
        with patch.object(
            format_paper_module,
            "MAX_THEME_RELATIONSHIP_GRAPH_EXPANSIONS",
            4,
        ):
            with self.assertRaisesRegex(ValueError, "too many expansions"):
                format_paper_module._theme_relationship_signature(
                    expansion_root
                )

    def test_concatenate_documents_rejects_chart_theme_relationship_payload_drift(self):
        cover_theme_image = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC"
            "AAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )
        body_theme_image = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC"
            "AAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        )

        def attach_theme_image(document, payload):
            theme_part = document.part.part_related_by(
                format_paper_module.RT.THEME
            )
            theme_root = etree.fromstring(theme_part.blob)
            namespace = format_paper_module._THEME_NAMESPACE
            fill_style_list = theme_root.find(
                ".//a:fmtScheme/a:fillStyleLst",
                namespaces={"a": namespace},
            )
            replacement = etree.Element(f"{{{namespace}}}blipFill")
            blip = etree.SubElement(replacement, f"{{{namespace}}}blip")
            blip.set(
                f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}embed",
                "rIdThemeImage",
            )
            stretch = etree.SubElement(
                replacement,
                f"{{{namespace}}}stretch",
            )
            etree.SubElement(stretch, f"{{{namespace}}}fillRect")
            fill_style_list.replace(fill_style_list[0], replacement)
            theme_part._blob = etree.tostring(
                theme_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

            image_part = format_paper_module.Part(
                format_paper_module.PackURI(
                    "/word/media/theme-background.png"
                ),
                "image/png",
                payload,
                document.part.package,
            )
            theme_part.rels.add_relationship(
                format_paper_module.RT.IMAGE,
                image_part,
                "rIdThemeImage",
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "theme-rel-payload-first.docx"
            second_path = tmp / "theme-rel-payload-second.docx"
            output_path = tmp / "theme-rel-payload-merged.docx"

            first = Document()
            attach_theme_image(first, cover_theme_image)
            self.add_test_chart(first, "THEME RELATIONSHIP PAYLOAD")
            first.save(first_path)

            second = Document()
            attach_theme_image(second, body_theme_image)
            second.add_paragraph("BODY")
            second.save(second_path)

            with self.assertRaises(DocumentConcatError):
                concatenate_documents(
                    first_path,
                    second_path,
                    output_path,
                    restart_body_page_number=False,
                )

            override_first_path = tmp / "theme-rel-override-first.docx"
            override_second_path = tmp / "theme-rel-override-second.docx"
            override_output_path = tmp / "theme-rel-override-merged.docx"
            override_first = Document()
            complete_override = self.make_test_theme_override_blob(
                override_first
            )
            attach_theme_image(override_first, cover_theme_image)
            self.add_test_chart(
                override_first,
                "AUTHORITATIVE THEME OVERRIDE",
                override_specs=(
                    (
                        complete_override,
                        format_paper_module.CT.OFC_THEME_OVERRIDE,
                        False,
                    ),
                ),
            )
            override_first.save(override_first_path)

            override_second = Document()
            attach_theme_image(override_second, body_theme_image)
            override_second.add_paragraph("BODY")
            override_second.save(override_second_path)

            result = concatenate_documents(
                override_first_path,
                override_second_path,
                override_output_path,
                restart_body_page_number=False,
            )
            self.assertTrue(result["concatenated"])
            merged = Document(override_output_path)
            copied_chart = next(
                relationship.target_part
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.CHART
            )
            copied_override = next(
                relationship.target_part
                for relationship in copied_chart.rels.values()
                if relationship.reltype
                == format_paper_module.RT.THEME_OVERRIDE
            )
            self.assertEqual(copied_override.blob, complete_override)

    def test_concatenate_documents_uses_chart_theme_fast_path_and_requires_authority(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)

            first_path = tmp / "same-chart-theme-first.docx"
            second_path = tmp / "same-chart-theme-second.docx"
            output_path = tmp / "same-chart-theme-merged.docx"
            first = Document()
            self.customize_test_theme(
                first,
                accent1="335577",
                major_latin="SharedChartFont",
                format_name="SharedChartFormat",
                background_mapping="accent1",
            )
            source_chart = self.add_test_chart(first, "SAME CHART THEME")
            embedded_workbook = format_paper_module.Part(
                format_paper_module.PackURI(
                    "/word/embeddings/chart-data.xlsx"
                ),
                (
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                b"test workbook payload",
                first.part.package,
            )
            source_chart.rels.add_relationship(
                format_paper_module.RT.PACKAGE,
                embedded_workbook,
                "rIdExternalData",
            )
            source_chart_root = etree.fromstring(source_chart.blob)
            external_data = etree.SubElement(
                source_chart_root,
                f"{{{format_paper_module._CHART_NAMESPACE}}}externalData",
            )
            external_data.set(
                f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}id",
                "rIdExternalData",
            )
            source_chart._blob = etree.tostring(
                source_chart_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            source_chart_blob = source_chart.blob
            first.save(first_path)
            second = Document()
            self.customize_test_theme(
                second,
                accent1="335577",
                major_latin="SharedChartFont",
                format_name="SharedChartFormat",
                background_mapping="accent1",
            )
            second.add_paragraph("BODY")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )
            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            copied_chart = next(
                relationship.target_part
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.CHART
            )
            self.assertEqual(copied_chart.blob, source_chart_blob)
            self.assertFalse(
                any(
                    relationship.reltype
                    == format_paper_module.RT.THEME_OVERRIDE
                    for relationship in copied_chart.rels.values()
                )
            )
            self.assertEqual(
                copied_chart.rels["rIdExternalData"].reltype,
                format_paper_module.RT.PACKAGE,
            )

            for authoritative_override in (False, True):
                with self.subTest(
                    missing_source_theme=True,
                    authoritative_override=authoritative_override,
                ):
                    case_name = (
                        "override" if authoritative_override else "unscoped"
                    )
                    first_path = tmp / f"missing-theme-{case_name}-first.docx"
                    second_path = tmp / f"missing-theme-{case_name}-second.docx"
                    output_path = tmp / f"missing-theme-{case_name}-merged.docx"

                    first = Document()
                    override_specs = ()
                    if authoritative_override:
                        override_specs = (
                            (
                                self.make_test_theme_override_blob(first),
                                format_paper_module.CT.OFC_THEME_OVERRIDE,
                                False,
                            ),
                        )
                    self.add_test_chart(
                        first,
                        f"MISSING SOURCE THEME {case_name}",
                        override_specs=override_specs,
                    )
                    source_default_fonts = (
                        first.styles.element.find(qn("w:docDefaults"))
                        .find(qn("w:rPrDefault"))
                        .find(qn("w:rPr"))
                        .find(qn("w:rFonts"))
                    )
                    source_default_fonts.attrib.clear()
                    for attribute_name in (
                        "w:ascii",
                        "w:hAnsi",
                        "w:eastAsia",
                        "w:cs",
                    ):
                        source_default_fonts.set(
                            qn(attribute_name),
                            "MissingThemeFallback",
                        )
                    theme_relationship = next(
                        relationship
                        for relationship in first.part.rels.values()
                        if relationship.reltype == format_paper_module.RT.THEME
                    )
                    first.part.drop_rel(theme_relationship.rId)
                    first.save(first_path)

                    second = Document()
                    self.customize_test_theme(
                        second,
                        accent1="1020D0",
                        major_latin="BodyChartFont",
                        format_name="BodyChartFormat",
                        background_mapping="accent1",
                    )
                    second.add_paragraph("BODY")
                    second.save(second_path)

                    if authoritative_override:
                        result = concatenate_documents(
                            first_path,
                            second_path,
                            output_path,
                            restart_body_page_number=False,
                        )
                        self.assertTrue(result["concatenated"])
                    else:
                        with self.assertRaises(DocumentConcatError):
                            concatenate_documents(
                                first_path,
                                second_path,
                                output_path,
                                restart_body_page_number=False,
                            )

    def test_concatenate_documents_materializes_inserted_theme_semantics(self):
        drawing_namespace = (
            "http://schemas.openxmlformats.org/drawingml/2006/main"
        )
        drawing_ns = {"a": drawing_namespace}

        def customize_theme(
            document,
            *,
            accent1,
            major_latin,
            minor_east_asia,
        ):
            theme_part = document.part.part_related_by(
                format_paper_module.RT.THEME
            )
            theme_root = etree.fromstring(theme_part.blob)
            accent = theme_root.find(
                ".//a:clrScheme/a:accent1/*",
                namespaces=drawing_ns,
            )
            accent.tag = f"{{{drawing_namespace}}}srgbClr"
            accent.attrib.clear()
            accent.set("val", accent1)
            theme_root.find(
                ".//a:fontScheme/a:majorFont/a:latin",
                namespaces=drawing_ns,
            ).set("typeface", major_latin)
            theme_root.find(
                ".//a:fontScheme/a:minorFont/a:ea",
                namespaces=drawing_ns,
            ).set("typeface", minor_east_asia)
            theme_part._blob = etree.tostring(
                theme_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

        def add_themed_run(paragraph, label):
            run = paragraph.add_run(label)
            run_properties = run._element.get_or_add_rPr()
            run_properties.append(
                parse_xml(
                    f'<w:rFonts {nsdecls("w")} '
                    'w:asciiTheme="majorHAnsi" '
                    'w:hAnsiTheme="majorHAnsi" '
                    'w:eastAsiaTheme="minorEastAsia"/>'
                )
            )
            run_properties.append(
                parse_xml(
                    f'<w:color {nsdecls("w")} w:val="000000" '
                    'w:themeColor="accent1"/>'
                )
            )
            run_properties.append(
                parse_xml(
                    f'<w:u {nsdecls("w")} w:val="single" '
                    'w:color="000000" w:themeColor="accent1"/>'
                )
            )
            run_properties.append(
                parse_xml(
                    f'<w:bdr {nsdecls("w")} w:val="single" w:sz="4" '
                    'w:space="0" w:color="000000" '
                    'w:themeColor="accent1"/>'
                )
            )
            run_properties.append(
                parse_xml(
                    f'<w:shd {nsdecls("w")} w:val="clear" '
                    'w:color="000000" w:themeColor="accent1" '
                    'w:fill="000000" w:themeFill="accent1"/>'
                )
            )
            run._element.append(
                parse_xml(
                    f'<w:drawing {nsdecls("w", "a")}> '
                    '<a:graphic><a:graphicData uri="urn:test:theme">'
                    '<a:solidFill><a:schemeClr val="accent1">'
                    '<a:tint val="20000"/>'
                    '</a:schemeClr></a:solidFill>'
                    '</a:graphicData></a:graphic></w:drawing>'
                )
            )
            return run

        def add_numbering_with_theme(document, paragraph):
            numbering_root = document.part.numbering_part.element
            abstract_num = parse_xml(
                f'<w:abstractNum {nsdecls("w")} w:abstractNumId="2000">'
                '<w:multiLevelType w:val="singleLevel"/>'
                '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
                '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>'
                '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" '
                'w:hAnsiTheme="majorHAnsi"/>'
                '<w:color w:val="000000" w:themeColor="accent1"/>'
                '</w:rPr></w:lvl></w:abstractNum>'
            )
            first_num = numbering_root.find(qn("w:num"))
            numbering_root.insert(
                numbering_root.index(first_num)
                if first_num is not None
                else len(numbering_root),
                abstract_num,
            )
            numbering_root.append(
                parse_xml(
                    f'<w:num {nsdecls("w")} w:numId="2000">'
                    '<w:abstractNumId w:val="2000"/></w:num>'
                )
            )
            paragraph._element.get_or_add_pPr().append(
                parse_xml(
                    f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/>'
                    '<w:numId w:val="2000"/></w:numPr>'
                )
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "theme-first.docx"
            second_path = tmp / "theme-second.docx"
            output_path = tmp / "theme-merged.docx"

            first = Document()
            customize_theme(
                first,
                accent1="E01020",
                major_latin="CoverMajor",
                minor_east_asia="CoverEastAsia",
            )
            cover_style = first.styles.add_style(
                "CoverThemeStyle",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            cover_style.element.append(
                parse_xml(
                    f'<w:rPr {nsdecls("w")}>'
                    '<w:rFonts w:asciiTheme="majorHAnsi" '
                    'w:hAnsiTheme="majorHAnsi"/>'
                    '<w:color w:val="000000" w:themeColor="accent1"/>'
                    '</w:rPr>'
                )
            )
            cover_paragraph = first.add_paragraph(
                "COVER THEME ",
                style=cover_style,
            )
            add_themed_run(cover_paragraph, "VISIBLE")
            add_numbering_with_theme(first, cover_paragraph)
            header_run = first.sections[0].header.paragraphs[0].add_run(
                "COVER THEME HEADER"
            )
            header_run._element.get_or_add_rPr().append(
                parse_xml(
                    f'<w:rFonts {nsdecls("w")} '
                    'w:asciiTheme="majorHAnsi" '
                    'w:hAnsiTheme="majorHAnsi"/>'
                )
            )
            first.save(first_path)

            second = Document()
            customize_theme(
                second,
                accent1="1020E0",
                major_latin="BodyMajor",
                minor_east_asia="BodyEastAsia",
            )
            body_paragraph = second.add_paragraph("BODY THEME ")
            add_themed_run(body_paragraph, "VISIBLE")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text.startswith("COVER THEME")
            )
            body = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text.startswith("BODY THEME")
            )
            cover_rfonts = next(cover._element.iter(qn("w:rFonts")))
            self.assertEqual(cover_rfonts.get(qn("w:ascii")), "CoverMajor")
            self.assertEqual(cover_rfonts.get(qn("w:hAnsi")), "CoverMajor")
            self.assertEqual(
                cover_rfonts.get(qn("w:eastAsia")),
                "CoverEastAsia",
            )
            for attribute in (
                "w:asciiTheme",
                "w:hAnsiTheme",
                "w:eastAsiaTheme",
            ):
                self.assertIsNone(cover_rfonts.get(qn(attribute)))

            cover_run_properties = cover._element.find(".//" + qn("w:rPr"))
            for tag, fallback_attribute in (
                ("w:color", "w:val"),
                ("w:u", "w:color"),
                ("w:bdr", "w:color"),
            ):
                element = cover_run_properties.find(qn(tag))
                self.assertEqual(element.get(qn(fallback_attribute)), "E01020")
                self.assertIsNone(element.get(qn("w:themeColor")))
            shading = cover_run_properties.find(qn("w:shd"))
            self.assertEqual(shading.get(qn("w:color")), "E01020")
            self.assertEqual(shading.get(qn("w:fill")), "E01020")
            self.assertIsNone(shading.get(qn("w:themeColor")))
            self.assertIsNone(shading.get(qn("w:themeFill")))

            self.assertFalse(list(cover._element.iter(qn("a:schemeClr"))))
            explicit_drawing_color = next(
                cover._element.iter(qn("a:srgbClr"))
            )
            self.assertEqual(explicit_drawing_color.get("val"), "E01020")
            self.assertEqual(
                explicit_drawing_color[0].tag,
                qn("a:tint"),
            )
            self.assertEqual(explicit_drawing_color[0].get("val"), "20000")

            body_rfonts = next(body._element.iter(qn("w:rFonts")))
            self.assertEqual(
                body_rfonts.get(qn("w:asciiTheme")),
                "majorHAnsi",
            )
            self.assertIsNone(body_rfonts.get(qn("w:ascii")))
            body_color = next(body._element.iter(qn("w:color")))
            self.assertEqual(body_color.get(qn("w:themeColor")), "accent1")

            merged_style = merged.styles["CoverThemeStyle"].element
            style_rfonts = next(merged_style.iter(qn("w:rFonts")))
            self.assertEqual(style_rfonts.get(qn("w:ascii")), "CoverMajor")
            self.assertIsNone(style_rfonts.get(qn("w:asciiTheme")))
            style_color = next(merged_style.iter(qn("w:color")))
            self.assertEqual(style_color.get(qn("w:val")), "E01020")
            self.assertIsNone(style_color.get(qn("w:themeColor")))

            num_id = cover._element.pPr.numPr.numId.val
            numbering_root = merged.part.numbering_part.element
            num = next(
                item
                for item in numbering_root.findall(qn("w:num"))
                if int(item.get(qn("w:numId"))) == num_id
            )
            abstract_id = num.find(qn("w:abstractNumId")).get(qn("w:val"))
            abstract_num = next(
                item
                for item in numbering_root.findall(qn("w:abstractNum"))
                if item.get(qn("w:abstractNumId")) == abstract_id
            )
            numbering_rfonts = next(abstract_num.iter(qn("w:rFonts")))
            self.assertEqual(
                numbering_rfonts.get(qn("w:ascii")),
                "CoverMajor",
            )
            self.assertIsNone(numbering_rfonts.get(qn("w:asciiTheme")))
            numbering_color = next(abstract_num.iter(qn("w:color")))
            self.assertEqual(numbering_color.get(qn("w:val")), "E01020")
            self.assertIsNone(numbering_color.get(qn("w:themeColor")))

            cover_header_part = next(
                relationship.target_part
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.HEADER
                and "COVER THEME HEADER"
                in "".join(relationship.target_part.element.itertext())
            )
            header_rfonts = next(
                cover_header_part.element.iter(qn("w:rFonts"))
            )
            self.assertEqual(header_rfonts.get(qn("w:ascii")), "CoverMajor")
            self.assertIsNone(header_rfonts.get(qn("w:asciiTheme")))

            theme_relationships = [
                relationship
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.THEME
            ]
            self.assertEqual(len(theme_relationships), 1)
            merged_theme = etree.fromstring(
                theme_relationships[0].target_part.blob
            )
            self.assertEqual(
                merged_theme.find(
                    ".//a:clrScheme/a:accent1/*",
                    namespaces=drawing_ns,
                ).get("val"),
                "1020E0",
            )
            self.assertEqual(
                merged_theme.find(
                    ".//a:fontScheme/a:majorFont/a:latin",
                    namespaces=drawing_ns,
                ).get("typeface"),
                "BodyMajor",
            )

    def test_concatenate_documents_resolves_theme_scripts_mapping_and_tints(self):
        drawing_namespace = (
            "http://schemas.openxmlformats.org/drawingml/2006/main"
        )
        drawing_ns = {"a": drawing_namespace}
        self.assertEqual(
            format_paper_module._apply_wml_theme_color_transform(
                "4F81BD",
                None,
                "BF",
            ),
            "365F91",
        )
        self.assertEqual(
            format_paper_module._apply_wml_theme_color_transform(
                "4F81BD",
                "99",
                None,
            ),
            "95B3D7",
        )

        def customize_theme(
            document,
            *,
            accent1,
            accent2,
            hans_font,
            arab_font,
            background_mapping,
        ):
            theme_part = document.part.part_related_by(
                format_paper_module.RT.THEME
            )
            theme_root = etree.fromstring(theme_part.blob)
            for color_name, value in (
                ("accent1", accent1),
                ("accent2", accent2),
            ):
                color = theme_root.find(
                    f".//a:clrScheme/a:{color_name}/*",
                    namespaces=drawing_ns,
                )
                color.tag = f"{{{drawing_namespace}}}srgbClr"
                color.attrib.clear()
                color.set("val", value)
            minor_font = theme_root.find(
                ".//a:fontScheme/a:minorFont",
                namespaces=drawing_ns,
            )
            minor_font.find("a:ea", namespaces=drawing_ns).set(
                "typeface",
                "",
            )
            minor_font.find("a:cs", namespaces=drawing_ns).set(
                "typeface",
                "",
            )
            next(
                font
                for font in minor_font.findall(
                    "a:font",
                    namespaces=drawing_ns,
                )
                if font.get("script") == "Hans"
            ).set("typeface", hans_font)
            next(
                font
                for font in minor_font.findall(
                    "a:font",
                    namespaces=drawing_ns,
                )
                if font.get("script") == "Arab"
            ).set("typeface", arab_font)
            theme_part._blob = etree.tostring(
                theme_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

            theme_languages = document.settings.element.find(
                qn("w:themeFontLang")
            )
            theme_languages.set(qn("w:eastAsia"), "zh-CN")
            theme_languages.set(qn("w:bidi"), "ar-SA")
            color_mapping = document.settings.element.find(
                qn("w:clrSchemeMapping")
            )
            color_mapping.set(qn("w:bg1"), background_mapping)

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "theme-script-first.docx"
            second_path = tmp / "theme-script-second.docx"
            output_path = tmp / "theme-script-merged.docx"

            first = Document()
            customize_theme(
                first,
                accent1="112233",
                accent2="778899",
                hans_font="CoverHans",
                arab_font="CoverArabic",
                background_mapping="accent2",
            )
            run = first.add_paragraph("COVER SCRIPT ").add_run("字体")
            run_properties = run._element.get_or_add_rPr()
            run_properties.append(
                parse_xml(
                    f'<w:rFonts {nsdecls("w")} '
                    'w:eastAsiaTheme="minorEastAsia" '
                    'w:cstheme="minorBidi"/>'
                )
            )
            run_properties.append(
                parse_xml(
                    f'<w:color {nsdecls("w")} w:val="ABCDEF" '
                    'w:themeColor="accent1" w:themeTint="99"/>'
                )
            )
            run_properties.append(
                parse_xml(
                    f'<w:shd {nsdecls("w")} w:val="clear" '
                    'w:fill="123456" w:themeFill="accent1" '
                    'w:themeFillShade="80"/>'
                )
            )
            run._element.append(
                parse_xml(
                    f'<w:drawing {nsdecls("w", "a")}> '
                    '<a:graphic><a:graphicData uri="urn:test:theme-map">'
                    '<a:solidFill><a:schemeClr val="bg1">'
                    '<a:shade val="60000"/>'
                    '</a:schemeClr></a:solidFill>'
                    '</a:graphicData></a:graphic></w:drawing>'
                )
            )
            first.save(first_path)

            second = Document()
            customize_theme(
                second,
                accent1="445566",
                accent2="AABBCC",
                hans_font="BodyHans",
                arab_font="BodyArabic",
                background_mapping="accent1",
            )
            second.add_paragraph("BODY")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            cover = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text.startswith("COVER SCRIPT")
            )
            rfonts = next(cover._element.iter(qn("w:rFonts")))
            self.assertEqual(rfonts.get(qn("w:eastAsia")), "CoverHans")
            self.assertEqual(rfonts.get(qn("w:cs")), "CoverArabic")
            self.assertIsNone(rfonts.get(qn("w:eastAsiaTheme")))
            self.assertIsNone(rfonts.get(qn("w:cstheme")))

            color = next(cover._element.iter(qn("w:color")))
            self.assertEqual(color.get(qn("w:val")), "3D7AB7")
            self.assertIsNone(color.get(qn("w:themeColor")))
            self.assertIsNone(color.get(qn("w:themeTint")))
            shading = next(cover._element.iter(qn("w:shd")))
            self.assertEqual(shading.get(qn("w:fill")), "081119")
            self.assertIsNone(shading.get(qn("w:themeFill")))
            self.assertIsNone(shading.get(qn("w:themeFillShade")))

            drawing_color = next(cover._element.iter(qn("a:srgbClr")))
            self.assertEqual(drawing_color.get("val"), "778899")
            self.assertEqual(drawing_color[0].tag, qn("a:shade"))
            self.assertEqual(drawing_color[0].get("val"), "60000")

    def test_concatenate_documents_rejects_contextual_theme_references(self):
        drawing_namespace = (
            "http://schemas.openxmlformats.org/drawingml/2006/main"
        )
        drawing_ns = {"a": drawing_namespace}

        def customize_theme(document, accent, font, format_name):
            theme_part = document.part.part_related_by(
                format_paper_module.RT.THEME
            )
            theme_root = etree.fromstring(theme_part.blob)
            color = theme_root.find(
                ".//a:clrScheme/a:accent1/*",
                namespaces=drawing_ns,
            )
            color.tag = f"{{{drawing_namespace}}}srgbClr"
            color.attrib.clear()
            color.set("val", accent)
            theme_root.find(
                ".//a:fontScheme/a:majorFont/a:latin",
                namespaces=drawing_ns,
            ).set("typeface", font)
            theme_root.find(
                ".//a:fmtScheme",
                namespaces=drawing_ns,
            ).set("name", format_name)
            theme_part._blob = etree.tostring(
                theme_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

        cases = {
            "placeholder-color": (
                '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            ),
            "font-reference": (
                '<a:fontRef idx="major">'
                '<a:schemeClr val="accent1"/></a:fontRef>'
            ),
            "format-reference": (
                '<a:fillRef idx="1">'
                '<a:schemeClr val="accent1"/></a:fillRef>'
            ),
        }
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            for case_name, drawing_content in cases.items():
                with self.subTest(case_name=case_name):
                    first_path = tmp / f"{case_name}-first.docx"
                    second_path = tmp / f"{case_name}-second.docx"
                    output_path = tmp / f"{case_name}-merged.docx"

                    first = Document()
                    customize_theme(
                        first,
                        "E01020",
                        "CoverMajor",
                        "CoverFormat",
                    )
                    first.add_paragraph(case_name)._element.append(
                        parse_xml(
                            f'<w:r {nsdecls("w", "a")}><w:drawing>'
                            '<a:graphic><a:graphicData uri="urn:test:context">'
                            f'{drawing_content}'
                            '</a:graphicData></a:graphic>'
                            '</w:drawing></w:r>'
                        )
                    )
                    first.save(first_path)

                    second = Document()
                    customize_theme(
                        second,
                        "1020E0",
                        "BodyMajor",
                        "BodyFormat",
                    )
                    second.add_paragraph("BODY")
                    second.save(second_path)

                    with self.assertRaises(DocumentConcatError):
                        concatenate_documents(
                            first_path,
                            second_path,
                            output_path,
                            restart_body_page_number=False,
                        )
                    self.assertFalse(output_path.exists())

    def test_concatenate_documents_rejects_format_refs_when_theme_relationships_differ(self):
        cover_theme_image = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC"
            "AAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )
        body_theme_image = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC"
            "AAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        )

        def attach_theme_image(document, payload):
            theme_part = document.part.part_related_by(
                format_paper_module.RT.THEME
            )
            theme_root = etree.fromstring(theme_part.blob)
            namespace = format_paper_module._THEME_NAMESPACE
            fill_styles = theme_root.find(
                ".//a:fmtScheme/a:fillStyleLst",
                namespaces={"a": namespace},
            )
            replacement = etree.Element(f"{{{namespace}}}blipFill")
            blip = etree.SubElement(replacement, f"{{{namespace}}}blip")
            blip.set(
                f"{{{format_paper_module._RELATIONSHIP_NAMESPACE}}}embed",
                "rIdThemeImage",
            )
            fill_styles.replace(fill_styles[0], replacement)
            theme_part._blob = etree.tostring(
                theme_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            theme_image = format_paper_module.Part(
                format_paper_module.PackURI(
                    "/word/media/theme-format-image.png"
                ),
                "image/png",
                payload,
                document.part.package,
            )
            theme_part.rels.add_relationship(
                format_paper_module.RT.IMAGE,
                theme_image,
                "rIdThemeImage",
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "theme-rel-format-first.docx"
            second_path = tmp / "theme-rel-format-second.docx"
            output_path = tmp / "theme-rel-format-merged.docx"

            first = Document()
            attach_theme_image(first, cover_theme_image)
            first.add_paragraph("THEME REL FORMAT")._element.append(
                parse_xml(
                    f'<w:r {nsdecls("w", "a")}><w:drawing>'
                    '<a:graphic><a:graphicData uri="urn:test:format-ref">'
                    '<a:fillRef idx="1"><a:schemeClr val="accent1"/>'
                    "</a:fillRef></a:graphicData></a:graphic>"
                    "</w:drawing></w:r>"
                )
            )
            first.save(first_path)

            second = Document()
            attach_theme_image(second, body_theme_image)
            second.add_paragraph("BODY")
            second.save(second_path)

            with self.assertRaises(DocumentConcatError):
                concatenate_documents(
                    first_path,
                    second_path,
                    output_path,
                    restart_body_page_number=False,
                )
            self.assertFalse(output_path.exists())

    def test_concatenate_documents_preserves_legacy_ole_relationships(self):
        relationship_type = (
            "http://schemas.openxmlformats.org/officeDocument/2006/"
            "relationships/oleObject"
        )
        content_type = (
            "application/vnd.openxmlformats-officedocument.oleObject"
        )

        def add_ole_object(document, label, payload):
            part = format_paper_module.Part(
                format_paper_module.PackURI(
                    "/word/embeddings/oleObject1.bin"
                ),
                content_type,
                payload,
                document.part.package,
            )
            relationship_id = document.part.relate_to(part, relationship_type)
            paragraph = document.add_paragraph(label)
            paragraph._element.append(
                parse_xml(
                    f'<w:r {nsdecls("w", "v", "o")}>'
                    "<w:object><v:shape/>"
                    f'<o:OLEObject Type="Embed" ProgID="Package" '
                    f'o:relid="{relationship_id}"/>'
                    "</w:object></w:r>"
                )
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "ole-first.docx"
            second_path = tmp / "ole-second.docx"
            output_path = tmp / "ole-merged.docx"
            first_payload = b"FIRST OLE PAYLOAD"
            second_payload = b"SECOND OLE PAYLOAD"

            first = Document()
            add_ole_object(first, "FIRST OLE", first_payload)
            first.save(first_path)

            second = Document()
            add_ole_object(second, "SECOND OLE", second_payload)
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            objects = list(merged.element.body.iter(qn("o:OLEObject")))
            self.assertEqual(len(objects), 2)
            relationships = [
                merged.part.rels[obj.get(qn("o:relid"))]
                for obj in objects
            ]
            self.assertTrue(
                all(rel.reltype == relationship_type for rel in relationships)
            )
            self.assertEqual(
                {rel.target_part.blob for rel in relationships},
                {first_payload, second_payload},
            )
            self.assertEqual(
                len({rel.target_part.partname for rel in relationships}),
                2,
            )

    def test_concatenate_documents_preserves_all_vml_image_relationship_attributes(self):
        primary_image = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )
        alternate_image = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "vml-first.docx"
            second_path = tmp / "plain-second.docx"
            output_path = tmp / "vml-merged.docx"

            first = Document()
            primary_part = first.part.package.get_or_add_image_part(
                BytesIO(primary_image)
            )
            alternate_part = first.part.package.get_or_add_image_part(
                BytesIO(alternate_image)
            )
            primary_rid = first.part.relate_to(
                primary_part,
                format_paper_module.RT.IMAGE,
            )
            alternate_rid = first.part.relate_to(
                alternate_part,
                format_paper_module.RT.IMAGE,
            )
            linked_image_url = "https://example.com/vml-linked-image"
            linked_image_rid = first.part.relate_to(
                linked_image_url,
                format_paper_module.RT.HYPERLINK,
                is_external=True,
            )

            o_relid_paragraph = first.add_paragraph("VML O RELID")
            o_relid_paragraph._element.append(
                parse_xml(
                    f'<w:r {nsdecls("w", "v", "o")}><w:pict><v:shape>'
                    f'<v:imagedata o:relid="{primary_rid}"/>'
                    "</v:shape></w:pict></w:r>"
                )
            )
            alternate_paragraph = first.add_paragraph("VML ALTERNATE")
            alternate_paragraph._element.append(
                parse_xml(
                    f'<w:r {nsdecls("w", "v", "r")}><w:pict><v:shape>'
                    f'<v:imagedata r:id="{primary_rid}" '
                    f'r:pict="{alternate_rid}" r:href="{linked_image_rid}"/>'
                    "</v:shape></w:pict></w:r>"
                )
            )
            fill_stroke_paragraph = first.add_paragraph("VML FILL STROKE")
            fill_stroke_paragraph._element.append(
                parse_xml(
                    f'<w:r {nsdecls("w", "v", "r")}><w:pict><v:rect>'
                    f'<v:fill r:id="{alternate_rid}"/>'
                    f'<v:stroke r:id="{primary_rid}"/>'
                    "</v:rect></w:pict></w:r>"
                )
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph("普通正文")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            image_nodes = list(
                merged.element.body.iter(qn("v:imagedata"))
            )
            self.assertEqual(len(image_nodes), 2)
            o_relid_relationship = merged.part.rels[
                image_nodes[0].get(qn("o:relid"))
            ]
            self.assertEqual(
                o_relid_relationship.target_part.blob,
                primary_image,
            )
            primary_relationship = merged.part.rels[
                image_nodes[1].get(qn("r:id"))
            ]
            alternate_relationship = merged.part.rels[
                image_nodes[1].get(qn("r:pict"))
            ]
            self.assertEqual(primary_relationship.target_part.blob, primary_image)
            self.assertEqual(
                alternate_relationship.target_part.blob,
                alternate_image,
            )
            self.assertNotEqual(
                primary_relationship.rId,
                alternate_relationship.rId,
            )
            linked_relationship = merged.part.rels[
                image_nodes[1].get(qn("r:href"))
            ]
            self.assertTrue(linked_relationship.is_external)
            self.assertEqual(
                linked_relationship.reltype,
                format_paper_module.RT.HYPERLINK,
            )
            self.assertEqual(
                linked_relationship.target_ref,
                linked_image_url,
            )
            fill = next(merged.element.body.iter(qn("v:fill")))
            stroke = next(merged.element.body.iter(qn("v:stroke")))
            self.assertEqual(
                merged.part.rels[fill.get(qn("r:id"))].target_part.blob,
                alternate_image,
            )
            self.assertEqual(
                merged.part.rels[stroke.get(qn("r:id"))].target_part.blob,
                primary_image,
            )
            with ZipFile(output_path) as archive:
                member_names = archive.namelist()
            self.assertEqual(len(member_names), len(set(member_names)))

    def test_concatenate_documents_preserves_drawingml_hyperlink_sound(self):
        image_payload = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )
        audio_payload = b"RIFF\x00\x00\x00\x00WAVEfmt DRAWINGML SOUND"
        audio_partname = "/word/media/hyperlink-sound.wav"
        audio_relationship_id = "rIdHyperlinkSound"
        hyperlink_url = "https://example.test/drawing-sound"
        video_url = "https://example.test/drawing-video.mp4"

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "drawing-sound-first.docx"
            second_path = tmp / "plain-second.docx"
            output_path = tmp / "drawing-sound-merged.docx"

            first = Document()
            picture_run = first.add_paragraph("DRAWING SOUND").add_run()
            picture_run.add_picture(BytesIO(image_payload))
            audio_part = format_paper_module.Part(
                format_paper_module.PackURI(audio_partname),
                "audio/wav",
                audio_payload,
                first.part.package,
            )
            first.part.rels.add_relationship(
                format_paper_module.RT.AUDIO,
                audio_part,
                audio_relationship_id,
                is_external=False,
            )
            hyperlink_relationship_id = first.part.rels.get_or_add_ext_rel(
                format_paper_module.RT.HYPERLINK,
                hyperlink_url,
            )
            video_relationship_id = first.part.rels.get_or_add_ext_rel(
                format_paper_module.RT.VIDEO,
                video_url,
            )
            drawing_properties = picture_run._element.find(
                ".//" + qn("wp:docPr")
            )
            hyperlink = drawing_properties.makeelement(
                qn("a:hlinkClick"),
                {qn("r:id"): hyperlink_relationship_id},
            )
            hyperlink.append(
                hyperlink.makeelement(
                    qn("a:snd"),
                    {
                        qn("r:embed"): audio_relationship_id,
                        "name": "hyperlink-sound.wav",
                    },
                )
            )
            drawing_properties.append(hyperlink)
            drawing_properties.append(
                drawing_properties.makeelement(
                    qn("a:videoFile"),
                    {qn("r:link"): video_relationship_id},
                )
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph("MASTER")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            merged_hyperlink = next(
                merged.element.body.iter(qn("a:hlinkClick"))
            )
            hyperlink_relationship = merged.part.rels[
                merged_hyperlink.get(qn("r:id"))
            ]
            self.assertTrue(hyperlink_relationship.is_external)
            self.assertEqual(
                hyperlink_relationship.reltype,
                format_paper_module.RT.HYPERLINK,
            )
            self.assertEqual(hyperlink_relationship.target_ref, hyperlink_url)

            merged_sound = next(merged_hyperlink.iter(qn("a:snd")))
            audio_relationship = merged.part.rels[
                merged_sound.get(qn("r:embed"))
            ]
            self.assertFalse(audio_relationship.is_external)
            self.assertEqual(
                audio_relationship.reltype,
                format_paper_module.RT.AUDIO,
            )
            self.assertEqual(audio_relationship.target_part.blob, audio_payload)
            merged_video = next(
                merged.element.body.iter(qn("a:videoFile"))
            )
            video_relationship = merged.part.rels[
                merged_video.get(qn("r:link"))
            ]
            self.assertTrue(video_relationship.is_external)
            self.assertEqual(
                video_relationship.reltype,
                format_paper_module.RT.VIDEO,
            )
            self.assertEqual(video_relationship.target_ref, video_url)

    def test_concatenate_documents_preserves_alternate_content_3d_model(self):
        model_namespace = (
            "http://schemas.microsoft.com/office/drawing/2017/model3d"
        )
        model_relationship_type = (
            "http://schemas.microsoft.com/office/2017/06/relationships/model3d"
        )
        model_payload = b"glTF BINARY MODEL PAYLOAD"
        image_layer_namespace = (
            "http://schemas.microsoft.com/office/drawing/2010/main"
        )
        raster_payload = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "model3d-first.docx"
            second_path = tmp / "plain-second.docx"
            output_path = tmp / "model3d-merged.docx"

            first = Document()
            model_part = format_paper_module.Part(
                format_paper_module.PackURI("/word/media/model3d1.glb"),
                "model/gltf-binary",
                model_payload,
                first.part.package,
            )
            first.part.rels.add_relationship(
                model_relationship_type,
                model_part,
                "modelRel",
                is_external=False,
            )
            raster_part = first.part.package.get_or_add_image_part(
                BytesIO(raster_payload)
            )
            raster_relationship_id = first.part.relate_to(
                raster_part,
                format_paper_module.RT.IMAGE,
            )
            first.add_paragraph("3D MODEL")._element.append(
                parse_xml(
                    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/'
                    'markup-compatibility/2006" '
                    f'xmlns:am3d="{model_namespace}" '
                    f'xmlns:a14="{image_layer_namespace}" '
                    f'{nsdecls("w", "a", "r")}>'
                    '<mc:Choice Requires="am3d"><w:r><w:drawing>'
                    '<a:graphic><a:graphicData '
                    f'uri="{model_namespace}">'
                    '<am3d:model3d r:embed="modelRel"/>'
                    f'<am3d:blip r:embed="{raster_relationship_id}"/>'
                    f'<a14:imgLayer r:embed="{raster_relationship_id}"/>'
                    '</a:graphicData></a:graphic>'
                    '</w:drawing></w:r></mc:Choice>'
                    '<mc:Fallback><w:r><w:t>3D FALLBACK</w:t></w:r>'
                    '</mc:Fallback></mc:AlternateContent>'
                )
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph("MASTER")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            model = next(
                merged.element.body.iter(f"{{{model_namespace}}}model3d")
            )
            model_relationship = merged.part.rels[
                model.get(qn("r:embed"))
            ]
            self.assertFalse(model_relationship.is_external)
            self.assertEqual(
                model_relationship.reltype,
                model_relationship_type,
            )
            self.assertEqual(model_relationship.target_part.blob, model_payload)
            for image_node in (
                next(merged.element.body.iter(f"{{{model_namespace}}}blip")),
                next(
                    merged.element.body.iter(
                        f"{{{image_layer_namespace}}}imgLayer"
                    )
                ),
            ):
                image_relationship = merged.part.rels[
                    image_node.get(qn("r:embed"))
                ]
                self.assertEqual(
                    image_relationship.reltype,
                    format_paper_module.RT.IMAGE,
                )
                self.assertEqual(
                    image_relationship.target_part.blob,
                    raster_payload,
                )
            self.assertIn(
                "3D FALLBACK",
                "".join(
                    text.text or ""
                    for text in merged.element.body.iter(qn("w:t"))
                ),
            )
            with ZipFile(output_path) as archive:
                member_names = archive.namelist()
                self.assertIsNone(archive.testzip())
            self.assertIn(
                str(model_relationship.target_part.partname).lstrip("/"),
                member_names,
            )
            self.assertEqual(
                len(member_names),
                len({name.casefold() for name in member_names}),
            )

    def test_concatenate_documents_preserves_custom_xml_data_bindings(self):
        custom_xml_namespace = (
            "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
        )
        item_id_attribute = f"{{{custom_xml_namespace}}}itemID"
        word_2012_namespace = (
            "http://schemas.microsoft.com/office/word/2012/wordml"
        )
        word_2012_data_binding_tag = (
            f"{{{word_2012_namespace}}}dataBinding"
        )
        fixture_path = Path(__file__).resolve().parent.parent / "test_input.docx"

        def make_bound_content_control(document, text, *, word_2012=False):
            custom_xml_part = document.part.rels.part_with_reltype(
                format_paper_module.RT.CUSTOM_XML
            )
            properties_part = custom_xml_part.rels.part_with_reltype(
                format_paper_module.RT.CUSTOM_XML_PROPS
            )
            properties_root = etree.fromstring(properties_part.blob)
            item_id = properties_root.get(item_id_attribute)
            binding_tag = "w15:dataBinding" if word_2012 else "w:dataBinding"
            content_control = parse_xml(
                f'<w:sdt {nsdecls("w")} xmlns:w15="{word_2012_namespace}">'
                "<w:sdtPr>"
                f'<{binding_tag} w:storeItemID="{item_id}" w:xpath="/b:Sources" '
                'w:prefixMappings="xmlns:b=&quot;http://schemas.openxmlformats.org/'
                'officeDocument/2006/bibliography&quot;"/>'
                "</w:sdtPr>"
                "<w:sdtContent><w:p><w:r>"
                f"<w:t>{text}</w:t>"
                "</w:r></w:p></w:sdtContent>"
                "</w:sdt>"
            )
            return item_id, content_control

        def add_bound_content_control(document, text, *, word_2012=False):
            item_id, content_control = make_bound_content_control(
                document,
                text,
                word_2012=word_2012,
            )
            document.element.body.insert(
                len(document.element.body) - 1,
                content_control,
            )
            return item_id

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "custom-xml-first.docx"
            second_path = tmp / "custom-xml-second.docx"
            output_path = tmp / "custom-xml-merged.docx"

            first = Document(fixture_path)
            original_item_id = add_bound_content_control(first, "FIRST BINDING")
            self.assertEqual(
                add_bound_content_control(
                    first,
                    "FIRST W15 BINDING",
                    word_2012=True,
                ),
                original_item_id,
            )
            header_item_id, header_control = make_bound_content_control(
                first,
                "FIRST HEADER BINDING",
            )
            self.assertEqual(header_item_id, original_item_id)
            first.sections[0].header._element.append(header_control)
            first.save(first_path)
            self.inject_simple_footnote(
                first_path,
                "FIRST FOOTNOTE BINDING",
                footnote_content_xml=(
                    f'<w:sdt {nsdecls("w")}><w:sdtPr>'
                    f'<w:dataBinding w:storeItemID="{original_item_id}"/>'
                    "</w:sdtPr><w:sdtContent><w:p><w:r>"
                    "<w:t>FIRST FOOTNOTE BINDING</w:t>"
                    "</w:r></w:p></w:sdtContent></w:sdt>"
                ),
            )

            second = Document(fixture_path)
            self.assertEqual(
                add_bound_content_control(second, "SECOND BINDING"),
                original_item_id,
            )
            self.assertEqual(
                add_bound_content_control(
                    second,
                    "SECOND W15 BINDING",
                    word_2012=True,
                ),
                original_item_id,
            )
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            self.assertTrue(output_path.exists())
            Document(output_path)
            with ZipFile(output_path) as archive:
                self.assertIsNone(archive.testzip())
                document_root = etree.fromstring(
                    archive.read("word/document.xml")
                )
                binding_roots = [
                    document_root,
                    etree.fromstring(archive.read("word/footnotes.xml")),
                ]
                binding_roots.extend(
                    etree.fromstring(archive.read(name))
                    for name in archive.namelist()
                    if re.fullmatch(r"word/header\d+\.xml", name)
                )
                item_property_names = sorted(
                    name
                    for name in archive.namelist()
                    if re.fullmatch(r"customXml/itemProps\d+\.xml", name)
                )
                item_names = sorted(
                    name
                    for name in archive.namelist()
                    if re.fullmatch(r"customXml/item\d+\.xml", name)
                )

                item_ids = {
                    etree.fromstring(archive.read(name)).get(item_id_attribute)
                    for name in item_property_names
                }
                binding_ids = {}
                for binding_root in binding_roots:
                    for content_control in binding_root.findall(
                        ".//" + qn("w:sdt")
                    ):
                        text = "".join(
                            node.text or ""
                            for node in content_control.findall(
                                ".//" + qn("w:t")
                            )
                        )
                        if text not in {
                            "FIRST BINDING",
                            "FIRST W15 BINDING",
                            "FIRST HEADER BINDING",
                            "FIRST FOOTNOTE BINDING",
                            "SECOND BINDING",
                            "SECOND W15 BINDING",
                        }:
                            continue
                        data_binding = content_control.find(
                            ".//" + qn("w:dataBinding")
                        )
                        if data_binding is None:
                            data_binding = content_control.find(
                                ".//" + word_2012_data_binding_tag
                            )
                        binding_ids[text] = data_binding.get(
                            qn("w:storeItemID")
                        )

            self.assertEqual(len(item_names), 2)
            self.assertEqual(len(item_property_names), 2)
            self.assertEqual(len(item_ids), 2)
            self.assertEqual(
                set(binding_ids),
                {
                    "FIRST BINDING",
                    "FIRST W15 BINDING",
                    "FIRST HEADER BINDING",
                    "FIRST FOOTNOTE BINDING",
                    "SECOND BINDING",
                    "SECOND W15 BINDING",
                },
            )
            self.assertEqual(set(binding_ids.values()), item_ids)
            self.assertEqual(binding_ids["SECOND BINDING"], original_item_id)
            self.assertEqual(
                binding_ids["SECOND W15 BINDING"],
                original_item_id,
            )
            self.assertNotEqual(binding_ids["FIRST BINDING"], original_item_id)
            self.assertEqual(
                binding_ids["FIRST W15 BINDING"],
                binding_ids["FIRST BINDING"],
            )
            self.assertEqual(
                binding_ids["FIRST HEADER BINDING"],
                binding_ids["FIRST BINDING"],
            )
            self.assertEqual(
                binding_ids["FIRST FOOTNOTE BINDING"],
                binding_ids["FIRST BINDING"],
            )

    def test_concatenate_documents_avoids_case_insensitive_partname_collisions(self):
        fixture_path = Path(__file__).resolve().parent.parent / "test_input.docx"
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "uppercase-custom-xml.docx"
            second_path = tmp / "lowercase-custom-xml.docx"
            output_path = tmp / "case-safe-merged.docx"

            first = Document(fixture_path)
            first_custom_xml = first.part.rels.part_with_reltype(
                format_paper_module.RT.CUSTOM_XML
            )
            first_custom_xml.partname = format_paper_module.PackURI(
                "/customXml/ITEM1.xml"
            )
            first.add_paragraph("大写 Part 名")
            first.save(first_path)

            second = Document(fixture_path)
            second.add_paragraph("小写 Part 名")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            with ZipFile(output_path) as archive:
                member_names = archive.namelist()
                self.assertIsNone(archive.testzip())
            self.assertEqual(len(member_names), len(set(member_names)))
            self.assertEqual(
                len(member_names),
                len({name.casefold() for name in member_names}),
            )

            merged = Document(output_path)
            custom_xml_relationships = [
                relationship
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.CUSTOM_XML
            ]
            self.assertEqual(len(custom_xml_relationships), 2)
            target_names = [
                str(relationship.target_part.partname)
                for relationship in custom_xml_relationships
            ]
            self.assertEqual(
                len(target_names),
                len({name.casefold() for name in target_names}),
            )

    def test_concatenate_documents_avoids_custom_graph_and_body_image_collisions(self):
        first_image = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )
        second_image = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        )
        custom_payload = b"CUSTOM XML PRIVATE RESOURCE"
        custom_resource_relationship = "https://example.test/custom-resource"
        fixture_path = Path(__file__).resolve().parent.parent / "test_input.docx"

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "custom-resource-first.docx"
            second_path = tmp / "body-image-second.docx"
            output_path = tmp / "image-collision-safe.docx"

            first = Document(fixture_path)
            custom_xml_part = first.part.rels.part_with_reltype(
                format_paper_module.RT.CUSTOM_XML
            )
            custom_resource = format_paper_module.Part(
                format_paper_module.PackURI("/word/media/image2.png"),
                "image/png",
                custom_payload,
                first.part.package,
            )
            custom_xml_part.relate_to(
                custom_resource,
                custom_resource_relationship,
            )
            first.add_paragraph("插入文档图片").add_run().add_picture(
                BytesIO(first_image),
                width=Inches(0.2),
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph("主文档图片").add_run().add_picture(
                BytesIO(second_image),
                width=Inches(0.2),
            )
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            with ZipFile(output_path) as archive:
                member_names = archive.namelist()
                self.assertIsNone(archive.testzip())
            self.assertEqual(len(member_names), len(set(member_names)))
            self.assertEqual(
                len(member_names),
                len({name.casefold() for name in member_names}),
            )

            merged = Document(output_path)
            body_image_payloads = {
                relationship.target_part.blob
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.IMAGE
            }
            self.assertEqual(
                body_image_payloads,
                {first_image, second_image},
            )
            copied_custom_xml = next(
                relationship.target_part
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.CUSTOM_XML
                and any(
                    nested.reltype == custom_resource_relationship
                    for nested in relationship.target_part.rels.values()
                )
            )
            copied_resource = copied_custom_xml.rels.part_with_reltype(
                custom_resource_relationship
            )
            self.assertEqual(copied_resource.blob, custom_payload)
            self.assertNotIn(copied_resource.blob, body_image_payloads)

    def test_concatenate_documents_reuses_images_shared_with_custom_xml_graphs(self):
        shared_image = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9VE3d2wAAAAASUVORK5CYII="
        )
        shared_relationship_type = "https://example.test/shared-image"
        fixture_path = Path(__file__).resolve().parent.parent / "test_input.docx"

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "shared-image-first.docx"
            second_path = tmp / "plain-second.docx"
            output_path = tmp / "shared-image-merged.docx"

            first = Document(fixture_path)
            run = first.add_paragraph("共享图片").add_run()
            run.add_picture(BytesIO(shared_image), width=Inches(0.2))
            image_relationship = next(
                relationship
                for relationship in first.part.rels.values()
                if relationship.reltype == format_paper_module.RT.IMAGE
                and relationship.target_part.blob == shared_image
            )
            custom_xml_part = first.part.rels.part_with_reltype(
                format_paper_module.RT.CUSTOM_XML
            )
            custom_xml_part.relate_to(
                image_relationship.target_part,
                shared_relationship_type,
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph("普通正文")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            copied_custom_xml = next(
                relationship.target_part
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.CUSTOM_XML
                and any(
                    nested.reltype == shared_relationship_type
                    for nested in relationship.target_part.rels.values()
                )
            )
            custom_target = copied_custom_xml.rels.part_with_reltype(
                shared_relationship_type
            )
            body_target = next(
                relationship.target_part
                for relationship in merged.part.rels.values()
                if relationship.reltype == format_paper_module.RT.IMAGE
                and relationship.target_part.blob == shared_image
            )
            self.assertEqual(custom_target.partname, body_target.partname)
            self.assertEqual(custom_target.blob, shared_image)

            with ZipFile(output_path) as archive:
                member_names = archive.namelist()
            shared_member = str(custom_target.partname).lstrip("/")
            self.assertEqual(member_names.count(shared_member), 1)
            self.assertEqual(len(member_names), len(set(member_names)))

    def test_concatenate_documents_preserves_link_only_drawingml_images(self):
        image_url = "https://example.test/linked-only.png"
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "linked-image-first.docx"
            second_path = tmp / "plain-second.docx"
            output_path = tmp / "linked-image-merged.docx"

            first = Document()
            relationship_id = first.part.relate_to(
                image_url,
                format_paper_module.RT.IMAGE,
                is_external=True,
            )
            paragraph = first.add_paragraph("仅外链图片")
            paragraph._element.append(
                parse_xml(
                    f'<w:r {nsdecls("w", "a", "r")}><w:drawing>'
                    f'<a:blip r:link="{relationship_id}"/>'
                    "</w:drawing></w:r>"
                )
            )
            first.save(first_path)

            second = Document()
            second.add_paragraph("普通正文")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            linked_blip = next(
                blip
                for blip in merged.element.body.iter(qn("a:blip"))
                if blip.get(qn("r:link"))
            )
            self.assertIsNone(linked_blip.get(qn("r:embed")))
            copied_relationship = merged.part.rels[
                linked_blip.get(qn("r:link"))
            ]
            self.assertTrue(copied_relationship.is_external)
            self.assertEqual(copied_relationship.reltype, format_paper_module.RT.IMAGE)
            self.assertEqual(copied_relationship.target_ref, image_url)

    def test_concatenate_documents_preserves_positive_id_footnote_separators(self):
        def convert_to_positive_special_ids(docx_path):
            with ZipFile(docx_path, "r") as source:
                entries = source.infolist()
                payloads = {
                    entry.filename: source.read(entry.filename)
                    for entry in entries
                }

            footnotes_root = etree.fromstring(payloads["word/footnotes.xml"])
            separator = next(
                note
                for note in footnotes_root.findall(qn("w:footnote"))
                if note.find(".//" + qn("w:separator")) is not None
            )
            continuation = next(
                note
                for note in footnotes_root.findall(qn("w:footnote"))
                if note.find(".//" + qn("w:continuationSeparator"))
                is not None
            )
            for note, positive_id in ((separator, "0"), (continuation, "1")):
                note.set(qn("w:id"), positive_id)
                note.attrib.pop(qn("w:type"), None)
            payloads["word/footnotes.xml"] = etree.tostring(
                footnotes_root,
                encoding="UTF-8",
                xml_declaration=True,
                standalone=True,
            )

            settings_root = etree.fromstring(payloads["word/settings.xml"])
            footnote_properties = settings_root.find(qn("w:footnotePr"))
            if footnote_properties is None:
                footnote_properties = etree.Element(qn("w:footnotePr"))
                compatibility = settings_root.find(qn("w:compat"))
                insert_index = (
                    settings_root.index(compatibility)
                    if compatibility is not None
                    else len(settings_root)
                )
                settings_root.insert(insert_index, footnote_properties)
            for reference in list(
                footnote_properties.findall(qn("w:footnote"))
            ):
                footnote_properties.remove(reference)
            for positive_id in ("0", "1"):
                reference = etree.SubElement(
                    footnote_properties,
                    qn("w:footnote"),
                )
                reference.set(qn("w:id"), positive_id)
            payloads["word/settings.xml"] = etree.tostring(
                settings_root,
                encoding="UTF-8",
                xml_declaration=True,
                standalone=True,
            )

            with ZipFile(docx_path, "w") as target:
                for entry in entries:
                    target.writestr(entry, payloads[entry.filename])

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "positive-separators-first.docx"
            second_path = tmp / "plain-second.docx"
            output_path = tmp / "positive-separators-merged.docx"

            first = Document()
            first.add_paragraph("带 WPS 脚注分隔符")
            first.save(first_path)
            self.inject_simple_footnote(first_path, "用户脚注", footnote_id=2)
            convert_to_positive_special_ids(first_path)

            # 脚注后处理只应计数/格式化用户脚注，不得改写分隔符。
            self.assertEqual(
                format_paper_module.format_docx_footnotes(first_path),
                1,
            )

            second = Document()
            second.add_paragraph("普通正文")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            with ZipFile(output_path) as archive:
                footnotes_root = etree.fromstring(archive.read("word/footnotes.xml"))
                settings_root = etree.fromstring(archive.read("word/settings.xml"))
                document_root = etree.fromstring(archive.read("word/document.xml"))
            self.assertEqual(
                {
                    note.get(qn("w:id"))
                    for note in footnotes_root.findall(qn("w:footnote"))
                },
                {"0", "1", "2"},
            )
            self.assertEqual(
                {
                    reference.get(qn("w:id"))
                    for reference in settings_root.findall(
                        ".//" + qn("w:footnotePr") + "/" + qn("w:footnote")
                    )
                },
                {"0", "1"},
            )
            self.assertEqual(
                [
                    reference.get(qn("w:id"))
                    for reference in document_root.findall(
                        ".//" + qn("w:footnoteReference")
                    )
                ],
                ["2"],
            )

    def test_concatenate_documents_preserves_positive_id_endnote_separators(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "positive-endnote-separators.docx"
            second_path = tmp / "plain-second.docx"
            output_path = tmp / "positive-endnotes-merged.docx"

            first = Document()
            first.add_paragraph("带 WPS 尾注分隔符")
            first.save(first_path)
            self.inject_simple_endnote(first_path, "用户尾注", endnote_id=2)

            source = Document(first_path)
            endnotes_part = source.part.rels.part_with_reltype(
                format_paper_module.RT.ENDNOTES
            )
            endnotes_root = etree.fromstring(endnotes_part.blob)
            separator = next(
                note
                for note in endnotes_root.findall(qn("w:endnote"))
                if note.find(".//" + qn("w:separator")) is not None
            )
            continuation = next(
                note
                for note in endnotes_root.findall(qn("w:endnote"))
                if note.find(".//" + qn("w:continuationSeparator"))
                is not None
            )
            for note, positive_id in ((separator, "0"), (continuation, "1")):
                note.set(qn("w:id"), positive_id)
                note.attrib.pop(qn("w:type"), None)
            endnotes_part._blob = etree.tostring(
                endnotes_root,
                encoding="UTF-8",
                xml_declaration=True,
                standalone=True,
            )

            settings = source.settings.element
            endnote_properties = settings.find(qn("w:endnotePr"))
            if endnote_properties is None:
                endnote_properties = OxmlElement("w:endnotePr")
                settings.insert_element_before(
                    endnote_properties,
                    "w:compat",
                    "w:docVars",
                    "w:rsids",
                )
            for positive_id in ("0", "1"):
                reference = OxmlElement("w:endnote")
                reference.set(qn("w:id"), positive_id)
                endnote_properties.append(reference)
            source.save(first_path)

            second = Document()
            second.add_paragraph("普通正文")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            with ZipFile(output_path) as archive:
                endnotes_root = etree.fromstring(archive.read("word/endnotes.xml"))
                settings_root = etree.fromstring(archive.read("word/settings.xml"))
            self.assertEqual(
                {
                    note.get(qn("w:id"))
                    for note in endnotes_root.findall(qn("w:endnote"))
                },
                {"0", "1", "2"},
            )
            self.assertEqual(
                {
                    reference.get(qn("w:id"))
                    for reference in settings_root.findall(
                        ".//" + qn("w:endnotePr") + "/" + qn("w:endnote")
                    )
                },
                {"0", "1"},
            )

    def test_concatenate_documents_avoids_case_collisions_for_created_story_parts(self):
        resource_relationship_type = "https://example.test/uppercase-footnotes"
        resource_payload = b"<resource>NOT A FOOTNOTES STORY</resource>"
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "footnote-first.docx"
            second_path = tmp / "uppercase-resource-second.docx"
            output_path = tmp / "story-part-case-safe.docx"

            first = Document()
            first.add_paragraph("插入文档脚注")
            first.save(first_path)
            self.inject_simple_footnote(first_path, "脚注内容", footnote_id=2)

            second = Document()
            custom_xml_part = second.part.rels.part_with_reltype(
                format_paper_module.RT.CUSTOM_XML
            )
            uppercase_resource = format_paper_module.Part(
                format_paper_module.PackURI("/word/FOOTNOTES.xml"),
                "application/xml",
                resource_payload,
                second.part.package,
            )
            custom_xml_part.relate_to(
                uppercase_resource,
                resource_relationship_type,
            )
            second.add_paragraph("主文档正文")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            with ZipFile(output_path) as archive:
                member_names = archive.namelist()
                self.assertIsNone(archive.testzip())
            self.assertEqual(
                len(member_names),
                len({name.casefold() for name in member_names}),
            )

            merged = Document(output_path)
            footnotes_part = merged.part.rels.part_with_reltype(
                format_paper_module.RT.FOOTNOTES
            )
            self.assertNotEqual(
                str(footnotes_part.partname).casefold(),
                "/word/footnotes.xml".casefold(),
            )
            copied_resource = next(
                relationship.target_part
                for custom_relationship in merged.part.rels.values()
                if custom_relationship.reltype == format_paper_module.RT.CUSTOM_XML
                for relationship in custom_relationship.target_part.rels.values()
                if relationship.reltype == resource_relationship_type
            )
            self.assertEqual(copied_resource.blob, resource_payload)

    def test_concatenate_documents_avoids_case_collisions_for_created_numbering_part(self):
        resource_relationship_type = "https://example.test/uppercase-numbering"
        resource_payload = b"<resource>NOT A NUMBERING STORY</resource>"
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "numbered-first.docx"
            second_path = tmp / "uppercase-resource-second.docx"
            output_path = tmp / "numbering-case-safe.docx"

            first = Document()
            numbered = first.add_paragraph("插入文档编号")
            num_pr = OxmlElement("w:numPr")
            ilvl = OxmlElement("w:ilvl")
            ilvl.set(qn("w:val"), "0")
            num_id = OxmlElement("w:numId")
            num_id.set(qn("w:val"), "5")
            num_pr.extend((ilvl, num_id))
            numbered._element.get_or_add_pPr().append(num_pr)
            first.save(first_path)

            second = Document()
            numbering_rid = next(
                rid
                for rid, relationship in second.part.rels.items()
                if relationship.reltype == format_paper_module.RT.NUMBERING
            )
            second.part.drop_rel(numbering_rid)
            custom_xml_part = second.part.rels.part_with_reltype(
                format_paper_module.RT.CUSTOM_XML
            )
            custom_xml_part.relate_to(
                format_paper_module.Part(
                    format_paper_module.PackURI("/word/NUMBERING.xml"),
                    "application/xml",
                    resource_payload,
                    second.part.package,
                ),
                resource_relationship_type,
            )
            second.add_paragraph("主文档")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            numbering_part = merged.part.rels.part_with_reltype(
                format_paper_module.RT.NUMBERING
            )
            self.assertNotEqual(
                str(numbering_part.partname).casefold(),
                "/word/numbering.xml",
            )
            merged_numbered = next(
                paragraph
                for paragraph in merged.paragraphs
                if paragraph.text == "插入文档编号"
            )
            self.assertIsNotNone(
                merged_numbered._element.find(
                    ".//" + qn("w:numId")
                )
            )
            with ZipFile(output_path) as archive:
                member_names = archive.namelist()
            self.assertEqual(
                len(member_names),
                len({name.casefold() for name in member_names}),
            )
            copied_resource = next(
                relationship.target_part
                for custom_relationship in merged.part.rels.values()
                if custom_relationship.reltype == format_paper_module.RT.CUSTOM_XML
                for relationship in custom_relationship.target_part.rels.values()
                if relationship.reltype == resource_relationship_type
            )
            self.assertEqual(copied_resource.blob, resource_payload)

    def test_concatenate_documents_remaps_sparse_footnote_ids_without_collision(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "first-footnote.docx"
            second_path = tmp / "second-footnote.docx"
            output_path = tmp / "combined-footnotes.docx"

            first = Document()
            first.add_paragraph("第一份文档")
            first.save(str(first_path))
            self.inject_simple_footnote(first_path, "第一份脚注", footnote_id=2)

            second = Document()
            second.add_paragraph("第二份文档")
            second.save(str(second_path))
            # 目标仅有 ID=4 时，docxcompose 原先用 len(root)+1 也会选到 4。
            self.inject_simple_footnote(second_path, "第二份脚注", footnote_id=4)

            result = concatenate_documents(
                str(first_path),
                str(second_path),
                str(output_path),
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            with ZipFile(output_path, "r") as archive:
                document_root = etree.fromstring(archive.read("word/document.xml"))
                footnotes_root = etree.fromstring(archive.read("word/footnotes.xml"))

            reference_ids = [
                reference.get(qn("w:id"))
                for reference in document_root.findall(".//" + qn("w:footnoteReference"))
            ]
            user_footnotes = [
                footnote
                for footnote in footnotes_root.findall(qn("w:footnote"))
                if footnote.get(qn("w:type")) is None
            ]
            footnote_ids = [footnote.get(qn("w:id")) for footnote in user_footnotes]
            footnote_texts = {
                "".join(footnote.itertext()).strip()
                for footnote in user_footnotes
            }

            self.assertEqual(len(reference_ids), 2)
            self.assertEqual(len(set(reference_ids)), 2)
            self.assertEqual(set(reference_ids), set(footnote_ids))
            self.assertEqual(footnote_texts, {"第一份脚注", "第二份脚注"})

    def test_concatenate_documents_migrates_comments_and_remaps_colliding_ids(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "first-comment.docx"
            second_path = tmp / "second-comment.docx"
            output_path = tmp / "combined-comments.docx"

            first = Document()
            first_run = first.add_paragraph().add_run("第一份被批注文字")
            first.add_comment(first_run, "第一份批注", author="审阅者")
            first.save(str(first_path))

            second = Document()
            second_run = second.add_paragraph().add_run("第二份被批注文字")
            second.add_comment(second_run, "第二份批注", author="审阅者")
            second.save(str(second_path))

            result = concatenate_documents(
                str(first_path),
                str(second_path),
                str(output_path),
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            with ZipFile(output_path, "r") as archive:
                document_root = etree.fromstring(archive.read("word/document.xml"))
                comments_root = etree.fromstring(archive.read("word/comments.xml"))
                relationships_xml = archive.read("word/_rels/document.xml.rels")

            comments = comments_root.findall(qn("w:comment"))
            comment_ids = [comment.get(qn("w:id")) for comment in comments]
            comment_texts = {"".join(comment.itertext()) for comment in comments}
            self.assertEqual(len(comment_ids), 2)
            self.assertEqual(len(set(comment_ids)), 2)
            self.assertEqual(comment_texts, {"第一份批注", "第二份批注"})
            self.assertIn(b"/comments", relationships_xml)
            reopened_comments = Document(str(output_path)).comments
            self.assertEqual(
                {comment.text for comment in reopened_comments},
                {"第一份批注", "第二份批注"},
            )

            for tag in (
                "w:commentRangeStart",
                "w:commentRangeEnd",
                "w:commentReference",
            ):
                semantic_ids = [
                    node.get(qn("w:id"))
                    for node in document_root.findall(".//" + qn(tag))
                ]
                self.assertEqual(len(semantic_ids), 2)
                self.assertEqual(set(semantic_ids), set(comment_ids))

    def test_concatenate_documents_migrates_comments_from_restored_header(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "header-comment-first.docx"
            second_path = tmp / "header-comment-second.docx"
            output_path = tmp / "header-comment-merged.docx"

            first = Document()
            header_run = first.sections[0].header.paragraphs[0].add_run(
                "页眉批注目标"
            )
            first.add_comment(header_run, "页眉批注内容", author="审阅者")
            first.add_paragraph("第一份文档")
            first.save(first_path)

            second = Document()
            second.add_paragraph("第二份文档")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            header_root = merged.sections[0].header._element
            header_ids = [
                node.get(qn("w:id"))
                for node in header_root.findall(".//" + qn("w:commentReference"))
            ]
            self.assertEqual(len(header_ids), 1)
            self.assertEqual(
                {comment.text for comment in merged.comments},
                {"页眉批注内容"},
            )
            with ZipFile(output_path) as archive:
                comments_root = etree.fromstring(archive.read("word/comments.xml"))
            self.assertEqual(
                header_ids,
                [
                    comment.get(qn("w:id"))
                    for comment in comments_root.findall(qn("w:comment"))
                ],
            )

    def test_concatenate_documents_migrates_endnotes_and_remaps_colliding_ids(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "first-endnote.docx"
            second_path = tmp / "second-endnote.docx"
            output_path = tmp / "combined-endnotes.docx"

            first = Document()
            first.add_paragraph("第一份文档")
            first.save(str(first_path))
            self.inject_simple_endnote(first_path, "第一份尾注", endnote_id=2)

            second = Document()
            second.add_paragraph("第二份文档")
            second.save(str(second_path))
            self.inject_simple_endnote(second_path, "第二份尾注", endnote_id=2)

            result = concatenate_documents(
                str(first_path),
                str(second_path),
                str(output_path),
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            with ZipFile(output_path, "r") as archive:
                document_root = etree.fromstring(archive.read("word/document.xml"))
                endnotes_root = etree.fromstring(archive.read("word/endnotes.xml"))
                relationships_xml = archive.read("word/_rels/document.xml.rels")

            reference_ids = [
                reference.get(qn("w:id"))
                for reference in document_root.findall(".//" + qn("w:endnoteReference"))
            ]
            user_endnotes = [
                endnote
                for endnote in endnotes_root.findall(qn("w:endnote"))
                if endnote.get(qn("w:type")) is None
            ]
            endnote_ids = [endnote.get(qn("w:id")) for endnote in user_endnotes]
            endnote_texts = {"".join(endnote.itertext()) for endnote in user_endnotes}

            self.assertEqual(len(reference_ids), 2)
            self.assertEqual(len(set(reference_ids)), 2)
            self.assertEqual(set(reference_ids), set(endnote_ids))
            self.assertEqual(endnote_texts, {"第一份尾注", "第二份尾注"})
            self.assertIn(b"/endnotes", relationships_xml)

    def test_concatenate_documents_rejects_modern_threaded_comments_without_data_loss(self):
        relationship_type = (
            "http://schemas.microsoft.com/office/2011/relationships/"
            "commentsExtended"
        )
        content_type = "application/vnd.ms-word.commentsExtended+xml"
        w14_namespace = "http://schemas.microsoft.com/office/word/2010/wordml"
        w15_namespace = "http://schemas.microsoft.com/office/word/2012/wordml"

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "threaded-comment-first.docx"
            second_path = tmp / "plain-second.docx"
            output_path = tmp / "must-not-be-written.docx"

            first = Document()
            commented_run = first.add_paragraph().add_run("现代批注目标")
            comment = first.add_comment(
                commented_run,
                "已解决的批注",
                author="审阅者",
            )
            comment.paragraphs[0]._element.set(
                f"{{{w14_namespace}}}paraId",
                "1234ABCD",
            )
            comments_extended = format_paper_module.Part(
                format_paper_module.PackURI(
                    "/word/commentsExtended.xml"
                ),
                content_type,
                (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    f'<w15:commentsEx xmlns:w15="{w15_namespace}">'
                    '<w15:commentEx w15:paraId="1234ABCD" w15:done="1"/>'
                    "</w15:commentsEx>"
                ).encode(),
                first.part.package,
            )
            first.part.relate_to(comments_extended, relationship_type)
            first.save(first_path)

            second = Document()
            second.add_paragraph("普通正文")
            second.save(second_path)

            with self.assertRaisesRegex(
                DocumentConcatError,
                "现代线程批注",
            ):
                concatenate_documents(
                    first_path,
                    second_path,
                    output_path,
                    restart_body_page_number=False,
                )
            self.assertFalse(output_path.exists())

    def test_concatenate_documents_rejects_classic_comment_into_modern_master(self):
        relationship_type = (
            "http://schemas.microsoft.com/office/2011/relationships/"
            "commentsExtended"
        )
        content_type = "application/vnd.ms-word.commentsExtended+xml"
        w14_namespace = "http://schemas.microsoft.com/office/word/2010/wordml"
        w15_namespace = "http://schemas.microsoft.com/office/word/2012/wordml"
        duplicate_para_id = "1234ABCD"

        def add_classic_comment(document, text):
            commented_run = document.add_paragraph().add_run(f"{text}目标")
            comment = document.add_comment(
                commented_run,
                text,
                author="审阅者",
            )
            comment.paragraphs[0]._element.set(
                f"{{{w14_namespace}}}paraId",
                duplicate_para_id,
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "classic-comment-first.docx"
            second_path = tmp / "modern-comment-master.docx"
            output_path = tmp / "must-not-be-written.docx"

            first = Document()
            add_classic_comment(first, "插入文档普通批注")
            first.save(first_path)

            second = Document()
            add_classic_comment(second, "目标文档现代批注")
            comments_extended = format_paper_module.Part(
                format_paper_module.PackURI(
                    "/word/commentsExtended.xml"
                ),
                content_type,
                (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    f'<w15:commentsEx xmlns:w15="{w15_namespace}">'
                    f'<w15:commentEx w15:paraId="{duplicate_para_id}" '
                    'w15:done="1"/>'
                    "</w15:commentsEx>"
                ).encode(),
                second.part.package,
            )
            second.part.relate_to(comments_extended, relationship_type)
            second.save(second_path)

            with self.assertRaisesRegex(
                DocumentConcatError,
                "目标文档包含现代线程批注",
            ):
                concatenate_documents(
                    first_path,
                    second_path,
                    output_path,
                    restart_body_page_number=False,
                )
            self.assertFalse(output_path.exists())

    def test_concatenate_documents_rejects_modern_comment_people_sidecar(self):
        relationship_type = (
            "http://schemas.microsoft.com/office/2011/relationships/people"
        )
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "people-sidecar-first.docx"
            second_path = tmp / "plain-second.docx"
            output_path = tmp / "must-not-be-written.docx"

            first = Document()
            first.add_paragraph("含现代批注人员元数据")
            first.part.relate_to(
                format_paper_module.Part(
                    format_paper_module.PackURI("/word/people.xml"),
                    "application/vnd.ms-word.people+xml",
                    (
                        b'<?xml version="1.0" encoding="UTF-8"?>'
                        b'<w15:people xmlns:w15="http://schemas.microsoft.com/'
                        b'office/word/2012/wordml"/>'
                    ),
                    first.part.package,
                ),
                relationship_type,
            )
            first.save(first_path)

            second = Document()
            second.save(second_path)

            with self.assertRaisesRegex(
                DocumentConcatError,
                "现代线程批注",
            ):
                concatenate_documents(
                    first_path,
                    second_path,
                    output_path,
                    restart_body_page_number=False,
                )
            self.assertFalse(output_path.exists())

    def test_concatenate_documents_preserves_endnote_separator_styles(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "styled-endnote-first.docx"
            second_path = tmp / "plain-second.docx"
            output_path = tmp / "styled-endnote-merged.docx"

            first = Document()
            separator_style = first.styles.add_style(
                "UniqueEndSeparator",
                format_paper_module.WD_STYLE_TYPE.PARAGRAPH,
            )
            first.add_paragraph("含自定义尾注分隔符")
            first.save(first_path)
            self.inject_simple_endnote(
                first_path,
                "尾注内容",
                separator_style=separator_style.style_id,
            )

            second = Document()
            second.add_paragraph("普通正文")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            with ZipFile(output_path) as archive:
                endnotes_root = etree.fromstring(archive.read("word/endnotes.xml"))
                styles_root = etree.fromstring(archive.read("word/styles.xml"))
            separator = next(
                note
                for note in endnotes_root.findall(qn("w:endnote"))
                if note.get(qn("w:type")) == "separator"
            )
            migrated_style_id = separator.find(
                ".//" + qn("w:pStyle")
            ).get(qn("w:val"))
            self.assertTrue(
                any(
                    style.get(qn("w:styleId")) == migrated_style_id
                    for style in styles_root.findall(qn("w:style"))
                )
            )

    def test_concatenate_documents_trims_cover_trailing_blank_paragraphs(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "cover.docx"
            second_path = tmp / "body.docx"
            output_path = tmp / "merged.docx"

            cover = Document()
            cover.add_paragraph("封面标题")
            for _ in range(4):
                cover.add_paragraph("")  # 尾部多余空段落
            cover.save(str(first_path))

            body = Document()
            body.add_paragraph("正文内容")
            body.save(str(second_path))

            result = concatenate_documents(str(first_path), str(second_path), str(output_path))
            self.assertTrue(result["concatenated"])

            merged_texts = [p.text for p in Document(str(output_path)).paragraphs]
            cover_idx = merged_texts.index("封面标题")
            body_idx = merged_texts.index("正文内容")
            # 封面与正文之间的空段落已被清理（仅余分节符产生的至多 1 个空段）
            gap_blanks = sum(1 for t in merged_texts[cover_idx + 1:body_idx] if not t.strip())
            self.assertLessEqual(gap_blanks, 1)

    def test_concatenate_documents_preserves_terminal_equation_paragraph(self):
        equation = parse_xml(
            r"""
            <m:oMathPara %s>
              <m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath>
            </m:oMathPara>
            """ % nsdecls("m")
        )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "equation-cover.docx"
            second_path = tmp / "body.docx"
            output_path = tmp / "merged.docx"

            cover = Document()
            cover.add_paragraph("公式封面")
            cover.add_paragraph()._element.append(equation)
            cover.save(str(first_path))

            body = Document()
            body.add_paragraph("正文内容")
            body.save(str(second_path))

            result = concatenate_documents(str(first_path), str(second_path), str(output_path))

            self.assertTrue(result["concatenated"])
            merged = Document(str(output_path))
            self.assertEqual(len(merged._element.findall(".//" + qn("m:oMath"))), 1)
            self.assertIn("正文内容", [paragraph.text for paragraph in merged.paragraphs])

    def test_concatenate_documents_propagates_output_limit(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "cover.docx"
            second_path = tmp / "body.docx"
            output_path = tmp / "merged.docx"
            Document().save(str(first_path))
            Document().save(str(second_path))

            with self.assertRaises(OutputSizeLimitExceeded):
                concatenate_documents(
                    str(first_path),
                    str(second_path),
                    str(output_path),
                    max_output_bytes=1024,
                )

            self.assertFalse(output_path.exists())

    def test_merge_cover_and_body_keeps_cover_and_formats_body(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            cover_path = tmp / "cover.docx"
            body_path = tmp / "body.docx"
            output_path = tmp / "merged.docx"

            cover = Document()
            cover.add_paragraph("课程论文封面")
            cover.save(str(cover_path))

            body = Document()
            body.add_paragraph("基于多元回归模型的城市化研究")
            body.add_paragraph("摘要：本文研究城市化进程。")
            body.add_paragraph("关键词：城市化 回归")
            body.add_paragraph("1 引言")
            body.add_paragraph("这是正文内容。")
            body.save(str(body_path))

            with patch.object(
                format_paper_module.tempfile,
                "NamedTemporaryFile",
                wraps=tempfile.NamedTemporaryFile,
            ) as named_temp_file:
                result = merge_cover_and_body(
                    str(cover_path),
                    str(body_path),
                    str(output_path),
                    max_output_bytes=64 * 1024,
                )

            self.assertTrue(result)
            self.assertTrue(output_path.exists())
            with ZipFile(output_path) as archive:
                self.assertIsNone(archive.testzip())
            self.assertEqual(
                Path(named_temp_file.call_args.kwargs["dir"]),
                body_path.parent,
            )

            merged_texts = [p.text for p in Document(str(output_path)).paragraphs]
            # 封面原样保留
            self.assertIn("课程论文封面", merged_texts)
            # 正文内容（经排版）仍在
            self.assertTrue(any("引言" in t for t in merged_texts))
            self.assertTrue(any("城市化研究" in t for t in merged_texts))

    def test_merge_cover_and_body_preserves_formatted_body_section_layout(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            cover_path = tmp / "cover-layout.docx"
            body_path = tmp / "body-layout.docx"
            output_path = tmp / "merged-layout.docx"

            cover = Document()
            cover_section = cover.sections[0]
            cover_section.left_margin = Cm(1)
            cover_section.header.paragraphs[0].text = "封面旧页眉"
            cover_section.footer.paragraphs[0].text = "封面旧页脚"
            cover.add_paragraph("课程论文封面")
            cover.save(str(cover_path))

            body = Document()
            body.add_paragraph("基于多元回归模型的城市化研究")
            body.add_paragraph("摘要：本文研究城市化进程。")
            body.add_paragraph("关键词：城市化 回归")
            body.add_paragraph("正文内容。")
            body.save(str(body_path))

            result = merge_cover_and_body(
                str(cover_path),
                str(body_path),
                str(output_path),
            )

            self.assertIsInstance(result, dict)
            merged = Document(str(output_path))
            self.assertEqual(len(merged.sections), 2)
            merged_cover, merged_body = merged.sections

            self.assertAlmostEqual(merged_cover.left_margin.cm, 1.0, places=1)
            self.assertEqual(
                merged_cover.header.paragraphs[0].text,
                "封面旧页眉",
            )
            self.assertAlmostEqual(merged_body.page_width.cm, 21.0, places=1)
            self.assertAlmostEqual(merged_body.page_height.cm, 29.7, places=1)
            self.assertAlmostEqual(merged_body.left_margin.cm, 3.18, places=1)
            self.assertEqual(
                merged_body.header.paragraphs[0].text,
                "基于多元回归模型的城市化研究",
            )
            self.assertIn("PAGE", merged_body.footer._element.xml)
            pg_num_type = merged_body._sectPr.find(qn("w:pgNumType"))
            self.assertIsNotNone(pg_num_type)
            self.assertEqual(pg_num_type.get(qn("w:start")), "1")

    def test_merge_cover_and_body_preserves_cover_odd_even_headers(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            cover_path = tmp / "odd-even-cover.docx"
            body_path = tmp / "body.docx"
            output_path = tmp / "odd-even-merged.docx"

            cover = Document()
            cover.settings.odd_and_even_pages_header_footer = True
            cover.sections[0].header.paragraphs[0].text = "COVER ODD"
            cover.sections[0].even_page_header.paragraphs[0].text = (
                "COVER EVEN"
            )
            cover.add_paragraph("封面内容")
            cover.save(cover_path)

            body = Document()
            body.add_paragraph("基于多元回归模型的城市化研究")
            body.add_paragraph("摘要：本文研究城市化进程。")
            body.add_paragraph("关键词：城市化 回归")
            body.add_paragraph("正文内容。")
            body.save(body_path)

            result = merge_cover_and_body(
                cover_path,
                body_path,
                output_path,
            )

            self.assertIsInstance(result, dict)
            merged = Document(output_path)
            self.assertTrue(
                merged.settings.odd_and_even_pages_header_footer
            )
            cover_section, body_section = merged.sections
            self.assertEqual(
                cover_section.header.paragraphs[0].text,
                "COVER ODD",
            )
            self.assertEqual(
                cover_section.even_page_header.paragraphs[0].text,
                "COVER EVEN",
            )
            self.assertEqual(
                body_section.even_page_header.paragraphs[0].text,
                "基于多元回归模型的城市化研究",
            )
            self.assertIn(
                "PAGE",
                body_section.even_page_footer._element.xml,
            )

    def test_merge_cover_and_body_creates_missing_comment_and_endnote_parts(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            cover_path = tmp / "annotated-cover.docx"
            body_path = tmp / "plain-body.docx"
            output_path = tmp / "annotated-merge.docx"

            cover = Document()
            commented_run = cover.add_paragraph().add_run("封面批注目标")
            cover.add_comment(commented_run, "封面批注内容", author="审阅者")
            cover.save(str(cover_path))
            self.inject_simple_endnote(cover_path, "封面尾注内容", endnote_id=2)

            body = Document()
            body.add_paragraph("正文内容")
            body.save(str(body_path))

            result = merge_cover_and_body(
                str(cover_path),
                str(body_path),
                str(output_path),
            )

            self.assertIsInstance(result, dict)
            with ZipFile(output_path, "r") as archive:
                self.assertIn("word/comments.xml", archive.namelist())
                self.assertIn("word/endnotes.xml", archive.namelist())
                document_root = etree.fromstring(archive.read("word/document.xml"))
                comments_root = etree.fromstring(archive.read("word/comments.xml"))
                endnotes_root = etree.fromstring(archive.read("word/endnotes.xml"))

            self.assertEqual(
                len(document_root.findall(".//" + qn("w:commentReference"))),
                1,
            )
            self.assertEqual(
                len(document_root.findall(".//" + qn("w:endnoteReference"))),
                1,
            )
            self.assertIn("封面批注内容", "".join(comments_root.itertext()))
            self.assertIn("封面尾注内容", "".join(endnotes_root.itertext()))

    def test_merge_cover_and_body_propagates_intermediate_output_limit(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            cover_path = tmp / "cover.docx"
            body_path = tmp / "body.docx"
            output_path = tmp / "merged.docx"
            Document().save(str(cover_path))
            Document().save(str(body_path))

            with self.assertRaises(OutputSizeLimitExceeded):
                merge_cover_and_body(
                    str(cover_path),
                    str(body_path),
                    str(output_path),
                    max_output_bytes=1024,
                )

            self.assertFalse(output_path.exists())

    def test_merge_cover_and_body_leases_and_releases_intermediate_file(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            cover_path = tmp / "cover.docx"
            body_path = tmp / "body.docx"
            output_path = tmp / "merged.docx"
            Document().save(str(cover_path))
            Document().save(str(body_path))
            leases_before = format_paper_module.get_active_temp_paths()
            observed_paths = []

            def reject_formatted_body(_body_path, formatted_path, **_kwargs):
                formatted_path = Path(formatted_path)
                observed_paths.append(formatted_path)
                self.assertTrue(format_paper_module.is_active_temp_path(formatted_path))
                formatted_path.write_bytes(b"partial-intermediate")
                return False

            with patch.object(
                format_paper_module,
                "format_academic_paper",
                side_effect=reject_formatted_body,
            ):
                result = merge_cover_and_body(
                    str(cover_path),
                    str(body_path),
                    str(output_path),
                )

            self.assertFalse(result)
            self.assertEqual(len(observed_paths), 1)
            self.assertFalse(observed_paths[0].exists())
            self.assertEqual(format_paper_module.get_active_temp_paths(), leases_before)

    def test_merge_cover_and_body_propagates_final_output_limit(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            cover_path = tmp / "cover.docx"
            body_path = tmp / "body.docx"
            output_path = tmp / "merged.docx"
            Document().save(str(cover_path))
            Document().save(str(body_path))

            def fake_body_formatter(_body_path, formatted_path, **_kwargs):
                formatted = Document()
                formatted.add_paragraph("已排版正文")
                formatted.save(formatted_path)
                return {"stats": {}}

            with patch.object(
                format_paper_module,
                "format_academic_paper",
                side_effect=fake_body_formatter,
            ):
                with self.assertRaises(OutputSizeLimitExceeded):
                    merge_cover_and_body(
                        str(cover_path),
                        str(body_path),
                        str(output_path),
                        max_output_bytes=1024,
                    )

            self.assertFalse(output_path.exists())

    def test_concatenate_documents_raises_for_invalid_docx(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "cover.docx"
            second_path = tmp / "broken.docx"
            output_path = tmp / "merged.docx"

            Document().save(str(first_path))
            second_path.write_bytes(b"this is not a docx zip")

            with self.assertRaises(DocumentConcatError):
                concatenate_documents(str(first_path), str(second_path), str(output_path))

    def test_concatenate_documents_maps_lazy_ooxml_errors_for_both_inputs(self):
        def set_invalid_even_and_odd_headers_value(xml_bytes):
            root = etree.fromstring(xml_bytes)
            setting = root.find(qn("w:evenAndOddHeaders"))
            if setting is None:
                setting = etree.Element(qn("w:evenAndOddHeaders"))
                root.append(setting)
            setting.set(qn("w:val"), "not-a-boolean")
            return etree.tostring(
                root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            for malformed_position in ("first", "second"):
                with self.subTest(malformed_position=malformed_position):
                    first_path = tmp / f"{malformed_position}-first.docx"
                    second_path = tmp / f"{malformed_position}-second.docx"
                    output_path = tmp / f"{malformed_position}-merged.docx"
                    for path, text in (
                        (first_path, "封面内容"),
                        (second_path, "正文内容"),
                    ):
                        document = Document()
                        document.add_paragraph(text)
                        document.save(path)
                    malformed_path = (
                        first_path if malformed_position == "first" else second_path
                    )
                    self.rewrite_docx_member(
                        malformed_path,
                        "word/settings.xml",
                        set_invalid_even_and_odd_headers_value,
                    )

                    with self.assertRaisesRegex(
                        DocumentConcatError,
                        "重新导出",
                    ) as caught:
                        concatenate_documents(
                            first_path,
                            second_path,
                            output_path,
                            restart_body_page_number=False,
                        )

                    self.assertIsInstance(
                        caught.exception.__cause__,
                        format_paper_module.InvalidXmlError,
                    )
                    self.assertFalse(output_path.exists())

    def test_concatenate_documents_can_keep_original_page_numbers(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "cover.docx"
            second_path = tmp / "body.docx"
            output_path = tmp / "merged.docx"

            Document().save(str(first_path))
            body = Document()
            body.add_paragraph("正文内容")
            body.save(str(second_path))

            result = concatenate_documents(
                str(first_path), str(second_path), str(output_path),
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            self.assertFalse(result["page_number_restarted"])
            self.assertTrue(output_path.exists())

    def test_concatenate_documents_preserves_conflicting_docproperty_fields(self):
        def append_property_field(
            paragraph,
            property_name: str,
            cached_value: str,
            *,
            quoted: bool = True,
        ):
            property_token = f'"{property_name}"' if quoted else property_name
            paragraph._element.append(
                parse_xml(
                    f'<w:fldSimple {nsdecls("w")} '
                    f"w:instr=' DOCPROPERTY {property_token} \\* MERGEFORMAT '>"
                    f"<w:r><w:t>{cached_value}</w:t></w:r>"
                    "</w:fldSimple>"
                )
            )

        def append_split_property_field(
            paragraph,
            property_name: str,
            cached_value: str,
        ):
            for fragment in (
                f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="begin"/></w:r>',
                f'<w:r {nsdecls("w")}><w:instrText xml:space="preserve"> DOCPROPERTY </w:instrText></w:r>',
                f'<w:r {nsdecls("w")}><w:instrText>"{property_name}"</w:instrText></w:r>',
                f'<w:r {nsdecls("w")}><w:instrText xml:space="preserve"> \\* MERGEFORMAT </w:instrText></w:r>',
                f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="separate"/></w:r>',
                f'<w:r {nsdecls("w")}><w:t>{cached_value}</w:t></w:r>',
                f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="end"/></w:r>',
            ):
                paragraph._element.append(parse_xml(fragment))

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "cover-properties.docx"
            second_path = tmp / "body-properties.docx"
            output_path = tmp / "merged-properties.docx"

            cover = Document()
            cover_properties = CustomProperties(cover)
            cover_properties.add("SharedCode", "COVER")
            cover_properties.add("CoverCount", 7)
            append_property_field(
                cover.add_paragraph("封面属性："),
                "SharedCode",
                "COVER",
                quoted=False,
            )
            append_split_property_field(
                cover.sections[0].header.paragraphs[0],
                "SharedCode",
                "COVER HEADER",
            )
            append_split_property_field(
                cover.sections[0].footer.paragraphs[0],
                "SharedCode",
                "COVER FOOTER",
            )
            comment_run = cover.add_paragraph().add_run("封面批注目标")
            cover_comment = cover.add_comment(
                comment_run,
                "封面批注属性：",
                author="审阅者",
            )
            append_property_field(
                cover_comment.paragraphs[0],
                "SharedCode",
                "COVER COMMENT",
            )
            cover.save(first_path)
            self.inject_simple_footnote(
                first_path,
                "COVER FOOTNOTE",
                footnote_content_xml=(
                    f'<w:fldSimple {nsdecls("w")} '
                    'w:instr=\' DOCPROPERTY "SharedCode" \\* MERGEFORMAT \'> '
                    "<w:r><w:t>COVER FOOTNOTE</w:t></w:r>"
                    "</w:fldSimple>"
                ),
            )

            body = Document()
            CustomProperties(body).add("SharedCode", "BODY")
            append_property_field(
                body.add_paragraph("正文属性："),
                "SharedCode",
                "BODY",
            )
            body.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            body_instructions = [
                field.get(qn("w:instr"))
                for field in merged.element.body.findall(".//" + qn("w:fldSimple"))
            ]
            self.assertEqual(len(body_instructions), 2)
            self.assertTrue(
                any('DOCPROPERTY "SharedCode"' in value for value in body_instructions)
            )
            properties = dict(CustomProperties(merged).items())
            renamed_name = next(
                name
                for name in properties
                if name.startswith("SharedCode__inserted_")
            )
            renamed_instruction = next(
                value
                for value in body_instructions
                if renamed_name in value
            )
            self.assertIn(f"DOCPROPERTY {renamed_name}", renamed_instruction)
            self.assertNotIn(f'"{renamed_name}"', renamed_instruction)

            cover_header_instruction = "".join(
                node.text or ""
                for node in merged.sections[0].header._element.findall(
                    ".//" + qn("w:instrText")
                )
            )
            self.assertIn(f'DOCPROPERTY "{renamed_name}"', cover_header_instruction)
            cover_footer_instruction = "".join(
                node.text or ""
                for node in merged.sections[0].footer._element.findall(
                    ".//" + qn("w:instrText")
                )
            )
            self.assertIn(f'DOCPROPERTY "{renamed_name}"', cover_footer_instruction)
            with ZipFile(output_path) as archive:
                comments_root = etree.fromstring(archive.read("word/comments.xml"))
                footnotes_root = etree.fromstring(archive.read("word/footnotes.xml"))
            comment_instructions = [
                field.get(qn("w:instr"))
                for field in comments_root.findall(".//" + qn("w:fldSimple"))
            ]
            footnote_instructions = [
                field.get(qn("w:instr"))
                for field in footnotes_root.findall(".//" + qn("w:fldSimple"))
            ]
            self.assertEqual(
                comment_instructions,
                [f' DOCPROPERTY "{renamed_name}" \\* MERGEFORMAT '],
            )
            self.assertEqual(
                footnote_instructions,
                [f' DOCPROPERTY "{renamed_name}" \\* MERGEFORMAT '],
            )

            self.assertEqual(properties["SharedCode"], "BODY")
            self.assertEqual(properties[renamed_name], "COVER")
            self.assertEqual(properties["CoverCount"], 7)
            with ZipFile(output_path) as archive:
                self.assertIsNone(archive.testzip())
                package_relationships = etree.fromstring(
                    archive.read("_rels/.rels")
                )
                custom_relationships = [
                    relationship
                    for relationship in package_relationships
                    if relationship.get("Type")
                    == format_paper_module.RT.CUSTOM_PROPERTIES
                ]
                self.assertEqual(len(custom_relationships), 1)
                custom_member = custom_relationships[0].get("Target").lstrip("/")
                self.assertIn(custom_member, archive.namelist())

                content_types = etree.fromstring(
                    archive.read("[Content_Types].xml")
                )
                custom_overrides = [
                    override
                    for override in content_types.findall(
                        f"{{{self.CONTENT_TYPES_NS}}}Override"
                    )
                    if override.get("PartName") == f"/{custom_member}"
                    and override.get("ContentType")
                    == format_paper_module.CT.OFC_CUSTOM_PROPERTIES
                ]
                self.assertEqual(len(custom_overrides), 1)
                properties_root = etree.fromstring(
                    archive.read(custom_member)
                )
            self.assertEqual(
                [item.get("pid") for item in properties_root],
                ["2", "3", "4"],
            )
            cover_count = next(
                item
                for item in properties_root
                if item.get("name") == "CoverCount"
            )
            self.assertEqual(cover_count[0].tag.rsplit("}", 1)[-1], "i4")
            self.assertEqual(cover_count[0].text, "7")

    def test_concatenate_documents_avoids_case_collisions_for_custom_properties_part(self):
        resource_relationship_type = "https://example.test/uppercase-custom-props"
        resource_payload = b"<resource>NOT CUSTOM PROPERTIES</resource>"
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "source-properties.docx"
            second_path = tmp / "uppercase-resource.docx"
            output_path = tmp / "custom-properties-case-safe.docx"

            first = Document()
            CustomProperties(first).add("InsertedCode", "SOURCE")
            first.add_paragraph("带自定义属性的文档")
            first.save(first_path)

            second = Document()
            custom_xml_part = second.part.rels.part_with_reltype(
                format_paper_module.RT.CUSTOM_XML
            )
            uppercase_resource = format_paper_module.Part(
                format_paper_module.PackURI("/docProps/CUSTOM.xml"),
                "application/xml",
                resource_payload,
                second.part.package,
            )
            custom_xml_part.relate_to(
                uppercase_resource,
                resource_relationship_type,
            )
            second.add_paragraph("主文档")
            second.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            with ZipFile(output_path) as archive:
                member_names = archive.namelist()
            self.assertEqual(
                len(member_names),
                len({name.casefold() for name in member_names}),
            )

            merged = Document(output_path)
            self.assertEqual(
                dict(CustomProperties(merged).items())["InsertedCode"],
                "SOURCE",
            )
            custom_properties_part = merged.part.package.part_related_by(
                format_paper_module.RT.CUSTOM_PROPERTIES
            )
            self.assertNotEqual(
                str(custom_properties_part.partname).casefold(),
                "/docprops/custom.xml",
            )
            copied_resource = next(
                relationship.target_part
                for custom_relationship in merged.part.rels.values()
                if custom_relationship.reltype == format_paper_module.RT.CUSTOM_XML
                for relationship in custom_relationship.target_part.rels.values()
                if relationship.reltype == resource_relationship_type
            )
            self.assertEqual(copied_resource.blob, resource_payload)

    def test_merge_cover_and_body_preserves_inserted_docproperty_field(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            cover_path = tmp / "cover-property.docx"
            body_path = tmp / "body-property.docx"
            output_path = tmp / "merged-property.docx"

            cover = Document()
            CustomProperties(cover).add("CoverCode", "DYNAMIC-COVER")
            property_paragraph = cover.add_paragraph("封面编号：")
            property_paragraph._element.append(
                parse_xml(
                    f'<w:fldSimple {nsdecls("w")} '
                    'w:instr=\' DOCPROPERTY "CoverCode" \\* MERGEFORMAT \'>'
                    "<w:r><w:t>DYNAMIC-COVER</w:t></w:r>"
                    "</w:fldSimple>"
                )
            )
            cover.save(cover_path)

            body = Document()
            body.add_paragraph("属性字段合并测试")
            body.add_paragraph("摘要：验证动态属性字段。")
            body.add_paragraph("关键词：属性 字段")
            body.add_paragraph("正文内容。")
            body.save(body_path)

            result = merge_cover_and_body(cover_path, body_path, output_path)

            self.assertIsInstance(result, dict)
            merged = Document(output_path)
            fields = merged.element.body.findall(".//" + qn("w:fldSimple"))
            self.assertEqual(len(fields), 1)
            self.assertIn("DOCPROPERTY", fields[0].get(qn("w:instr")))
            self.assertEqual(
                dict(CustomProperties(merged).items())["CoverCode"],
                "DYNAMIC-COVER",
            )

    def test_concatenate_documents_deduplicates_custom_properties_across_prefixes(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            first_path = tmp / "prefix-cover.docx"
            second_path = tmp / "prefix-body.docx"
            output_path = tmp / "prefix-merged.docx"

            cover = Document()
            CustomProperties(cover).add("SharedCount", 7)
            cover.save(first_path)
            with ZipFile(first_path, "r") as source:
                entries = source.infolist()
                payloads = {
                    entry.filename: source.read(entry.filename)
                    for entry in entries
                }
            custom_xml = payloads["docProps/custom.xml"]
            custom_xml = custom_xml.replace(b"xmlns:vt=", b"xmlns:v=")
            custom_xml = custom_xml.replace(b"<vt:", b"<v:")
            custom_xml = custom_xml.replace(b"</vt:", b"</v:")
            payloads["docProps/custom.xml"] = custom_xml
            with ZipFile(first_path, "w") as target:
                for entry in entries:
                    target.writestr(entry, payloads[entry.filename])

            body = Document()
            CustomProperties(body).add("SharedCount", 7)
            body.add_paragraph("正文")
            body.save(second_path)

            result = concatenate_documents(
                first_path,
                second_path,
                output_path,
                restart_body_page_number=False,
            )

            self.assertTrue(result["concatenated"])
            merged = Document(output_path)
            self.assertEqual(
                dict(CustomProperties(merged).items())["SharedCount"],
                7,
            )
            with ZipFile(output_path) as archive:
                properties_root = etree.fromstring(
                    archive.read("docProps/custom.xml")
                )
            self.assertEqual(
                [item.get("name") for item in properties_root],
                ["SharedCount"],
            )

    def test_concatenate_documents_returns_false_for_missing_file(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            existing = tmp / "exists.docx"
            Document().save(str(existing))

            result = concatenate_documents(
                str(existing), str(tmp / "missing.docx"), str(tmp / "out.docx")
            )

            self.assertFalse(result)

    def test_expired_temp_cleanup_is_idempotent_when_path_disappears(self):
        """A concurrent worker may remove a staged file after enumeration."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            missing_path = Path(tmp_dir) / ".docx-output-raced.tmp"
            self.assertFalse(
                format_paper_module.remove_expired_inactive_temp_path(
                    missing_path,
                    cutoff=float("inf"),
                )
            )


if __name__ == "__main__":
    unittest.main()
