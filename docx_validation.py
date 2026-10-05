"""Shared, side-effect-free DOCX package validation.

This module intentionally depends only on the Python standard library so both
Web uploads and direct/CLI formatter entry points enforce the same OOXML and
ZIP resource budgets.
"""

import posixpath
import re
import stat
import zipfile
import zlib
import xml.etree.ElementTree as ElementTree
from contextlib import suppress
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from urllib.parse import unquote, urlsplit

REQUIRED_DOCX_MEMBERS = ("[Content_Types].xml", "_rels/.rels")
MAX_DOCX_ARCHIVE_MEMBERS = 2048
MAX_DOCX_MEMBER_UNCOMPRESSED_BYTES = 50 * 1024 * 1024
MAX_DOCX_TOTAL_UNCOMPRESSED_BYTES = 64 * 1024 * 1024
MAX_DOCX_XML_MEMBER_UNCOMPRESSED_BYTES = 8 * 1024 * 1024
MAX_DOCX_TOTAL_XML_UNCOMPRESSED_BYTES = 16 * 1024 * 1024
MAX_DOCX_CONTENT_TYPES_BYTES = 256 * 1024
MAX_DOCX_RELATIONSHIP_PART_BYTES = 1 * 1024 * 1024
MAX_DOCX_TOTAL_RELATIONSHIP_BYTES = 4 * 1024 * 1024
MAX_DOCX_RELATIONSHIPS_PER_PART = 4096
MAX_DOCX_TOTAL_RELATIONSHIPS = 8192
# python-docx 1.2.0 以递归 DFS 加载 OPC 关系图；限制可达链深度，避免
# 小体积文档通过近千层关系链耗尽 Python 调用栈。
MAX_DOCX_RELATIONSHIP_GRAPH_DEPTH = 256
FORBIDDEN_XML_DECLARATION_MARKERS = (b"<!doctype", b"<!entity")
MAX_DOCX_XML_ELEMENTS = 200_000
MAX_DOCX_XML_DEPTH = 256
MAX_DOCX_TOTAL_XML_ELEMENTS = 500_000
# 段落会在 python-docx 中物化为 XML 节点包装、Paragraph 对象和
# 格式分析对象。单独的 XML 元素预算无法阻止数万个空段落被
# 高压缩后以极小归档进入处理链路，因此需要独立的结构预算。
MAX_DOCX_TOTAL_PARAGRAPHS = 10_000
# 单个段落同样可以包含数万个空 w:r；格式化会为每个 run 扩展
# 字体属性。50,000 个 run 足以容纳复杂合法论文，同时让两个并发
# 任务仍保持在小型实例的内存预算内。
MAX_DOCX_TOTAL_RUNS = 50_000
# XML 已受单部件 8 MiB、XML 总量 16 MiB、深度与元素数预算约束；
# 合法表格/重复文本可有很高压缩率，因此比率门槛只用于非 XML 成员。
MAX_DOCX_NON_XML_COMPRESSION_RATIO = 200
MAX_DOCX_MEMBER_NAME_LENGTH = 512
MAX_DOCX_TABLE_GRID_COLUMNS = 256
MAX_DOCX_TABLE_LOGICAL_CELLS = 20_000
DOCX_MAIN_DOCUMENT_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)
# 合并/拼接最多汇集两个已通过上传预算的包；产物校验需覆盖该合法峰值。
MAX_GENERATED_DOCX_ARCHIVE_MEMBERS = MAX_DOCX_ARCHIVE_MEMBERS * 2
MAX_GENERATED_DOCX_MEMBER_UNCOMPRESSED_BYTES = MAX_DOCX_MEMBER_UNCOMPRESSED_BYTES * 2
MAX_GENERATED_DOCX_TOTAL_UNCOMPRESSED_BYTES = MAX_DOCX_TOTAL_UNCOMPRESSED_BYTES * 2
MAX_GENERATED_DOCX_CONTENT_TYPES_BYTES = MAX_DOCX_CONTENT_TYPES_BYTES * 2
MAX_GENERATED_DOCX_XML_MEMBER_UNCOMPRESSED_BYTES = (
    MAX_DOCX_XML_MEMBER_UNCOMPRESSED_BYTES * 2
)
MAX_GENERATED_DOCX_TOTAL_XML_UNCOMPRESSED_BYTES = (
    MAX_DOCX_TOTAL_XML_UNCOMPRESSED_BYTES * 2
)
MAX_GENERATED_DOCX_RELATIONSHIP_PART_BYTES = MAX_DOCX_RELATIONSHIP_PART_BYTES * 2
MAX_GENERATED_DOCX_TOTAL_RELATIONSHIP_BYTES = MAX_DOCX_TOTAL_RELATIONSHIP_BYTES * 2
MAX_GENERATED_DOCX_RELATIONSHIPS_PER_PART = MAX_DOCX_RELATIONSHIPS_PER_PART * 2
MAX_GENERATED_DOCX_TOTAL_RELATIONSHIPS = MAX_DOCX_TOTAL_RELATIONSHIPS * 2
MAX_GENERATED_DOCX_DOCUMENT_XML_ELEMENTS = MAX_DOCX_TOTAL_XML_ELEMENTS * 2
MAX_GENERATED_DOCX_TABLE_LOGICAL_CELLS = MAX_DOCX_TABLE_LOGICAL_CELLS * 2
# 拼接/合并会汇集两份已通过输入预算的文档，并可能额外生成
# 封面、目录或分节段落；小幅余量避免合法的双文档产物在最后校验时被拒绝。
MAX_GENERATED_DOCX_TOTAL_PARAGRAPHS = MAX_DOCX_TOTAL_PARAGRAPHS * 2 + 100
MAX_GENERATED_DOCX_TOTAL_RUNS = MAX_DOCX_TOTAL_RUNS * 2 + 1_000
MAX_INPUT_DOCX_ARCHIVE_BYTES = 50 * 1024 * 1024
_ALLOWED_DOCX_COMPRESSION_TYPES = frozenset(
    (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED)
)
_ZIP_LOCAL_FILE_HEADER_SIGNATURE = b"PK\x03\x04"
_ZIP_DATA_DESCRIPTOR_SIGNATURE = b"PK\x07\x08"
_ZIP64_EXTRA_FIELD_ID = 0x0001
_ZIP_UINT32_MAX = (1 << 32) - 1
_ZIP_MEMBER_VALIDATION_CHUNK_BYTES = 64 * 1024
_ZIP_DATA_DESCRIPTOR_FLAG = 0x0008
_ZIP_UTF8_FILENAME_FLAG = 0x0800
_ZIP_DEFLATE_OPTION_FLAGS = 0x0006


@dataclass(frozen=True, slots=True)
class DocxValidationLimits:
    """All resource budgets applied while validating one DOCX package."""

    max_archive_bytes: int | None
    max_archive_members: int
    max_member_uncompressed_bytes: int
    max_total_uncompressed_bytes: int
    max_xml_member_uncompressed_bytes: int
    max_total_xml_uncompressed_bytes: int
    max_content_types_bytes: int
    max_relationship_part_bytes: int
    max_total_relationship_bytes: int
    max_relationships_per_part: int
    max_total_relationships: int
    max_relationship_graph_depth: int
    max_xml_elements: int
    max_xml_depth: int
    max_total_xml_elements: int
    max_total_paragraphs: int
    max_total_runs: int
    max_non_xml_compression_ratio: int | None
    max_member_name_length: int
    max_table_grid_columns: int
    max_table_logical_cells: int


def _valid_docx_validation_limits(limits: DocxValidationLimits) -> bool:
    """Reject malformed custom budgets before they reach arithmetic/parsers.

    The public validator accepts a limits profile, so callers can supply a
    dataclass produced with ``dataclasses.replace``.  A negative budget (or a
    boolean, which is an ``int`` subclass) could otherwise make a later
    subtraction/division raise unexpectedly instead of returning ``False``.
    """
    if not isinstance(limits, DocxValidationLimits):
        return False
    fields = (
        "max_archive_members",
        "max_member_uncompressed_bytes",
        "max_total_uncompressed_bytes",
        "max_xml_member_uncompressed_bytes",
        "max_total_xml_uncompressed_bytes",
        "max_content_types_bytes",
        "max_relationship_part_bytes",
        "max_total_relationship_bytes",
        "max_relationships_per_part",
        "max_total_relationships",
        "max_relationship_graph_depth",
        "max_xml_elements",
        "max_xml_depth",
        "max_total_xml_elements",
        "max_total_paragraphs",
        "max_total_runs",
        "max_member_name_length",
        "max_table_grid_columns",
        "max_table_logical_cells",
    )
    for field_name in fields:
        value = getattr(limits, field_name)
        if isinstance(value, bool) or not isinstance(value, int) or value < 0:
            return False
    if limits.max_relationship_graph_depth < 1 or limits.max_xml_depth < 1:
        return False
    if limits.max_archive_bytes is not None and (
        isinstance(limits.max_archive_bytes, bool)
        or not isinstance(limits.max_archive_bytes, int)
        or limits.max_archive_bytes < 0
    ):
        return False
    if limits.max_non_xml_compression_ratio is not None and (
        isinstance(limits.max_non_xml_compression_ratio, bool)
        or not isinstance(limits.max_non_xml_compression_ratio, int)
        or limits.max_non_xml_compression_ratio < 1
    ):
        return False
    return True


INPUT_DOCX_LIMITS = DocxValidationLimits(
    max_archive_bytes=MAX_INPUT_DOCX_ARCHIVE_BYTES,
    max_archive_members=MAX_DOCX_ARCHIVE_MEMBERS,
    max_member_uncompressed_bytes=MAX_DOCX_MEMBER_UNCOMPRESSED_BYTES,
    max_total_uncompressed_bytes=MAX_DOCX_TOTAL_UNCOMPRESSED_BYTES,
    max_xml_member_uncompressed_bytes=MAX_DOCX_XML_MEMBER_UNCOMPRESSED_BYTES,
    max_total_xml_uncompressed_bytes=MAX_DOCX_TOTAL_XML_UNCOMPRESSED_BYTES,
    max_content_types_bytes=MAX_DOCX_CONTENT_TYPES_BYTES,
    max_relationship_part_bytes=MAX_DOCX_RELATIONSHIP_PART_BYTES,
    max_total_relationship_bytes=MAX_DOCX_TOTAL_RELATIONSHIP_BYTES,
    max_relationships_per_part=MAX_DOCX_RELATIONSHIPS_PER_PART,
    max_total_relationships=MAX_DOCX_TOTAL_RELATIONSHIPS,
    max_relationship_graph_depth=MAX_DOCX_RELATIONSHIP_GRAPH_DEPTH,
    max_xml_elements=MAX_DOCX_XML_ELEMENTS,
    max_xml_depth=MAX_DOCX_XML_DEPTH,
    max_total_xml_elements=MAX_DOCX_TOTAL_XML_ELEMENTS,
    max_total_paragraphs=MAX_DOCX_TOTAL_PARAGRAPHS,
    max_total_runs=MAX_DOCX_TOTAL_RUNS,
    max_non_xml_compression_ratio=MAX_DOCX_NON_XML_COMPRESSION_RATIO,
    max_member_name_length=MAX_DOCX_MEMBER_NAME_LENGTH,
    max_table_grid_columns=MAX_DOCX_TABLE_GRID_COLUMNS,
    max_table_logical_cells=MAX_DOCX_TABLE_LOGICAL_CELLS,
)

GENERATED_DOCX_LIMITS = DocxValidationLimits(
    max_archive_bytes=None,
    max_archive_members=MAX_GENERATED_DOCX_ARCHIVE_MEMBERS,
    max_member_uncompressed_bytes=MAX_GENERATED_DOCX_MEMBER_UNCOMPRESSED_BYTES,
    max_total_uncompressed_bytes=MAX_GENERATED_DOCX_TOTAL_UNCOMPRESSED_BYTES,
    max_xml_member_uncompressed_bytes=MAX_GENERATED_DOCX_XML_MEMBER_UNCOMPRESSED_BYTES,
    max_total_xml_uncompressed_bytes=MAX_GENERATED_DOCX_TOTAL_XML_UNCOMPRESSED_BYTES,
    max_content_types_bytes=MAX_GENERATED_DOCX_CONTENT_TYPES_BYTES,
    max_relationship_part_bytes=MAX_GENERATED_DOCX_RELATIONSHIP_PART_BYTES,
    max_total_relationship_bytes=MAX_GENERATED_DOCX_TOTAL_RELATIONSHIP_BYTES,
    max_relationships_per_part=MAX_GENERATED_DOCX_RELATIONSHIPS_PER_PART,
    max_total_relationships=MAX_GENERATED_DOCX_TOTAL_RELATIONSHIPS,
    max_relationship_graph_depth=MAX_DOCX_RELATIONSHIP_GRAPH_DEPTH,
    max_xml_elements=MAX_GENERATED_DOCX_DOCUMENT_XML_ELEMENTS,
    max_xml_depth=MAX_DOCX_XML_DEPTH,
    max_total_xml_elements=MAX_GENERATED_DOCX_DOCUMENT_XML_ELEMENTS,
    max_total_paragraphs=MAX_GENERATED_DOCX_TOTAL_PARAGRAPHS,
    max_total_runs=MAX_GENERATED_DOCX_TOTAL_RUNS,
    max_non_xml_compression_ratio=None,
    max_member_name_length=MAX_DOCX_MEMBER_NAME_LENGTH,
    max_table_grid_columns=MAX_DOCX_TABLE_GRID_COLUMNS,
    max_table_logical_cells=MAX_GENERATED_DOCX_TABLE_LOGICAL_CELLS,
)
WORDPROCESSINGML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PACKAGE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
STRICT_OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships"
)
DRAWINGML_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main"
CORE_PROPERTIES_NAMESPACE = (
    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
)
EXTENDED_PROPERTIES_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
)
CUSTOM_PROPERTIES_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
)
WORD_DOCUMENT_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}document"
WORD_BACKGROUND_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}background"
WORD_BODY_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}body"
WORD_STYLES_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}styles"
WORD_NUMBERING_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}numbering"
WORD_SETTINGS_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}settings"
WORD_FOOTNOTES_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}footnotes"
WORD_ENDNOTES_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}endnotes"
WORD_HEADER_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}hdr"
WORD_FOOTER_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}ftr"
WORD_COMMENTS_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}comments"
WORD_GLOSSARY_DOCUMENT_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}glossaryDocument"
WORD_FONTS_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}fonts"
WORD_WEB_SETTINGS_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}webSettings"
PACKAGE_RELATIONSHIPS_TAG = (
    f"{{{PACKAGE_RELATIONSHIPS_NAMESPACE}}}Relationships"
)
DRAWING_THEME_TAG = f"{{{DRAWINGML_NAMESPACE}}}theme"
CORE_PROPERTIES_TAG = f"{{{CORE_PROPERTIES_NAMESPACE}}}coreProperties"
EXTENDED_PROPERTIES_TAG = f"{{{EXTENDED_PROPERTIES_NAMESPACE}}}Properties"
CUSTOM_PROPERTIES_TAG = f"{{{CUSTOM_PROPERTIES_NAMESPACE}}}Properties"
WORD_TABLE_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}tbl"
WORD_PARAGRAPH_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}p"
WORD_RUN_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}r"
WORD_TABLE_ROW_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}tr"
WORD_TABLE_CELL_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}tc"
WORD_TABLE_CELL_PROPERTIES_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}tcPr"
WORD_GRID_SPAN_TAG = f"{{{WORDPROCESSINGML_NAMESPACE}}}gridSpan"
WORD_VALUE_ATTRIBUTE = f"{{{WORDPROCESSINGML_NAMESPACE}}}val"
ASCII_PART_NAME_CASE_TABLE = str.maketrans(
    "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
    "abcdefghijklmnopqrstuvwxyz",
)
PACKAGE_RELATIONSHIP_CONTENT_TYPE = (
    "application/vnd.openxmlformats-package.relationships+xml"
)
DOCX_XML_CONTENT_TYPE_ROOTS = {
    DOCX_MAIN_DOCUMENT_CONTENT_TYPE: WORD_DOCUMENT_TAG,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml": WORD_STYLES_TAG,
    "application/vnd.ms-word.styleswitheffects+xml": WORD_STYLES_TAG,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml": WORD_NUMBERING_TAG,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml": WORD_SETTINGS_TAG,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml": WORD_FOOTNOTES_TAG,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml": WORD_ENDNOTES_TAG,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml": WORD_HEADER_TAG,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml": WORD_FOOTER_TAG,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml": WORD_COMMENTS_TAG,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml": (
        WORD_GLOSSARY_DOCUMENT_TAG
    ),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.fonttable+xml": WORD_FONTS_TAG,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.websettings+xml": (
        WORD_WEB_SETTINGS_TAG
    ),
    PACKAGE_RELATIONSHIP_CONTENT_TYPE: PACKAGE_RELATIONSHIPS_TAG,
    "application/vnd.openxmlformats-officedocument.theme+xml": DRAWING_THEME_TAG,
    "application/vnd.openxmlformats-package.core-properties+xml": CORE_PROPERTIES_TAG,
    "application/vnd.openxmlformats-officedocument.extended-properties+xml": (
        EXTENDED_PROPERTIES_TAG
    ),
    "application/vnd.openxmlformats-officedocument.custom-properties+xml": (
        CUSTOM_PROPERTIES_TAG
    ),
}
OFFICE_DOCUMENT_RELATIONSHIP_TYPES = {
    f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/officeDocument",
}
PACKAGE_ROOT_RELATIONSHIP_EXPECTED_CONTENT_TYPES = {
    "coreProperties": {
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties":
            "application/vnd.openxmlformats-package.core-properties+xml",
    },
    "extendedProperties": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/extended-properties":
            "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    },
    "customProperties": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/custom-properties":
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
    },
}
DOCX_RELATIONSHIP_EXPECTED_CONTENT_TYPES = {
    "styles": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/styles":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    },
    "numbering": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/numbering":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    },
    "settings": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/settings":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
    },
    "footnotes": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/footnotes":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
    },
    "endnotes": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/endnotes":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
    },
    "header": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/header":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    },
    "footer": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/footer":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    },
    "comments": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/comments":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    },
    "fontTable": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/fontTable":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
    },
    "webSettings": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/webSettings":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
    },
    "theme": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/theme":
            "application/vnd.openxmlformats-officedocument.theme+xml",
    },
    "glossaryDocument": {
        f"{OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/glossaryDocument":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml",
    },
    "stylesWithEffects": {
        "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects":
            "application/vnd.ms-word.stylesWithEffects+xml",
    },
}
# python-docx only registers specialized Part classes for the document, styles,
# numbering, settings, core-properties, comments, headers and footers parts.  The
# following related parts are loaded as the generic ``Part`` class, so a media type
# whose ASCII letters differ only in case remains safely consumable by the current
# formatter.  Keep these allow-lists narrow; a case alias for a registered part would
# silently downgrade it to ``Part`` and make the corresponding high-level API fail.
CASE_INSENSITIVE_PACKAGE_ROOT_RELATIONSHIP_CATEGORIES = frozenset(
    {"extendedProperties", "customProperties"}
)
CASE_INSENSITIVE_DOCX_RELATIONSHIP_CATEGORIES = frozenset(
    {
        "footnotes",
        "endnotes",
        "fontTable",
        "webSettings",
        "theme",
        "glossaryDocument",
        "stylesWithEffects",
    }
)
KNOWN_CONSUMED_DOCX_RELATIONSHIP_TYPES = frozenset(
    set(OFFICE_DOCUMENT_RELATIONSHIP_TYPES)
    | {
        relationship_type
        for category in (
            PACKAGE_ROOT_RELATIONSHIP_EXPECTED_CONTENT_TYPES,
            DOCX_RELATIONSHIP_EXPECTED_CONTENT_TYPES,
        )
        for relationship_types in category.values()
        for relationship_type in relationship_types
    }
)
KNOWN_CONSUMED_DOCX_RELATIONSHIP_TYPES_ASCII_LOWER = frozenset(
    relationship_type.translate(ASCII_PART_NAME_CASE_TABLE)
    for relationship_type in KNOWN_CONSUMED_DOCX_RELATIONSHIP_TYPES
)
STRICT_RELATIONSHIP_PREFIX_ASCII_LOWER = (
    f"{STRICT_OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE}/".translate(
        ASCII_PART_NAME_CASE_TABLE
    )
)
def contains_forbidden_xml_declaration(xml_bytes: bytes) -> bool:
    """拒绝 DTD/实体声明，避免标准库 XML 解析器展开不受信任实体。"""
    # OOXML 通常使用 UTF-8，但 XML 也允许 UTF-16/32；移除这些编码在
    # ASCII 标记间插入的 NUL 后再匹配，避免仅靠改编码绕过检查。
    lowered = bytes(xml_bytes).replace(b"\x00", b"").lower()
    return any(marker in lowered for marker in FORBIDDEN_XML_DECLARATION_MARKERS)


def stream_contains_forbidden_xml_declaration(xml_stream, chunk_size: int = 64 * 1024) -> bool:
    """分块检查 DTD/实体声明，避免为大型生成 XML 一次性分配整份内存。"""
    overlap_size = max(len(marker) for marker in FORBIDDEN_XML_DECLARATION_MARKERS) - 1
    overlap = b""
    try:
        while True:
            chunk = xml_stream.read(chunk_size)
            if not chunk:
                return False
            normalized = (overlap + chunk).replace(b"\x00", b"").lower()
            if any(marker in normalized for marker in FORBIDDEN_XML_DECLARATION_MARKERS):
                return True
            overlap = normalized[-overlap_size:] if overlap_size else b""
    except (OSError, RuntimeError, ValueError, zlib.error):
        return True


def declared_docx_member_content_type(member_name: str, defaults: dict, overrides: dict) -> str:
    """返回 ZIP 成员声明的原始 media type，保留消费端可见的大小写与空白。"""
    basename = member_name.rsplit("/", 1)[-1].lower() if isinstance(member_name, str) else ""
    extension = basename.rsplit(".", 1)[-1] if "." in basename else ""
    # python-docx 的 ContentTypeMap 使用 CaseInsensitiveDict 查找
    # Override PartName，其键语义正是 str.lower()。与之完全对齐可
    # 避免 Unicode 大小写别名在验证与实际加载时命中不同的
    # Override；仍不解码百分号别名，因为消费端也不解码。
    override_key = member_name.lower() if isinstance(member_name, str) else ""
    content_type = overrides.get(override_key, defaults.get(extension, ""))
    if not isinstance(content_type, str):
        return ""
    return content_type


def docx_content_type_matches(
    actual: str,
    expected: str,
    *,
    allow_ascii_case_alias: bool = False,
) -> bool:
    """比较 OPC media type，同时可选地放行纯 ASCII 大小写别名。

    MIME type/subtype 的字母大小写不改变其语义，但 ``python-docx`` 的
    ``PartFactory`` 按原始字符串精确选择专用 Part 类。因而不能把所有已知
    Content-Type 一概做 ``lower()`` 比较：注册的部件一旦带大小写别名就会
    降级为通用 ``Part``，在访问 ``Document.styles`` 等 API 时才失败。

    仅当调用方明确允许时才比较 ASCII 大小写归一化后的值；不做
    ``strip()``、参数解析或 Unicode case-fold，因此前导/尾随空白、参数和
    非 ASCII 变体仍会被拒绝。
    """
    if not isinstance(actual, str) or not isinstance(expected, str):
        return False
    if actual == expected:
        return True
    if not allow_ascii_case_alias or not actual.isascii() or not expected.isascii():
        return False
    return actual.translate(ASCII_PART_NAME_CASE_TABLE) == expected.translate(
        ASCII_PART_NAME_CASE_TABLE
    )


def docx_member_content_type(member_name: str, defaults: dict, overrides: dict) -> str:
    """返回用于 XML 分类的 media type 主段，忽略可选 MIME 参数。"""
    content_type = declared_docx_member_content_type(member_name, defaults, overrides)
    # MIME 类型可带参数；分类 XML 时只比较 media type 主段，防止
    # 用 ``application/xml; charset=...`` 把受信任 XML 部件伪装成二进制文件。
    return content_type.split(";", 1)[0].strip().lower()


def canonical_docx_part_name(member_name: str) -> str:
    """按 OPC Part URI 及 python-docx 键语义归一化大小写与百分号编码。"""
    decoded_name = unquote(member_name) if isinstance(member_name, str) else ""
    return decoded_name.lower()


def expected_docx_xml_root(
    member_name: str,
    defaults: dict | None = None,
    overrides: dict | None = None,
) -> str | None:
    """返回核心 OOXML 部件必须使用的根标签。"""
    exact_roots = {
        "word/document.xml": WORD_DOCUMENT_TAG,
        "word/styles.xml": WORD_STYLES_TAG,
        "word/stylesWithEffects.xml": WORD_STYLES_TAG,
        "word/numbering.xml": WORD_NUMBERING_TAG,
        "word/settings.xml": WORD_SETTINGS_TAG,
        "word/footnotes.xml": WORD_FOOTNOTES_TAG,
        "word/endnotes.xml": WORD_ENDNOTES_TAG,
        "word/comments.xml": WORD_COMMENTS_TAG,
        "word/fontTable.xml": WORD_FONTS_TAG,
        "word/webSettings.xml": WORD_WEB_SETTINGS_TAG,
        "word/glossary/document.xml": WORD_GLOSSARY_DOCUMENT_TAG,
        "docProps/core.xml": CORE_PROPERTIES_TAG,
        "docProps/app.xml": EXTENDED_PROPERTIES_TAG,
        "docProps/custom.xml": CUSTOM_PROPERTIES_TAG,
    }
    expected = exact_roots.get(member_name)
    if expected is not None:
        return expected
    if re.fullmatch(r"word/header\d+\.xml", member_name):
        return WORD_HEADER_TAG
    if re.fullmatch(r"word/footer\d+\.xml", member_name):
        return WORD_FOOTER_TAG
    if re.fullmatch(r"word/theme/theme\d+\.xml", member_name):
        return DRAWING_THEME_TAG
    if member_name.lower().endswith(".rels"):
        return PACKAGE_RELATIONSHIPS_TAG
    if defaults is not None and overrides is not None:
        return DOCX_XML_CONTENT_TYPE_ROOTS.get(
            docx_member_content_type(member_name, defaults, overrides)
        )
    return None


def is_bounded_xml_structure(
    xml_stream,
    element_budget=None,
    paragraph_budget=None,
    run_budget=None,
    *,
    expected_root: str | None = None,
    max_elements: int | None = None,
    max_depth: int | None = None,
    max_total_elements: int | None = None,
    max_total_paragraphs: int | None = None,
    max_total_runs: int | None = None,
) -> bool:
    """流式限制通用 OOXML 部件的元素、段落与嵌套深度。"""
    max_elements = MAX_DOCX_XML_ELEMENTS if max_elements is None else max_elements
    max_depth = MAX_DOCX_XML_DEPTH if max_depth is None else max_depth
    max_total_elements = MAX_DOCX_TOTAL_XML_ELEMENTS if max_total_elements is None else max_total_elements
    max_total_paragraphs = (
        MAX_DOCX_TOTAL_PARAGRAPHS
        if max_total_paragraphs is None
        else max_total_paragraphs
    )
    max_total_runs = (
        MAX_DOCX_TOTAL_RUNS if max_total_runs is None else max_total_runs
    )
    if paragraph_budget is None:
        paragraph_budget = [0]
    if run_budget is None:
        run_budget = [0]
    depth = 0
    element_count = 0
    root_seen = False
    try:
        for event, element in ElementTree.iterparse(xml_stream, events=("start", "end")):
            if event == "start":
                depth += 1
                element_count += 1
                if not root_seen:
                    root_seen = True
                    if expected_root is not None and element.tag != expected_root:
                        return False
                if element_budget is not None:
                    element_budget[0] += 1
                if element.tag == WORD_PARAGRAPH_TAG:
                    paragraph_budget[0] += 1
                elif element.tag == WORD_RUN_TAG:
                    run_budget[0] += 1
                if (
                    depth > max_depth
                    or element_count > max_elements
                    or paragraph_budget[0] > max_total_paragraphs
                    or run_budget[0] > max_total_runs
                    or (
                        element_budget is not None
                        and element_budget[0] > max_total_elements
                    )
                ):
                    return False
                continue

            depth -= 1
            if depth < 0:
                return False
            element.clear()
    except (OSError, RuntimeError, LookupError, ElementTree.ParseError, zlib.error):
        return False

    return root_seen and depth == 0 and element_count > 0


def is_safe_docx_member_name(member: zipfile.ZipInfo) -> bool:
    """限制 OOXML ZIP 成员为相对、分段明确的普通文件/目录。"""
    name = member.filename
    original_name = getattr(member, "orig_filename", name)
    if (
        not isinstance(name, str)
        or not isinstance(original_name, str)
        or not name
        or original_name != name
        or "\\" in name
        or "\x00" in name
        or any(not character.isprintable() for character in name)
        or name.startswith("/")
        or re.match(r"^[A-Za-z]:", name)
    ):
        return False

    decoded_name = unquote(name)
    if decoded_name != name and (
        "\\" in decoded_name
        or "\x00" in decoded_name
        or any(not character.isprintable() for character in decoded_name)
        or decoded_name.count("/") != name.count("/")
        or decoded_name.startswith("/")
        or re.match(r"^[A-Za-z]:", decoded_name)
    ):
        return False
    if decoded_name != name:
        decoded_is_directory = decoded_name.endswith("/")
        decoded_segments = (
            decoded_name[:-1].split("/")
            if decoded_is_directory
            else decoded_name.split("/")
        )
        if any(segment in {"", ".", ".."} for segment in decoded_segments):
            return False

    is_directory = name.endswith("/")
    segments = name[:-1].split("/") if is_directory else name.split("/")
    if not segments or any(segment in {"", ".", ".."} for segment in segments):
        return False

    unix_mode = (member.external_attr >> 16) & 0xFFFF
    file_type = stat.S_IFMT(unix_mode)
    if file_type not in {0, stat.S_IFREG, stat.S_IFDIR}:
        return False
    if is_directory and file_type == stat.S_IFREG:
        return False
    if not is_directory and file_type == stat.S_IFDIR:
        return False
    return True


def parse_docx_content_types(
    archive: zipfile.ZipFile,
    *,
    max_bytes: int = MAX_DOCX_CONTENT_TYPES_BYTES,
    max_items: int = MAX_DOCX_ARCHIVE_MEMBERS,
):
    """读取 OOXML 内容类型表，用于识别扩展名被修改的 XML 部件。"""
    try:
        content_types_info = archive.getinfo("[Content_Types].xml")
        if content_types_info.file_size > max_bytes:
            return None

        content_types_bytes = archive.read(content_types_info)
        if contains_forbidden_xml_declaration(content_types_bytes):
            return None
        root = ElementTree.fromstring(content_types_bytes)
    except (
        KeyError,
        OSError,
        RuntimeError,
        NotImplementedError,
        zipfile.BadZipFile,
        zlib.error,
        LookupError,
        ElementTree.ParseError,
    ):
        return None

    namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
    if root.tag != f"{{{namespace}}}Types":
        return None
    if isinstance(root.text, str) and root.text.strip():
        return None

    defaults = {}
    overrides = {}
    default_tag = f"{{{namespace}}}Default"
    override_tag = f"{{{namespace}}}Override"
    for child in root:
        if (
            len(child)
            or (isinstance(child.text, str) and child.text.strip())
            or (isinstance(child.tail, str) and child.tail.strip())
        ):
            return None
        if child.tag == default_tag:
            raw_extension = child.attrib.get("Extension", "")
            extension = raw_extension.strip().lower()
            content_type = child.attrib.get("ContentType", "")
            if (
                not extension
                or raw_extension != raw_extension.strip()
                or len(extension) > 64
                or any(
                    character in "/\\."
                    or character.isspace()
                    or not character.isprintable()
                    for character in extension
                )
                or not content_type.strip()
                or extension in defaults
            ):
                return None
            defaults[extension] = content_type
        elif child.tag == override_tag:
            raw_part_name = child.attrib.get("PartName", "")
            part_name = raw_part_name.strip()
            content_type = child.attrib.get("ContentType", "")
            if (
                raw_part_name != part_name
                or not part_name.startswith("/")
                or not content_type.strip()
            ):
                return None
            normalized_part_name = part_name[1:]
            # python-docx 的 CaseInsensitiveDict 以 str.lower() 存取 PartName。
            # 在入口用相同规则拒绝重复键，避免验证器与
            # 消费端对“最终生效 Override”产生分歧。
            override_key = normalized_part_name.lower()
            if not normalized_part_name or override_key in overrides:
                return None
            overrides[override_key] = content_type
        else:
            return None

        if len(defaults) + len(overrides) > max_items:
            return None

    return defaults, overrides


def _relationship_source_part_name(relationship_part_name: str) -> str | None:
    """返回关系部件对应的源部件名；包根关系的源部件为空字符串。"""
    if relationship_part_name == "_rels/.rels":
        return ""
    if not isinstance(relationship_part_name, str) or not relationship_part_name.endswith(".rels"):
        return None
    if relationship_part_name.startswith("_rels/"):
        basename = relationship_part_name[len("_rels/") :]
        if not basename or "/" in basename or not basename.endswith(".rels"):
            return None
        source_basename = basename[:-5]
        return source_basename or None
    if "/_rels/" not in relationship_part_name:
        return None
    directory, basename = relationship_part_name.rsplit("/_rels/", 1)
    if not directory or not basename or not basename.endswith(".rels"):
        return None
    source_basename = basename[:-5]
    if not source_basename:
        return None
    return f"{directory}/{source_basename}"


def _canonical_relationship_type_alias(value: str) -> str:
    """规范化关系 Type，仅用于识别已知 URI 的消费端不兼容别名。"""
    if not isinstance(value, str):
        return ""
    try:
        parsed = urlsplit(unquote(value))
    except (UnicodeError, ValueError):
        return ""
    if not parsed.scheme or not parsed.netloc:
        return ""

    normalized_path = posixpath.normpath(parsed.path)
    canonical = f"{parsed.scheme}://{parsed.netloc}{normalized_path}"
    return canonical.translate(ASCII_PART_NAME_CASE_TABLE)


def _resolve_internal_relationship_target(source_part_name: str, target: str) -> str | None:
    """解析 OPC 内部关系目标，并拒绝越过包根的路径。"""
    if not isinstance(target, str) or not target or target != target.strip():
        return None
    if any(not character.isprintable() or character == "\\" for character in target):
        return None

    try:
        parsed = urlsplit(target)
    except ValueError:
        return None
    # 内部目标只能是相对/绝对 part URI；查询串、片段和 scheme 会被
    # python-docx 当作路径处理，容易造成关系语义与实际 ZIP 成员不一致。
    if parsed.scheme or parsed.netloc or parsed.query or parsed.fragment or not parsed.path:
        return None

    raw_path = parsed.path
    decoded_path = unquote(raw_path)
    if any(
        not character.isprintable() or character == "\\"
        for character in decoded_path
    ):
        return None
    # 编码后的路径分隔符/点段仍不能被用来伪造包外路径。
    decoded_segments = decoded_path.lstrip("/").split("/")
    if any(segment == "" for segment in decoded_segments):
        return None

    if raw_path.startswith("/"):
        joined = raw_path.lstrip("/")
    else:
        base_directory = posixpath.dirname(source_part_name)
        joined = posixpath.join(base_directory, raw_path)
    normalized = posixpath.normpath(joined)
    if (
        not normalized
        or normalized in {".", ".."}
        or normalized.startswith("../")
        or normalized.startswith("/")
    ):
        return None
    return normalized


def _relationship_target_is_external(target_mode: str | None) -> bool | None:
    """规范化 TargetMode；None 表示 OPC 默认的 Internal。"""
    if target_mode is None or target_mode == "Internal":
        return False
    if target_mode == "External":
        return True
    return None


def _docx_relationship_graph_within_depth(
    relationship_records: dict[str, list[dict]],
    max_depth: int,
) -> bool:
    """迭代模拟 python-docx 的关系 DFS，避免验证器自身递归。"""
    if isinstance(max_depth, bool) or not isinstance(max_depth, int) or max_depth < 1:
        return False

    root_records = relationship_records.get("")
    if root_records is None:
        return False

    visited_parts: set[str] = set()
    stack = [(iter(root_records), 0)]
    while stack:
        records, source_depth = stack[-1]
        try:
            record = next(records)
        except StopIteration:
            stack.pop()
            continue

        target_part = record["target_part"]
        if record["external"] or not target_part or target_part in visited_parts:
            continue

        target_depth = source_depth + 1
        if target_depth > max_depth:
            return False
        visited_parts.add(target_part)
        stack.append((iter(relationship_records.get(target_part, ())), target_depth))

    return True


def validate_docx_relationships(
    archive: zipfile.ZipFile,
    member_names: set[str],
    defaults: dict,
    overrides: dict,
    *,
    max_part_bytes: int = MAX_DOCX_RELATIONSHIP_PART_BYTES,
    max_total_bytes: int = MAX_DOCX_TOTAL_RELATIONSHIP_BYTES,
    max_relationships_per_part: int = MAX_DOCX_RELATIONSHIPS_PER_PART,
    max_total_relationships: int = MAX_DOCX_TOTAL_RELATIONSHIPS,
    max_graph_depth: int = MAX_DOCX_RELATIONSHIP_GRAPH_DEPTH,
) -> str | None:
    """校验 OPC 关系图，成功时返回主文档 Part 名。"""
    # 没有包根关系时 python-docx 无法找到 officeDocument part，后续
    # Document() 会抛 KeyError；这类 ZIP 不是可处理的 DOCX。
    if "_rels/.rels" not in member_names:
        return None

    relationship_tag = f"{{{PACKAGE_RELATIONSHIPS_NAMESPACE}}}Relationship"
    relationship_records: dict[str, list[dict]] = {}
    relationship_part_names = {
        name
        for name in member_names
        if isinstance(name, str) and name.lower().endswith(".rels")
    }
    # Build the alias index once.  Without this, each orphan relationship
    # part would scan and normalize every package member, making validation
    # quadratic for packages containing many harmless leftovers.
    canonical_member_names = {
        canonical_docx_part_name(member_name) for member_name in member_names
    }
    total_relationship_bytes = 0
    total_relationships = 0
    for relationship_part_name in relationship_part_names:
        source_part_name = _relationship_source_part_name(relationship_part_name)
        if source_part_name is None:
            return None
        orphan_relationship_part = False
        if source_part_name:
            # Relationship parts are only loaded by python-docx when reached
            # from the package relationship graph.  Some producers (notably
            # documents edited repeatedly in Word/LibreOffice) leave an
            # orphan ``*.rels`` part behind after its source part is removed;
            # python-docx ignores such an unreachable part.  Preserve that
            # compatibility by ignoring its relationship records, while still
            # applying the same size, XML, and relationship-shape budgets.
            if source_part_name not in member_names:
                if source_part_name + "/" in member_names:
                    # A relationship part whose source resolves to a package
                    # directory is malformed, not an orphaned part.
                    return None
                # A missing source is only an orphan when no package member
                # resolves to the same OPC name after decoding/case folding.
                # Case/percent aliases (for example ``Document.xml`` or
                # ``%64ocument.xml``) are otherwise silently ignored by
                # python-docx and must remain rejected by this validator.
                source_canonical_name = canonical_docx_part_name(source_part_name)
                if source_canonical_name in canonical_member_names:
                    return None
                orphan_relationship_part = True
            elif archive.getinfo(source_part_name).is_dir():
                return None

        info = archive.getinfo(relationship_part_name)
        if (
            info.is_dir()
            or info.file_size <= 0
            or info.file_size > max_part_bytes
        ):
            return None
        total_relationship_bytes += info.file_size
        if total_relationship_bytes > max_total_bytes:
            return None
        if not docx_content_type_matches(
            declared_docx_member_content_type(
                relationship_part_name, defaults, overrides
            ),
            PACKAGE_RELATIONSHIP_CONTENT_TYPE,
            # Relationship parts are parsed by their ``.rels`` name and are not
            # dispatched through python-docx's specialized PartFactory classes.
            # A pure ASCII case alias is therefore safe, while parameters/space
            # variants remain rejected by the exact-match helper above.
            allow_ascii_case_alias=True,
        ):
            return None
        try:
            relationship_xml = archive.read(info)
            if contains_forbidden_xml_declaration(relationship_xml):
                return None
            root = ElementTree.fromstring(relationship_xml)
        except (
            KeyError,
            OSError,
            RuntimeError,
            NotImplementedError,
            zipfile.BadZipFile,
            zlib.error,
            LookupError,
            ElementTree.ParseError,
        ):
            return None
        if (
            root.tag != PACKAGE_RELATIONSHIPS_TAG
            or (isinstance(root.text, str) and root.text.strip())
        ):
            return None

        relationship_count = len(root)
        if relationship_count > max_relationships_per_part:
            return None
        total_relationships += relationship_count
        if total_relationships > max_total_relationships:
            return None

        ids = set()
        records = []
        for child in root:
            if (
                child.tag != relationship_tag
                or len(child)
                or (isinstance(child.text, str) and child.text.strip())
                or (isinstance(child.tail, str) and child.tail.strip())
            ):
                return None
            relationship_id = child.attrib.get("Id")
            relationship_type = child.attrib.get("Type")
            target = child.attrib.get("Target")
            target_mode = child.attrib.get("TargetMode")
            canonical_relationship_type_alias = (
                _canonical_relationship_type_alias(relationship_type)
            )
            if (
                not isinstance(relationship_id, str)
                or not relationship_id
                or relationship_id != relationship_id.strip()
                or any(character.isspace() for character in relationship_id)
                or relationship_id in ids
                or not isinstance(relationship_type, str)
                or not relationship_type
                or relationship_type != relationship_type.strip()
                or any(not character.isprintable() for character in relationship_type)
                # python-docx 1.2.0 的 RT 常量及 Part 查找只支持
                # Transitional 关系命名空间。接受 Strict 关系会
                # 让主文档直接打开失败，或让样式/编号悄然降级。
                or canonical_relationship_type_alias.startswith(
                    STRICT_RELATIONSHIP_PREFIX_ASCII_LOWER
                )
                # URI 大小写别名会让 python-docx 把已知关系当成
                # 无关的自定义关系，进而静默丢失原样式/编号。百分号、
                # 查询串、片段和点段别名具有相同风险，也统一拒绝。
                or (
                    canonical_relationship_type_alias
                    in KNOWN_CONSUMED_DOCX_RELATIONSHIP_TYPES_ASCII_LOWER
                    and relationship_type not in KNOWN_CONSUMED_DOCX_RELATIONSHIP_TYPES
                )
                or not isinstance(target, str)
                or not target
                or target != target.strip()
            ):
                return None
            target_mode_external = _relationship_target_is_external(target_mode)
            if target_mode_external is None:
                return None
            try:
                parsed_type = urlsplit(relationship_type)
            except ValueError:
                return None
            if not parsed_type.scheme:
                return None

            target_part_name = None
            if not target_mode_external:
                target_part_name = _resolve_internal_relationship_target(
                    source_part_name,
                    target,
                )
                if target_part_name is None:
                    return None
                # ``python-docx`` resolves relationship targets with POSIX path
                # normalization but then looks up the resulting ZIP member name
                # exactly.  Treating percent-decoded or ASCII-case-equivalent
                # names as aliases here would therefore accept a package that
                # passes validation and immediately fails while being opened.
                if (
                    not orphan_relationship_part
                    and (
                        target_part_name not in member_names
                        or archive.getinfo(target_part_name).is_dir()
                    )
                ):
                    return None
            else:
                # 外部目标可以是 http(s)、mailto 等 URI，但不能包含控制字符。
                if any(not character.isprintable() for character in target):
                    return None

            ids.add(relationship_id)
            records.append(
                {
                    "id": relationship_id,
                    "type": relationship_type,
                    "target": target,
                    "external": target_mode_external,
                    "target_part": target_part_name,
                }
            )
        if not orphan_relationship_part:
            relationship_records[source_part_name] = records

    root_records = relationship_records.get("")
    if root_records is None:
        return None
    if not _docx_relationship_graph_within_depth(
        relationship_records,
        max_graph_depth,
    ):
        return None
    office_document_records = [
        record
        for record in root_records
        if record["type"] in OFFICE_DOCUMENT_RELATIONSHIP_TYPES
    ]
    if len(office_document_records) != 1 or office_document_records[0]["external"]:
        return None

    main_document_part_name = office_document_records[0]["target_part"]
    if (
        not main_document_part_name
        or declared_docx_member_content_type(
            main_document_part_name,
            defaults,
            overrides,
        )
        != DOCX_MAIN_DOCUMENT_CONTENT_TYPE
    ):
        return None

    for category, relationship_types in (
        PACKAGE_ROOT_RELATIONSHIP_EXPECTED_CONTENT_TYPES.items()
    ):
        matching_records = [
            record for record in root_records if record["type"] in relationship_types
        ]
        if len(matching_records) > 1:
            return None
        expected_content_types = set(relationship_types.values())
        for record in matching_records:
            if record["external"] or not record["target_part"]:
                return None
            actual_content_type = declared_docx_member_content_type(
                record["target_part"],
                defaults,
                overrides,
            )
            if not any(
                docx_content_type_matches(
                    actual_content_type,
                    expected_content_type,
                    allow_ascii_case_alias=(
                        category
                        in CASE_INSENSITIVE_PACKAGE_ROOT_RELATIONSHIP_CATEGORIES
                    ),
                )
                for expected_content_type in expected_content_types
            ):
                return None

    document_records = relationship_records.get(main_document_part_name, ())
    for category, relationship_types in DOCX_RELATIONSHIP_EXPECTED_CONTENT_TYPES.items():
        matching_records = [
            record for record in document_records if record["type"] in relationship_types
        ]
        # styles/numbering/settings/footnotes/endnotes are singleton relations;
        # multiple headers/footers are legal because each has its own part.
        if category not in {"header", "footer"} and len(matching_records) > 1:
            return None
        expected_content_types = set(relationship_types.values())
        for record in matching_records:
            if record["external"] or not record["target_part"]:
                return None
            actual_content_type = declared_docx_member_content_type(
                record["target_part"], defaults, overrides
            )
            if not any(
                docx_content_type_matches(
                    actual_content_type,
                    expected_content_type,
                    allow_ascii_case_alias=(
                        category in CASE_INSENSITIVE_DOCX_RELATIONSHIP_CATEGORIES
                    ),
                )
                for expected_content_type in expected_content_types
            ):
                return None

    return main_document_part_name


def is_docx_xml_member(member_name: str, defaults: dict, overrides: dict) -> bool:
    normalized_name = member_name.lower()
    if normalized_name.endswith((".xml", ".rels")):
        return True

    content_type = docx_member_content_type(member_name, defaults, overrides)
    return content_type in {"application/xml", "text/xml"} or content_type.endswith("+xml")


def parse_bounded_grid_span(value) -> int | None:
    """解析 OOXML 表格跨度，避免把超大十进制值交给 int() 或 python-docx。"""
    if not isinstance(value, str):
        return None

    digits = value.strip()
    if digits.startswith("+"):
        digits = digits[1:]
    if not digits or not digits.isascii() or not digits.isdecimal():
        return None

    significant_digits = digits.lstrip("0") or "0"
    maximum_text = str(MAX_DOCX_TABLE_GRID_COLUMNS)
    if len(significant_digits) > len(maximum_text):
        return None

    span = int(significant_digits)
    if span < 1 or span > MAX_DOCX_TABLE_GRID_COLUMNS:
        return None
    return span


def is_bounded_docx_document_xml(
    xml_stream,
    element_budget=None,
    paragraph_budget=None,
    run_budget=None,
    *,
    max_elements: int | None = None,
    max_depth: int | None = None,
    max_total_elements: int | None = None,
    max_total_paragraphs: int | None = None,
    max_total_runs: int | None = None,
    max_grid_columns: int | None = None,
    max_logical_cells: int | None = None,
) -> bool:
    """流式校验主文档结构、段落与表格规模，阻断对象放大。"""
    max_elements = MAX_DOCX_XML_ELEMENTS if max_elements is None else max_elements
    max_depth = MAX_DOCX_XML_DEPTH if max_depth is None else max_depth
    max_total_elements = MAX_DOCX_TOTAL_XML_ELEMENTS if max_total_elements is None else max_total_elements
    max_total_paragraphs = (
        MAX_DOCX_TOTAL_PARAGRAPHS
        if max_total_paragraphs is None
        else max_total_paragraphs
    )
    max_total_runs = (
        MAX_DOCX_TOTAL_RUNS if max_total_runs is None else max_total_runs
    )
    max_grid_columns = MAX_DOCX_TABLE_GRID_COLUMNS if max_grid_columns is None else max_grid_columns
    max_logical_cells = MAX_DOCX_TABLE_LOGICAL_CELLS if max_logical_cells is None else max_logical_cells
    if paragraph_budget is None:
        paragraph_budget = [0]
    if run_budget is None:
        run_budget = [0]
    tag_stack = []
    table_states = []
    row_states = []
    cell_states = []
    total_logical_cells = 0
    element_count = 0
    root_seen = False
    background_seen = False
    body_seen = False

    try:
        for event, element in ElementTree.iterparse(xml_stream, events=("start", "end")):
            tag = element.tag
            if event == "start":
                element_count += 1
                if element_budget is not None:
                    element_budget[0] += 1
                if tag == WORD_PARAGRAPH_TAG:
                    paragraph_budget[0] += 1
                elif tag == WORD_RUN_TAG:
                    run_budget[0] += 1
                if (
                    element_count > max_elements
                    or len(tag_stack) >= max_depth
                    or paragraph_budget[0] > max_total_paragraphs
                    or run_budget[0] > max_total_runs
                    or (
                        element_budget is not None
                        and element_budget[0] > max_total_elements
                    )
                ):
                    return False
                if not root_seen:
                    root_seen = True
                    if tag != WORD_DOCUMENT_TAG:
                        return False
                elif tag == WORD_BODY_TAG:
                    # CT_Document 要求 body 是唯一且最后的直接子元素；
                    # 同时拒绝嵌套 body，避免格式化器与校验器视图不一致。
                    if tag_stack != [WORD_DOCUMENT_TAG] or body_seen:
                        return False
                    body_seen = True
                elif tag_stack == [WORD_DOCUMENT_TAG]:
                    # body 之前仅允许规范中可选的单个 background。
                    # body_seen 之后到达这里的任何元素均为越位直接子元素。
                    if tag != WORD_BACKGROUND_TAG or background_seen or body_seen:
                        return False
                    background_seen = True
                tag_stack.append(tag)

                if tag == WORD_TABLE_TAG and (
                    tag_stack[-3:] == [WORD_DOCUMENT_TAG, WORD_BODY_TAG, WORD_TABLE_TAG]
                    or (
                        tag_stack[-2:] == [WORD_TABLE_CELL_TAG, WORD_TABLE_TAG]
                        and cell_states
                        and cell_states[-1]["element_depth"] == len(tag_stack) - 1
                    )
                ):
                    table_states.append({"element_depth": len(tag_stack)})
                elif (
                    tag == WORD_TABLE_ROW_TAG
                    and tag_stack[-2:] == [WORD_TABLE_TAG, WORD_TABLE_ROW_TAG]
                    and table_states
                    and table_states[-1]["element_depth"] == len(tag_stack) - 1
                ):
                    row_states.append(
                        {
                            "element_depth": len(tag_stack),
                            "width": 0,
                        }
                    )
                elif (
                    tag == WORD_TABLE_CELL_TAG
                    and tag_stack[-2:] == [WORD_TABLE_ROW_TAG, WORD_TABLE_CELL_TAG]
                    and row_states
                    and row_states[-1]["element_depth"] == len(tag_stack) - 1
                ):
                    cell_states.append(
                        {
                            "element_depth": len(tag_stack),
                            "span": 1,
                            "span_seen": False,
                        }
                    )
                continue

            if (
                tag == WORD_GRID_SPAN_TAG
                and tag_stack[-3:]
                == [WORD_TABLE_CELL_TAG, WORD_TABLE_CELL_PROPERTIES_TAG, WORD_GRID_SPAN_TAG]
                and cell_states
                and cell_states[-1]["element_depth"] == len(tag_stack) - 2
            ):
                cell_state = cell_states[-1]
                if cell_state["span_seen"]:
                    return False

                span = parse_bounded_grid_span(element.attrib.get(WORD_VALUE_ATTRIBUTE))
                if span is None:
                    return False
                cell_state["span"] = span
                cell_state["span_seen"] = True

            elif (
                tag == WORD_TABLE_CELL_TAG
                and cell_states
                and cell_states[-1]["element_depth"] == len(tag_stack)
            ):
                cell_state = cell_states.pop()
                if not row_states or row_states[-1]["element_depth"] != len(tag_stack) - 1:
                    return False

                span = cell_state["span"]
                row_states[-1]["width"] += span
                total_logical_cells += span
                if (
                    row_states[-1]["width"] > max_grid_columns
                    or total_logical_cells > max_logical_cells
                ):
                    return False

            elif (
                tag == WORD_TABLE_ROW_TAG
                and row_states
                and row_states[-1]["element_depth"] == len(tag_stack)
            ):
                row_states.pop()

            elif (
                tag == WORD_TABLE_TAG
                and table_states
                and table_states[-1]["element_depth"] == len(tag_stack)
            ):
                table_states.pop()

            element.clear()
            if not tag_stack or tag_stack.pop() != tag:
                return False
    except (OSError, RuntimeError, LookupError, ElementTree.ParseError, zlib.error):
        return False

    return (
        root_seen
        and body_seen
        and not tag_stack
        and not table_states
        and not row_states
        and not cell_states
    )


def _stream_size_within_limit(stream, max_archive_bytes: int | None) -> bool:
    """Check raw archive size without reading it into memory."""
    if max_archive_bytes is None:
        return True
    if max_archive_bytes < 0:
        return False

    try:
        stream.seek(0, 2)
        archive_size = stream.tell()
        stream.seek(0)
    except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
        return False

    return (
        isinstance(archive_size, int)
        and not isinstance(archive_size, bool)
        and 0 <= archive_size <= max_archive_bytes
    )


def _has_supported_zip_flags(member: zipfile.ZipInfo) -> bool:
    """Allow only flags consumed consistently by this validator and ZipFile."""
    allowed_flags = _ZIP_DATA_DESCRIPTOR_FLAG | _ZIP_UTF8_FILENAME_FLAG
    if member.compress_type == zipfile.ZIP_DEFLATED:
        allowed_flags |= _ZIP_DEFLATE_OPTION_FLAGS
    return member.flag_bits & ~allowed_flags == 0


def _resolve_local_zip64_sizes(
    local_file_size: int,
    local_compressed_size: int,
    local_extra: bytes,
) -> tuple[int, int] | None:
    """Resolve ZIP64 size sentinels from one local header extra field."""
    if (
        local_file_size != _ZIP_UINT32_MAX
        and local_compressed_size != _ZIP_UINT32_MAX
    ):
        return local_file_size, local_compressed_size

    zip64_payload = None
    offset = 0
    while offset < len(local_extra):
        if len(local_extra) - offset < 4:
            return None
        field_id = int.from_bytes(local_extra[offset : offset + 2], "little")
        field_size = int.from_bytes(
            local_extra[offset + 2 : offset + 4],
            "little",
        )
        offset += 4
        field_end = offset + field_size
        if field_end > len(local_extra):
            return None
        if field_id == _ZIP64_EXTRA_FIELD_ID:
            if zip64_payload is not None:
                return None
            zip64_payload = local_extra[offset:field_end]
        offset = field_end

    if zip64_payload is None:
        return None

    payload_offset = 0

    def take_uint64() -> int | None:
        nonlocal payload_offset
        value_end = payload_offset + 8
        if value_end > len(zip64_payload):
            return None
        value = int.from_bytes(
            zip64_payload[payload_offset:value_end],
            "little",
        )
        payload_offset = value_end
        return value

    resolved_file_size = local_file_size
    if resolved_file_size == _ZIP_UINT32_MAX:
        resolved_file_size = take_uint64()
        if resolved_file_size is None:
            return None

    resolved_compressed_size = local_compressed_size
    if resolved_compressed_size == _ZIP_UINT32_MAX:
        resolved_compressed_size = take_uint64()
        if resolved_compressed_size is None:
            return None

    return resolved_file_size, resolved_compressed_size


def _matches_zip_data_descriptor(
    descriptor: bytes,
    member: zipfile.ZipInfo,
) -> bool:
    """Match a bounded standard or ZIP64 data descriptor to central metadata."""
    if len(descriptor) in {16, 24}:
        signature = descriptor[:4]
        descriptor = descriptor[4:]
        if signature != _ZIP_DATA_DESCRIPTOR_SIGNATURE:
            return False
    elif len(descriptor) not in {12, 20}:
        return False

    if len(descriptor) == 12:
        crc = int.from_bytes(descriptor[:4], "little")
        compressed_size = int.from_bytes(descriptor[4:8], "little")
        file_size = int.from_bytes(descriptor[8:12], "little")
    elif len(descriptor) == 20:
        crc = int.from_bytes(descriptor[:4], "little")
        compressed_size = int.from_bytes(descriptor[4:12], "little")
        file_size = int.from_bytes(descriptor[12:20], "little")
    else:
        return False

    return (
        crc == member.CRC
        and compressed_size == member.compress_size
        and file_size == member.file_size
    )


def _build_zip_member_layouts(
    archive: zipfile.ZipFile,
    members: list[zipfile.ZipInfo],
) -> dict[str, tuple[int, int]] | None:
    """Validate local headers and return each member's compressed-data span.

    Central-directory sizes alone cannot distinguish a complete stored member
    from an attacker-selected prefix. Requiring every local entry (including a
    possible data descriptor) to end at the next local header closes that gap.
    """
    archive_stream = archive.fp
    central_directory_offset = archive.start_dir
    if archive_stream is None or not isinstance(central_directory_offset, int):
        return None

    ordered_members = sorted(members, key=lambda member: member.header_offset)
    header_offsets = [member.header_offset for member in ordered_members]
    if (
        any(
            not isinstance(offset, int)
            or isinstance(offset, bool)
            or offset < 0
            or offset >= central_directory_offset
            for offset in header_offsets
        )
        or len(set(header_offsets)) != len(header_offsets)
    ):
        return None

    layouts = {}
    boundaries = header_offsets[1:] + [central_directory_offset]
    for member, boundary in zip(ordered_members, boundaries, strict=True):
        header_offset = member.header_offset
        if boundary - header_offset < 30:
            return None

        archive_stream.seek(header_offset)
        local_header = archive_stream.read(30)
        if (
            not isinstance(local_header, bytes)
            or len(local_header) != 30
            or local_header[:4] != _ZIP_LOCAL_FILE_HEADER_SIGNATURE
        ):
            return None

        local_flags = int.from_bytes(local_header[6:8], "little")
        local_compression = int.from_bytes(local_header[8:10], "little")
        local_crc = int.from_bytes(local_header[14:18], "little")
        local_compressed_size = int.from_bytes(local_header[18:22], "little")
        local_file_size = int.from_bytes(local_header[22:26], "little")
        filename_length = int.from_bytes(local_header[26:28], "little")
        extra_length = int.from_bytes(local_header[28:30], "little")
        local_metadata_size = filename_length + extra_length
        local_metadata = archive_stream.read(local_metadata_size)
        if (
            not isinstance(local_metadata, bytes)
            or len(local_metadata) != local_metadata_size
            or local_flags != member.flag_bits
            or local_compression != member.compress_type
        ):
            return None

        filename_bytes = local_metadata[:filename_length]
        local_extra = local_metadata[filename_length:]
        filename_encoding = "utf-8" if local_flags & 0x800 else "cp437"
        if filename_bytes.decode(filename_encoding) != member.orig_filename:
            return None

        data_offset = header_offset + 30 + local_metadata_size
        compressed_data_end = data_offset + member.compress_size
        if compressed_data_end > boundary:
            return None

        if local_flags & 0x08:
            if (
                local_crc not in {0, member.CRC}
                or local_compressed_size
                not in {0, _ZIP_UINT32_MAX, member.compress_size}
                or local_file_size not in {0, _ZIP_UINT32_MAX, member.file_size}
            ):
                return None
            descriptor_size = boundary - compressed_data_end
            if descriptor_size not in {12, 16, 20, 24}:
                return None
            archive_stream.seek(compressed_data_end)
            descriptor = archive_stream.read(descriptor_size)
            if (
                not isinstance(descriptor, bytes)
                or not _matches_zip_data_descriptor(descriptor, member)
            ):
                return None
        else:
            if compressed_data_end != boundary or local_crc != member.CRC:
                return None
            resolved_sizes = _resolve_local_zip64_sizes(
                local_file_size,
                local_compressed_size,
                local_extra,
            )
            if resolved_sizes != (member.file_size, member.compress_size):
                return None

        layouts[member.filename] = (data_offset, compressed_data_end)

    return layouts


def _verify_zip_member_output(
    archive: zipfile.ZipFile,
    member: zipfile.ZipInfo,
    layout: tuple[int, int],
    *,
    max_output_bytes: int,
) -> int | None:
    """Boundedly verify actual STORE/DEFLATE output length and CRC.

    This intentionally bypasses ``ZipExtFile`` because it truncates output at
    attacker-controlled ``ZipInfo.file_size``. At most one 64 KiB output chunk
    is materialized at a time, plus one byte beyond the active budget so an
    overflow is detected rather than silently accepted.
    """
    if max_output_bytes < 0:
        return None

    archive_stream = archive.fp
    if archive_stream is None:
        return None
    data_offset, compressed_data_end = layout
    remaining_compressed = compressed_data_end - data_offset
    if remaining_compressed != member.compress_size:
        return None

    archive_stream.seek(data_offset)
    actual_size = 0
    actual_crc = 0

    if member.compress_type == zipfile.ZIP_STORED:
        while remaining_compressed:
            compressed_chunk = archive_stream.read(
                min(_ZIP_MEMBER_VALIDATION_CHUNK_BYTES, remaining_compressed)
            )
            if not isinstance(compressed_chunk, bytes) or not compressed_chunk:
                return None
            remaining_compressed -= len(compressed_chunk)
            actual_size += len(compressed_chunk)
            if actual_size > max_output_bytes:
                return None
            actual_crc = zlib.crc32(compressed_chunk, actual_crc)
    elif member.compress_type == zipfile.ZIP_DEFLATED:
        decompressor = zlib.decompressobj(-zlib.MAX_WBITS)
        while remaining_compressed:
            compressed_chunk = archive_stream.read(
                min(_ZIP_MEMBER_VALIDATION_CHUNK_BYTES, remaining_compressed)
            )
            if not isinstance(compressed_chunk, bytes) or not compressed_chunk:
                return None
            remaining_compressed -= len(compressed_chunk)
            pending = compressed_chunk
            while pending:
                previous_pending_size = len(pending)
                output_limit = min(
                    _ZIP_MEMBER_VALIDATION_CHUNK_BYTES,
                    max_output_bytes - actual_size + 1,
                )
                output = decompressor.decompress(pending, output_limit)
                if decompressor.unused_data:
                    return None
                actual_size += len(output)
                if actual_size > max_output_bytes:
                    return None
                actual_crc = zlib.crc32(output, actual_crc)
                pending = decompressor.unconsumed_tail
                if (
                    pending
                    and len(pending) >= previous_pending_size
                    and not output
                ):
                    return None
            if decompressor.eof and remaining_compressed:
                return None

        while not decompressor.eof:
            output_limit = min(
                _ZIP_MEMBER_VALIDATION_CHUNK_BYTES,
                max_output_bytes - actual_size + 1,
            )
            output = decompressor.decompress(b"", output_limit)
            actual_size += len(output)
            if actual_size > max_output_bytes:
                return None
            actual_crc = zlib.crc32(output, actual_crc)
            if not output and not decompressor.eof:
                return None
        if decompressor.unused_data or decompressor.unconsumed_tail:
            return None
    else:
        return None

    if actual_size != member.file_size or actual_crc != member.CRC:
        return None
    return actual_size


def is_valid_docx_stream(
    stream,
    *,
    limits: DocxValidationLimits = INPUT_DOCX_LIMITS,
) -> bool:
    """Validate a seekable DOCX stream and always rewind it to byte zero."""
    try:
        if stream is None or not _valid_docx_validation_limits(limits):
            return False
        if not _stream_size_within_limit(stream, limits.max_archive_bytes):
            return False
        if not zipfile.is_zipfile(stream):
            return False

        stream.seek(0)
        with zipfile.ZipFile(stream) as archive:
            members = archive.infolist()
            if not members or len(members) > limits.max_archive_members:
                return False

            seen_names = set()
            seen_canonical_names = set()
            total_uncompressed = 0
            for member in members:
                member_name = member.filename
                canonical_name = canonical_docx_part_name(member_name)
                if (
                    not is_safe_docx_member_name(member)
                    or len(member_name) > limits.max_member_name_length
                    or member_name in seen_names
                    or canonical_name in seen_canonical_names
                    or member.compress_type
                    not in _ALLOWED_DOCX_COMPRESSION_TYPES
                    or not _has_supported_zip_flags(member)
                ):
                    return False

                seen_names.add(member_name)
                seen_canonical_names.add(canonical_name)
                file_size = member.file_size
                compressed_size = member.compress_size
                if file_size < 0 or compressed_size < 0:
                    return False
                if file_size > limits.max_member_uncompressed_bytes:
                    return False

                total_uncompressed += file_size
                if total_uncompressed > limits.max_total_uncompressed_bytes:
                    return False

                if file_size and compressed_size <= 0:
                    return False

            if not set(REQUIRED_DOCX_MEMBERS).issubset(seen_names):
                return False

            layouts = _build_zip_member_layouts(archive, members)
            if layouts is None:
                return False

            content_types_info = archive.getinfo("[Content_Types].xml")
            if content_types_info.file_size > limits.max_content_types_bytes:
                return False
            content_types_actual_size = _verify_zip_member_output(
                archive,
                content_types_info,
                layouts[content_types_info.filename],
                max_output_bytes=min(
                    content_types_info.file_size,
                    limits.max_member_uncompressed_bytes,
                    limits.max_total_uncompressed_bytes,
                    limits.max_xml_member_uncompressed_bytes,
                    limits.max_total_xml_uncompressed_bytes,
                    limits.max_content_types_bytes,
                ),
            )
            if content_types_actual_size is None:
                return False

            content_types = parse_docx_content_types(
                archive,
                max_bytes=limits.max_content_types_bytes,
                max_items=limits.max_archive_members,
            )
            if content_types is None:
                return False
            defaults, overrides = content_types

            for member in members:
                if member.is_dir() or member.filename == "[Content_Types].xml":
                    continue
                if not docx_member_content_type(member.filename, defaults, overrides):
                    return False
                if (
                    limits.max_non_xml_compression_ratio is not None
                    and member.file_size
                    and not is_docx_xml_member(member.filename, defaults, overrides)
                    and member.file_size / member.compress_size
                    > limits.max_non_xml_compression_ratio
                ):
                    return False

            total_xml_uncompressed = 0
            xml_members = []
            for member in members:
                if member.is_dir() or not is_docx_xml_member(
                    member.filename,
                    defaults,
                    overrides,
                ):
                    continue
                if member.file_size > limits.max_xml_member_uncompressed_bytes:
                    return False
                total_xml_uncompressed += member.file_size
                if total_xml_uncompressed > limits.max_total_xml_uncompressed_bytes:
                    return False
                xml_members.append(member)

            xml_member_names = {member.filename for member in xml_members}
            verified_total_uncompressed = content_types_actual_size
            verified_total_xml_uncompressed = content_types_actual_size
            for member in members:
                if member.filename == "[Content_Types].xml":
                    continue

                max_output_bytes = min(
                    member.file_size,
                    limits.max_member_uncompressed_bytes,
                    limits.max_total_uncompressed_bytes
                    - verified_total_uncompressed,
                )
                if member.filename in xml_member_names:
                    max_output_bytes = min(
                        max_output_bytes,
                        limits.max_xml_member_uncompressed_bytes,
                        limits.max_total_xml_uncompressed_bytes
                        - verified_total_xml_uncompressed,
                    )

                actual_size = _verify_zip_member_output(
                    archive,
                    member,
                    layouts[member.filename],
                    max_output_bytes=max_output_bytes,
                )
                if actual_size is None:
                    return False
                verified_total_uncompressed += actual_size
                if member.filename in xml_member_names:
                    verified_total_xml_uncompressed += actual_size

            if (
                verified_total_uncompressed != total_uncompressed
                or verified_total_xml_uncompressed != total_xml_uncompressed
            ):
                return False

            main_document_part_name = validate_docx_relationships(
                archive,
                seen_names,
                defaults,
                overrides,
                max_part_bytes=limits.max_relationship_part_bytes,
                max_total_bytes=limits.max_total_relationship_bytes,
                max_relationships_per_part=limits.max_relationships_per_part,
                max_total_relationships=limits.max_total_relationships,
                max_graph_depth=limits.max_relationship_graph_depth,
            )
            if not main_document_part_name:
                return False

            element_budget = [0]
            paragraph_budget = [0]
            run_budget = [0]
            for member in xml_members:
                try:
                    with archive.open(member) as xml_stream:
                        xml_bytes = xml_stream.read(member.file_size + 1)
                    if (
                        len(xml_bytes) != member.file_size
                        or contains_forbidden_xml_declaration(xml_bytes)
                    ):
                        return False

                    if member.filename == main_document_part_name:
                        is_valid_xml = is_bounded_docx_document_xml(
                            BytesIO(xml_bytes),
                            element_budget,
                            paragraph_budget,
                            run_budget,
                            max_elements=limits.max_xml_elements,
                            max_depth=limits.max_xml_depth,
                            max_total_elements=limits.max_total_xml_elements,
                            max_total_paragraphs=limits.max_total_paragraphs,
                            max_total_runs=limits.max_total_runs,
                            max_grid_columns=limits.max_table_grid_columns,
                            max_logical_cells=limits.max_table_logical_cells,
                        )
                    elif member.filename == "[Content_Types].xml":
                        continue
                    else:
                        is_valid_xml = is_bounded_xml_structure(
                            BytesIO(xml_bytes),
                            element_budget,
                            paragraph_budget,
                            run_budget,
                            expected_root=expected_docx_xml_root(
                                member.filename,
                                defaults,
                                overrides,
                            ),
                            max_elements=limits.max_xml_elements,
                            max_depth=limits.max_xml_depth,
                            max_total_elements=limits.max_total_xml_elements,
                            max_total_paragraphs=limits.max_total_paragraphs,
                            max_total_runs=limits.max_total_runs,
                        )
                    if not is_valid_xml:
                        return False
                except (
                    OSError,
                    RuntimeError,
                    NotImplementedError,
                    zipfile.BadZipFile,
                    zlib.error,
                    LookupError,
                ):
                    return False

            return True
    except (
        AttributeError,
        EOFError,
        OSError,
        RuntimeError,
        TypeError,
        ValueError,
        LookupError,
        zipfile.BadZipFile,
        zlib.error,
    ):
        return False
    finally:
        with suppress(Exception):
            stream.seek(0)


def is_valid_docx_path(
    path: str | Path,
    *,
    limits: DocxValidationLimits = INPUT_DOCX_LIMITS,
) -> bool:
    """Validate one DOCX path using a single opened file handle."""
    try:
        with Path(path).open("rb") as stream:
            return is_valid_docx_stream(stream, limits=limits)
    except (OSError, RuntimeError, TypeError, ValueError):
        return False


def is_valid_generated_docx(
    output_path: str | Path,
    *,
    limits: DocxValidationLimits = GENERATED_DOCX_LIMITS,
) -> bool:
    """Validate a generated DOCX with the expanded output budget profile."""
    return is_valid_docx_path(output_path, limits=limits)
