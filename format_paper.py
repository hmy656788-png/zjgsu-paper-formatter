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

import colorsys
import hashlib
import math
import re
import sys
import logging
import stat
import tempfile
import threading
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from docx_validation import (
    DocxValidationLimits,
    GENERATED_DOCX_LIMITS,
    INPUT_DOCX_LIMITS,
    MAX_DOCX_TOTAL_PARAGRAPHS,
    is_valid_docx_stream,
    is_valid_generated_docx,
)
from docx import Document
from docx.exceptions import InvalidXmlError
from docx.shared import Pt, RGBColor, Cm
from docx.enum.section import WD_HEADER_FOOTER, WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.opc.exceptions import PackageNotFoundError
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml.exceptions import InvalidXmlError as OxmlInvalidXmlError
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml
from docx.parts.hdrftr import FooterPart, HeaderPart
from docx.parts.numbering import NumberingPart
from docx.table import _Cell
from docx.text.run import Run
from lxml import etree

# ============================================================
# 日志配置
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

FOOTNOTE_FONT_SIZE_PT = 10
FOOTNOTE_XML_PATH = "word/footnotes.xml"
FOOTNOTE_SKIP_TYPES = {"separator", "continuationSeparator", "continuationNotice"}
NOTE_SPECIAL_TYPES = frozenset(
    {"separator", "continuationSeparator", "continuationNotice"}
)
FORBIDDEN_OOXML_DECLARATION_MARKERS = (b"<!doctype", b"<!entity")
MAX_LOG_PATH_LENGTH = 160
# 纯文本与 DOCX 共用同一结构预算：两者都会同时物化 Word 段落、
# 段落列表和分析对象。
MAX_TEXT_PARAGRAPHS = MAX_DOCX_TOTAL_PARAGRAPHS
LOG_PATH_TOKEN_RE = re.compile(r"(?<![\w:/])(?:[A-Za-z]:)?[\\/][^\s\"'<>|,;:]+")
DRAWING_XML_TAGS = frozenset(
    {
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing",
        "{urn:schemas-microsoft-com:vml}shape",
        "{urn:schemas-microsoft-com:office:office}OLEObject",
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict",
    }
)
EQUATION_XML_TAGS = frozenset(
    {
        "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath",
        "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMathPara",
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}object",
        "{urn:schemas-microsoft-com:office:office}OLEObject",
        "{urn:schemas-microsoft-com:vml}shape",
    }
)
# Special hyphen nodes are intentionally excluded: Paragraph.text omits them,
# so rebuilding a labeled paragraph would silently remove their characters.
REWRITE_SAFE_RUN_CHILD_TAGS = frozenset(
    qn(tag_name)
    for tag_name in (
        "w:rPr",
        "w:t",
        "w:tab",
        "w:br",
        "w:cr",
        "w:lastRenderedPageBreak",
    )
)
# These run properties are rewritten by ``_set_run_font`` (or are purely
# typographic defaults), so they do not carry inline semantics that would be
# lost when a heading/caption is rebuilt from plain text.  Other properties
# (for example ``vertAlign``, ``highlight`` or ``smallCaps``) must keep the
# original run boundaries and are therefore treated as rewrite-sensitive.
REWRITE_SAFE_RUN_PROPERTY_TAGS = frozenset(
    qn(tag_name)
    for tag_name in (
        "w:rFonts",
        "w:sz",
        "w:szCs",
        "w:b",
        "w:bCs",
        "w:i",
        "w:iCs",
        "w:u",
        "w:color",
    )
)


class OutputSizeLimitExceeded(Exception):
    """生成的 DOCX 写入超过调用方允许的字节上限。"""

    def __init__(self, max_bytes: int):
        self.max_bytes = max_bytes
        super().__init__(f"generated document exceeds {max_bytes} bytes")


class InvalidInputDocxError(Exception):
    """An external DOCX failed the shared ZIP/OOXML validation policy."""


class InvalidGeneratedDocxError(Exception):
    """A rewritten DOCX failed the generated-output safety policy."""


class TextParagraphLimitExceeded(ValueError):
    """纯文本拆分后的段落数量超过安全预算。"""

    def __init__(self, max_paragraphs: int):
        self.max_paragraphs = max_paragraphs
        super().__init__(f"text exceeds {max_paragraphs} paragraphs")


def _load_validated_document(
    path: str | Path,
    *,
    limits: DocxValidationLimits = INPUT_DOCX_LIMITS,
):
    """Validate and load a DOCX through one file handle to avoid reopen races."""
    with Path(path).open("rb") as stream:
        if not is_valid_docx_stream(stream, limits=limits):
            raise InvalidInputDocxError("document failed DOCX safety validation")
        return Document(stream)


# 核心格式化器的临时文件会与 Web 层的 TTL 清理共用目录。用引用计数
# 而不是单纯 set，避免同一路径被嵌套操作重复登记时提前释放租约。
_ACTIVE_TEMP_PATH_COUNTS: dict[Path, int] = {}
_ACTIVE_TEMP_PATHS_LOCK = threading.Lock()


def _normalize_active_temp_path(path: str | Path) -> Path:
    """返回不解析符号链接的绝对路径，便于与清理器枚举的路径稳定比较。"""
    return Path(path).absolute()


def _register_active_temp_path_locked(path: str | Path) -> Path:
    """在已持有租约锁时登记路径。"""
    normalized = _normalize_active_temp_path(path)
    _ACTIVE_TEMP_PATH_COUNTS[normalized] = _ACTIVE_TEMP_PATH_COUNTS.get(normalized, 0) + 1
    return normalized


def register_active_temp_path(path: str | Path) -> Path:
    """登记正在使用的核心格式化临时路径，并返回标准化路径。"""
    with _ACTIVE_TEMP_PATHS_LOCK:
        return _register_active_temp_path_locked(path)


def release_active_temp_path(path: str | Path) -> None:
    """释放一次临时路径租约；重复释放不会报错。"""
    normalized = _normalize_active_temp_path(path)
    with _ACTIVE_TEMP_PATHS_LOCK:
        remaining = _ACTIVE_TEMP_PATH_COUNTS.get(normalized, 0) - 1
        if remaining > 0:
            _ACTIVE_TEMP_PATH_COUNTS[normalized] = remaining
        else:
            _ACTIVE_TEMP_PATH_COUNTS.pop(normalized, None)


def get_active_temp_paths() -> frozenset[Path]:
    """返回当前活动临时路径的线程安全快照。"""
    with _ACTIVE_TEMP_PATHS_LOCK:
        return frozenset(_ACTIVE_TEMP_PATH_COUNTS)


def is_active_temp_path(path: str | Path) -> bool:
    """在租约锁内复核指定路径是否仍在使用。"""
    normalized = _normalize_active_temp_path(path)
    with _ACTIVE_TEMP_PATHS_LOCK:
        return normalized in _ACTIVE_TEMP_PATH_COUNTS


def remove_expired_inactive_temp_path(path: str | Path, cutoff: float) -> bool:
    """仅在路径过期且无核心格式化租约时删除它。

    检查与删除共用租约锁，避免新租约在两步之间插入。返回值
    仅表示本次是否实际删除了路径。
    """
    normalized = _normalize_active_temp_path(path)
    with _ACTIVE_TEMP_PATHS_LOCK:
        if normalized in _ACTIVE_TEMP_PATH_COUNTS:
            return False

        # The path may disappear between directory enumeration and this
        # locked recheck (for example, a worker can finish and unlink its
        # staging file concurrently).  Treat that race as an idempotent
        # no-op; callers use the boolean to distinguish an actual deletion.
        try:
            metadata = normalized.lstat()
        except FileNotFoundError:
            return False
        if metadata.st_mtime >= cutoff:
            return False
        # Never unlink a hard-linked inode during automatic cleanup: the
        # sibling link may be an unrelated user file.
        if stat.S_ISREG(metadata.st_mode) and getattr(metadata, "st_nlink", 1) != 1:
            return False
        try:
            if stat.S_ISDIR(metadata.st_mode):
                normalized.rmdir()
            else:
                normalized.unlink()
        except FileNotFoundError:
            return False
        return True


class _BoundedOutputStream:
    """限制文件最大长度；超限后丢弃 ZipFile 析构清理写入，避免二次异常。"""

    def __init__(self, raw, max_bytes: int):
        self._raw = raw
        self._max_bytes = max_bytes
        self._max_extent = 0
        self._discarding = False
        self._virtual_position = 0
        self._virtual_size = 0

    @property
    def limit_exceeded(self) -> bool:
        return self._discarding

    def _enter_discard_mode(self) -> None:
        if self._discarding:
            return
        self._virtual_position = self._raw.tell()
        self._virtual_size = max(self._max_extent, self._virtual_position)
        self._discarding = True

    def write(self, data):
        if self._discarding:
            self._virtual_position += len(data)
            self._virtual_size = max(self._virtual_size, self._virtual_position)
            return len(data)

        end_position = self._raw.tell() + len(data)
        if end_position > self._max_bytes:
            self._enter_discard_mode()
            raise OutputSizeLimitExceeded(self._max_bytes)

        written = self._raw.write(data)
        self._max_extent = max(self._max_extent, self._raw.tell())
        return written

    def writelines(self, lines) -> None:
        for line in lines:
            self.write(line)

    def seek(self, offset, whence=0):
        if self._discarding:
            if whence == 0:
                base = 0
            elif whence == 1:
                base = self._virtual_position
            elif whence == 2:
                base = self._virtual_size
            else:
                raise ValueError(f"unsupported whence: {whence}")
            position = base + offset
            if position < 0:
                raise ValueError("negative seek position")
            self._virtual_position = position
            return position

        position = self._raw.seek(offset, whence)
        if position > self._max_bytes:
            self._enter_discard_mode()
            raise OutputSizeLimitExceeded(self._max_bytes)
        return position

    def tell(self):
        if self._discarding:
            return self._virtual_position
        return self._raw.tell()

    def truncate(self, size=None):
        target_size = self.tell() if size is None else size
        if target_size < 0:
            raise ValueError("negative truncate size")
        if self._discarding:
            self._virtual_size = target_size
            self._virtual_position = min(self._virtual_position, target_size)
            return target_size
        if target_size > self._max_bytes:
            self._enter_discard_mode()
            raise OutputSizeLimitExceeded(self._max_bytes)

        result = self._raw.truncate(size)
        self._max_extent = target_size
        return result

    def flush(self):
        if not self._discarding and not self._raw.closed:
            return self._raw.flush()
        return None

    def __getattr__(self, name):
        return getattr(self._raw, name)


def _normalize_output_limit(max_output_bytes) -> int | None:
    if max_output_bytes is None:
        return None
    # Keep the public byte-budget contract exact.  Coercing floats (for
    # example ``1.9`` -> ``1``) silently tightens a caller's limit and makes
    # validation depend on an implementation detail of ``int()``.  Strings
    # and other numeric lookalikes are rejected for the same reason.
    if isinstance(max_output_bytes, bool) or not isinstance(max_output_bytes, int):
        raise ValueError("max_output_bytes must be a positive integer")
    if max_output_bytes <= 0:
        raise ValueError("max_output_bytes must be a positive integer")
    return max_output_bytes


def _ensure_output_parent_directory(output_path: str | Path) -> Path:
    """创建输出目录前拒绝穿过符号链接的父级路径。

    暂存文件必须和最终目标位于同一个受控目录；如果父级目录在
    ``mkdir`` 或 ``NamedTemporaryFile`` 前被解析成符号链接，原子替换
    仍可能把产物写到调用方未预期的位置。逐级用 ``lstat`` 检查已存在
    的祖先，同时允许最后一级目录按需创建。
    """
    output_file = Path(output_path)
    parent = output_file.parent
    probe = parent
    while True:
        try:
            metadata = probe.lstat()
        except FileNotFoundError:
            if probe == probe.parent:
                break
            probe = probe.parent
            continue
        except OSError as exc:
            raise InvalidGeneratedDocxError(
                "output parent directory is unavailable"
            ) from exc
        if stat.S_ISLNK(metadata.st_mode):
            raise InvalidGeneratedDocxError(
                "output parent directory must not be a symbolic link"
            )
        if not stat.S_ISDIR(metadata.st_mode):
            raise InvalidGeneratedDocxError(
                "output parent path is not a directory"
            )
        break

    parent.mkdir(parents=True, exist_ok=True)
    try:
        metadata = parent.lstat()
    except OSError as exc:
        raise InvalidGeneratedDocxError(
            "output parent directory is unavailable"
        ) from exc
    if stat.S_ISLNK(metadata.st_mode) or not stat.S_ISDIR(metadata.st_mode):
        raise InvalidGeneratedDocxError(
            "output parent path is not a private directory"
        )
    return parent


def _create_output_staging_path(output_path: str | Path) -> Path:
    """在目标目录创建同文件系统的私有暂存路径，供原子替换使用。"""
    output_file = Path(output_path)
    parent = _ensure_output_parent_directory(output_file)
    # ``NamedTemporaryFile(dir=...)`` resolves the directory by pathname.  A
    # concurrent replacement of that pathname with a symlink could otherwise
    # create the staging file outside the validated output tree.  Keep the
    # directory identity and verify it again immediately after creation.
    try:
        parent_before = parent.lstat()
    except OSError as exc:
        raise InvalidGeneratedDocxError(
            "output parent directory is unavailable"
        ) from exc
    if not stat.S_ISDIR(parent_before.st_mode):
        raise InvalidGeneratedDocxError("output parent path is not a directory")
    # 创建与登记共用租约锁，清理器不会在两步之间
    # 看到一个尚未受保护的新文件。
    with _ACTIVE_TEMP_PATHS_LOCK:
        parent_changed = False
        with tempfile.NamedTemporaryFile(
            prefix=".docx-output-",
            suffix=".tmp",
            dir=output_file.parent,
            delete=False,
        ) as handle:
            staging_path = Path(handle.name)
            try:
                parent_after = parent.lstat()
            except OSError:
                parent_changed = True
                parent_after = None
            if parent_after is not None and (
                not stat.S_ISDIR(parent_after.st_mode)
                or parent_after.st_dev != parent_before.st_dev
                or parent_after.st_ino != parent_before.st_ino
            ):
                parent_changed = True
        if parent_changed:
            # The context manager has now closed the file.  Only attempt
            # pathname cleanup after confirming the directory identity was
            # restored; otherwise unlinking could target an unrelated
            # directory.  If the pathname remains replaced, leave the staging
            # inode for the owning directory's cleanup instead of following a
            # symlink during error handling.
            try:
                parent_restored = parent.lstat()
            except OSError:
                parent_restored = None
            if (
                parent_restored is not None
                and stat.S_ISDIR(parent_restored.st_mode)
                and parent_restored.st_dev == parent_before.st_dev
                and parent_restored.st_ino == parent_before.st_ino
            ):
                try:
                    staging_path.unlink(missing_ok=True)
                except OSError:
                    pass
            raise InvalidGeneratedDocxError(
                "output parent directory changed during staging"
            )
        try:
            return _register_active_temp_path_locked(staging_path)
        except Exception:
            try:
                staging_path.unlink(missing_ok=True)
            except (IsADirectoryError, OSError):
                # Best-effort cleanup; do not mask registration failures if
                # a callback or concurrent actor replaced the path.
                pass
            raise


def _write_with_output_limit(save_callable, output_path: str | Path, max_output_bytes=None) -> int:
    """直接写入私有暂存路径，并在写入过程中执行可选大小上限。"""
    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    limit = _normalize_output_limit(max_output_bytes)
    if limit is None:
        save_callable(str(output_file))
        return output_file.stat().st_size

    with output_file.open("w+b") as raw:
        bounded = _BoundedOutputStream(raw, limit)
        save_callable(bounded)
        if bounded.limit_exceeded:
            raise OutputSizeLimitExceeded(limit)
        bounded.flush()
        size = raw.seek(0, 2)
        if size > limit:  # pragma: no cover - 写入包装器的最终防线
            raise OutputSizeLimitExceeded(limit)
        return size


def _ensure_private_staging_file(staging_path: str | Path) -> None:
    """确认发布前暂存路径仍是本进程创建的单链接普通文件。

    保存器通常只会写入传入路径，但自定义保存回调或异常清理可能把该
    路径替换成符号链接/硬链接。若直接校验并 ``replace``，校验器可能
    跟随外部目标，破坏原子发布对输出边界的保证。
    """
    try:
        metadata = Path(staging_path).lstat()
    except OSError as exc:
        raise InvalidGeneratedDocxError(
            "output staging path is unavailable"
        ) from exc
    if not stat.S_ISREG(metadata.st_mode) or getattr(metadata, "st_nlink", 1) != 1:
        raise InvalidGeneratedDocxError(
            "output staging path is not a private regular file"
        )


def _cleanup_staging_path(path: str | Path) -> None:
    """Best-effort cleanup that never follows or recursively removes a path.

    A save/validation callback is an extension point and can replace the
    staging pathname with a directory before the caller's ``finally`` block
    runs.  Calling ``Path.unlink`` directly in that case raises
    ``IsADirectoryError`` and masks the real formatting error.  Only unlink
    non-directory entries; leave a replacement directory for the normal
    storage cleanup pass rather than removing user data.
    """
    normalized = Path(path)
    try:
        metadata = normalized.lstat()
    except (FileNotFoundError, OSError):
        return
    if stat.S_ISDIR(metadata.st_mode):
        return
    try:
        normalized.unlink()
    except (FileNotFoundError, IsADirectoryError, OSError):
        # Cleanup is deliberately best effort and must not mask the caller's
        # validation or save exception.
        return


def _save_with_output_limit(
    save_callable,
    output_path: str | Path,
    max_output_bytes=None,
    validate_callable=None,
) -> int:
    """把包原子写入路径；失败时保留已有目标且不暴露新半成品。

    ``validate_callable`` 在暂存文件发布前执行，因此生成器即使成功写出
    一个 ZIP，也不会把结构损坏的产物替换掉已有目标。
    """
    output_file = Path(output_path)
    staging_path = _create_output_staging_path(output_file)
    try:
        size = _write_with_output_limit(
            save_callable,
            staging_path,
            max_output_bytes,
        )
        _ensure_private_staging_file(staging_path)

        if validate_callable is not None:
            try:
                valid = bool(validate_callable(staging_path))
            except Exception as exc:
                raise InvalidGeneratedDocxError(
                    "generated document failed output validation"
                ) from exc
            if not valid:
                raise InvalidGeneratedDocxError(
                    "generated document failed output validation"
                )

        # Validation callbacks are extension points and may replace the path
        # while inspecting it.  Recheck immediately before publish so a
        # symlink/hard-link swap cannot turn the atomic rename into an
        # external-file publish.
        _ensure_private_staging_file(staging_path)
        staging_path.replace(output_file)
        return size
    finally:
        try:
            _cleanup_staging_path(staging_path)
        finally:
            release_active_temp_path(staging_path)


def format_log_path(path) -> str:
    """返回适合日志展示的文件名，避免暴露服务器临时目录或控制字符。"""
    if path is None:
        return "unknown"

    try:
        raw = str(path)
    except Exception:
        return "unknown"

    normalized = "".join(ch for ch in raw.replace("\\", "/") if ch.isprintable()).strip()
    if not normalized:
        return "unknown"

    try:
        name = Path(normalized).name
    except (OSError, ValueError):
        name = normalized.rsplit("/", 1)[-1]

    name = name.strip()[:MAX_LOG_PATH_LENGTH].rstrip()
    return name or "unknown"


def format_log_exception(exc: Exception, *paths) -> str:
    """整理异常消息，同时替换其中可能出现的已知路径。"""
    try:
        message = str(exc)
    except Exception:
        message = ""

    for path in paths:
        try:
            raw = str(path)
        except Exception:
            continue
        if raw:
            message = message.replace(raw, format_log_path(raw))

    message = LOG_PATH_TOKEN_RE.sub(lambda match: format_log_path(match.group(0)), message)
    message = "".join(ch if ch.isprintable() else " " for ch in message)
    message = re.sub(r"\s+", " ", message).strip()
    message = message[:240].rstrip()
    if message:
        return f"{exc.__class__.__name__}: {message}"
    return exc.__class__.__name__


def emit_progress(progress_callback, step: int, message: str, detail: str | None = None):
    """向上层报告排版进度，供 Web SSE 等场景复用。"""
    if progress_callback is None:
        return

    payload = {
        "step": step,
        "message": message,
    }
    if detail:
        payload["detail"] = detail

    try:
        progress_callback(payload)
    except Exception as exc:  # pragma: no cover - 回调失败不应影响主流程
        logger.warning(f"进度回调发送失败：{format_log_exception(exc)}")


def _parse_untrusted_ooxml_part(xml_bytes: bytes):
    """以禁用 DTD、实体与网络访问的方式解析 OOXML 部件。"""
    payload = bytes(xml_bytes)
    normalized = payload.replace(b"\x00", b"").lower()
    if any(marker in normalized for marker in FORBIDDEN_OOXML_DECLARATION_MARKERS):
        raise ValueError("DTD and entity declarations are not allowed in OOXML parts")

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        huge_tree=False,
        recover=False,
    )
    return etree.fromstring(payload, parser=parser)

# ============================================================
# 正则表达式定义（核心匹配逻辑）
# ============================================================

# --- 一级标题匹配 ---
# 匹配规则：数字或数字. + 一个或多个空格 + 至少一个中英文字符
# 示例匹配："1 引言"、"1. 引言"、"2 研究设计"、"3 模型的估计与检验"
# 要求该段落仅包含这一行内容（独占一行），因此使用 ^ 和 $ 锚定
RE_HEADING_L1 = re.compile(
    r"^\d+\.?\s+[A-Za-z\u4e00-\u9fff][\u4e00-\u9fffA-Za-z0-9\s\-—、（）()/&.,:：]*$"
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
RE_ENGLISH_ABSTRACT_HEADING = re.compile(r"^(?:英文摘要|abstract)\s*$", re.IGNORECASE)
RE_ENGLISH_ABSTRACT = re.compile(r"^abstract\s*[:：]\s*", re.IGNORECASE)
RE_ENGLISH_KEYWORDS = re.compile(r"^(?:keywords?|key\s*words?)\s*[:：]?\s*", re.IGNORECASE)
RE_REFERENCES_HEADING = re.compile(r"^参\s*考\s*文\s*献\s*[:：]?\s*$")
RE_CAPTION_NOTE = re.compile(r"^(?:注|说明|资料来源|数据来源|来源|source)\s*[:：]?\s*", re.IGNORECASE)
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
DEFAULT_HEADER_TEXT = ""
RUNNING_HEADER_MAX_LENGTH = 28
CAPTION_MAX_LENGTH = 60
MIN_TOC_HEADING_COUNT = 2
COVER_INFO_FIELDS = [
    ("学院", "college"),
    ("教师", "teacher"),
    ("班级", "class_name"),
    ("姓名", "student_name"),
    ("学号", "student_id"),
]
COVER_LAYOUT = {
    "logo_width_cm": 3.6,
    "school_name_width_cm": 10.6,
    "logo_space_before_pt": 30,
    "logo_space_after_pt": 10,
    "school_name_space_after_pt": 42,
    "title_size_pt": 26,
    "title_space_after_pt": 90,
    "info_spacer_after_pt": 16,
    "info_font_pt": 15,
    "label_width_cm": 2.2,
    "value_width_cm": 7.2,
}
COVER_IMAGE_CANDIDATES = {
    "logo": ("logo.png", "logo.jpg", "logo.jpeg"),
    "school_name": (
        "school_name.png",
        "school_name.jpg",
        "school_name.jpeg",
        "school_name_calligraphy.png",
        "school_name_calligraphy.jpg",
        "浙江工商大学.png",
        "浙江工商大学.jpg",
    ),
}


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
    CAPTION_NOTE = "caption_note"      # 图表附注/来源说明
    SECTION_HEADING = "section_heading"  # 非编号章节标题（如致谢/附录）
    REFERENCES_HEADING = "references_heading"  # 参考文献标题
    REFERENCE_ENTRY = "reference_entry"        # 参考文献条目
    ABSTRACT = "abstract"            # 摘要段落
    KEYWORDS = "keywords"            # 关键词段落
    ENGLISH_ABSTRACT_HEADING = "english_abstract_heading"  # 英文摘要标题
    ENGLISH_ABSTRACT = "english_abstract"                  # 英文摘要正文
    ENGLISH_KEYWORDS = "english_keywords"                  # 英文关键词
    BODY = "body"                    # 正文段落


@dataclass(slots=True)
class ParagraphAnalysis:
    """缓存单个段落的预分析结果，避免主流程重复做文本和 XML 扫描。"""

    index: int
    normalized_text: str
    classified_type: str
    caption_match: tuple[str, re.Match[str], str] | None
    inferred_heading_type: str | None
    has_drawing: bool
    has_equation: bool
    has_rewrite_sensitive_content: bool
    is_reference_entry_candidate: bool
    is_caption_note_candidate: bool


HEADING_LEVEL_BY_TYPE = {
    ParagraphType.HEADING_L1: 0,
    ParagraphType.HEADING_L2: 1,
    ParagraphType.HEADING_L3: 2,
}
HEADING_NUMBER_PATTERNS = {
    ParagraphType.HEADING_L1: re.compile(r"^(?P<number>\d+)(?:\.)?\s+(?P<title>.+)$"),
    ParagraphType.HEADING_L2: re.compile(r"^(?P<number>\d+\.\d+)\s*(?P<title>.+)$"),
    ParagraphType.HEADING_L3: re.compile(r"^(?P<number>\d+\.\d+\.\d+)\s*(?P<title>.+)$"),
}
HEADING_NUMBERING_NSID = "5A475355"
HEADING_NUMBERING_TEMPLATE = "5A475355"
HEADING_NUMBERING_LEVEL_TEXTS = ("%1", "%1.%2", "%1.%2.%3")
MAX_OOXML_DECIMAL_NUMBER = 2_147_483_647
# DrawingML ST_PositiveCoordinate / python-docx 允许的最大 EMU 坐标。
MAX_OOXML_COORDINATE = 27_273_042_316_900
MAX_HEADING_NUMBER_COMPONENT = MAX_OOXML_DECIMAL_NUMBER
UNNUMBERED_HEADING_MAX_LENGTH = 40
HEADING_STYLE_NAME_TO_LEVEL = {
    "heading1": 0,
    "heading2": 1,
    "heading3": 2,
    "标题1": 0,
    "标题2": 1,
    "标题3": 2,
}


# ============================================================
# 段落分类函数
# ============================================================
def match_caption_in_normalized_text(normalized_text: str):
    """识别已规范化文本中的图/表标题。"""
    normalized = (normalized_text or "").strip()
    if not normalized or len(normalized) > CAPTION_MAX_LENGTH:
        return None

    figure_match = RE_FIGURE_CAPTION.match(normalized)
    if figure_match:
        return ParagraphType.FIGURE_CAPTION, figure_match, normalized

    table_match = RE_TABLE_CAPTION.match(normalized)
    if table_match:
        return ParagraphType.TABLE_CAPTION, table_match, normalized

    return None


def classify_normalized_paragraph(
    normalized_text: str,
    *,
    caption_match: tuple[str, re.Match[str], str] | None = None,
) -> str:
    """
    根据已规范化的段落文本判断其类型。

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
        normalized_text: 已经做过空白规范化的段落文本

    Returns:
        ParagraphType 常量字符串
    """
    stripped = (normalized_text or "").strip()

    if not stripped:
        return ParagraphType.BODY  # 空段落当作正文处理

    # 优先匹配摘要和关键词
    if RE_ENGLISH_ABSTRACT_HEADING.match(stripped):
        return ParagraphType.ENGLISH_ABSTRACT_HEADING

    if RE_ABSTRACT.match(stripped):
        return ParagraphType.ABSTRACT

    if RE_KEYWORDS.match(stripped):
        return ParagraphType.KEYWORDS

    if RE_ENGLISH_ABSTRACT.match(stripped):
        return ParagraphType.ENGLISH_ABSTRACT

    if RE_ENGLISH_KEYWORDS.match(stripped):
        return ParagraphType.ENGLISH_KEYWORDS

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

    if caption_match is None:
        caption_match = match_caption_in_normalized_text(stripped)
    if caption_match:
        return caption_match[0]

    # 默认为正文
    return ParagraphType.BODY


def classify_paragraph(text: str) -> str:
    """根据段落纯文本内容判断其类型。"""
    normalized_text = normalize_text_for_matching(text)
    return classify_normalized_paragraph(normalized_text)


def text_exceeds_paragraph_limit(
    text: str,
    max_paragraphs: int = MAX_TEXT_PARAGRAPHS,
) -> bool:
    """以常量内存扫描换行，判断规范化后的有效段落数是否超限。

    CRLF 视为一个分隔符，单独的 CR/LF 也各算一个；末尾连续换行会在
    ``split_text_to_paragraphs`` 中被移除，因此只有后面再次出现正文时才
    把暂存分隔符计入段落数。
    """
    if not isinstance(max_paragraphs, int) or isinstance(max_paragraphs, bool) or max_paragraphs < 1:
        raise ValueError("max_paragraphs must be a positive integer")
    if not isinstance(text, str):
        return False

    paragraph_count = 1
    pending_breaks = 0
    index = 0
    text_length = len(text)
    while index < text_length:
        char = text[index]
        if char == "\r":
            if index + 1 < text_length and text[index + 1] == "\n":
                index += 1
            pending_breaks = min(max_paragraphs + 1, pending_breaks + 1)
        elif char == "\n":
            pending_breaks = min(max_paragraphs + 1, pending_breaks + 1)
        elif pending_breaks:
            paragraph_count += pending_breaks
            if paragraph_count > max_paragraphs:
                return True
            pending_breaks = 0
        index += 1

    return False


def split_text_to_paragraphs(
    text: str,
    max_paragraphs: int = MAX_TEXT_PARAGRAPHS,
) -> list[str]:
    """
    规范化纯文本输入并拆分为段落列表。

    - 统一处理 Windows/macOS/Linux 换行符
    - 保留中间空行，便于在导出的文档中保留段落间距
    - 避免把换行控制字符残留到段落正文里
    """
    if text_exceeds_paragraph_limit(text, max_paragraphs=max_paragraphs):
        raise TextParagraphLimitExceeded(max_paragraphs)

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


def _paragraph_text_for_matching(paragraph) -> str:
    """提取段落中可见文字，覆盖内容控件和简单域的结果文本。

    ``python-docx`` 的 ``Paragraph.text`` 只遍历段落直属 run/hyperlink；
    Word 常见的内容控件（``w:sdt``）和简单域（``w:fldSimple``）中的结果
    run 因此会被漏掉，导致标题/摘要等结构识别失败。这里按 OOXML 顺序
    读取可见文本节点，同时跳过修订删除和域指令，保留制表符/换行符。
    """
    text_parts = []
    hidden_subtrees = {
        qn("w:del"),
        qn("w:moveFrom"),
        qn("w:instrText"),
        qn("w:delText"),
    }
    def visit(node):
        tag = getattr(node, "tag", None)
        if tag in hidden_subtrees:
            return
        if tag == qn("w:t"):
            text_parts.append(node.text or "")
        elif tag == qn("w:noBreakHyphen"):
            text_parts.append("\u2011")
        elif tag == qn("w:softHyphen"):
            text_parts.append("\u00ad")
        elif tag in {qn("w:tab"), qn("w:ptab")}:
            text_parts.append("\t")
        elif tag in {qn("w:br"), qn("w:cr")}:
            text_parts.append("\n")
        for child in node:
            visit(child)

    visit(paragraph._element)
    return "".join(text_parts)


def match_caption(text: str):
    """
    识别图/表标题，并返回类型、匹配结果和规范化后的文本。

    通过长度限制降低误把正文识别为图表标题的风险。
    """
    normalized = normalize_text_for_matching(text)
    return match_caption_in_normalized_text(normalized)


def rebuild_caption_text(kind: str, number: int, match: re.Match) -> str:
    """按照统一编号规则重建图/表标题文本。"""
    label = "图" if kind == ParagraphType.FIGURE_CAPTION else "表"
    caption = (match.group("caption") or "").strip()
    return f"{label} {number}" if not caption else f"{label} {number} {caption}"


def is_caption_note_candidate(text: str, normalized_text: str | None = None) -> bool:
    """判断段落是否像图表后的附注/来源说明。"""
    normalized = normalized_text if normalized_text is not None else normalize_text_for_matching(text)
    return bool(RE_CAPTION_NOTE.match(normalized))


def is_title_candidate(
    text: str,
    *,
    normalized_text: str | None = None,
    classified_type: str | None = None,
) -> bool:
    """判断一个段落是否像论文标题。"""
    stripped = normalized_text if normalized_text is not None else normalize_text_for_matching(text)

    if not stripped:
        return False

    if not 6 <= len(stripped) <= 40:
        return False

    if stripped.endswith(("。", "！", "？", "!", "?", "；", ";")):
        return False

    if "@" in stripped or RE_TITLE_METADATA_PREFIX.match(stripped):
        return False

    para_type = classified_type if classified_type is not None else classify_normalized_paragraph(stripped)
    return para_type == ParagraphType.BODY


def is_reference_entry_text(text: str, normalized_text: str | None = None) -> bool:
    """判断参考文献段落是否具备常见的编号前缀。"""
    normalized = normalized_text if normalized_text is not None else normalize_text_for_matching(text)
    return bool(RE_REFERENCE_ENTRY_TEXT.match(normalized))


def find_title_paragraph_index(paragraphs, analyses: list[ParagraphAnalysis] | None = None) -> int | None:
    """
    尝试识别论文主标题。

    规则保持保守：
    - 只考虑前 3 个非空段落中的候选项
    - 段落本身必须像标题
    - 后续 3 个非空段落内需要出现“摘要”或“关键词”
    """
    if analyses is None:
        non_empty = []
        for index, paragraph in enumerate(paragraphs):
            normalized_text = normalize_text_for_matching(
                _paragraph_text_for_matching(paragraph)
            )
            if not normalized_text:
                continue
            non_empty.append(
                (
                    index,
                    normalized_text,
                    classify_normalized_paragraph(normalized_text),
                )
            )
    else:
        non_empty = [
            (analysis.index, analysis.normalized_text, analysis.classified_type)
            for analysis in analyses
            if analysis.normalized_text
        ]

    if not non_empty:
        return None

    for candidate_position, (candidate_index, candidate_text, candidate_type) in enumerate(non_empty[:3]):
        if not is_title_candidate(
            candidate_text,
            normalized_text=candidate_text,
            classified_type=candidate_type,
        ):
            continue

        for _, text, para_type in non_empty[candidate_position + 1:candidate_position + 4]:
            if para_type in {
                ParagraphType.ABSTRACT,
                ParagraphType.KEYWORDS,
                ParagraphType.ENGLISH_ABSTRACT_HEADING,
                ParagraphType.ENGLISH_ABSTRACT,
                ParagraphType.ENGLISH_KEYWORDS,
            }:
                return candidate_index

            if para_type in {
                ParagraphType.HEADING_L1,
                ParagraphType.HEADING_L2,
                ParagraphType.FIGURE_CAPTION,
                ParagraphType.TABLE_CAPTION,
            }:
                break

    return None


def _parse_bounded_decimal(value, maximum: int, minimum: int = 0) -> int | None:
    """解析 OOXML 十进制属性，先限长再调用 int()，避免超长数字放大。"""
    if not isinstance(value, str):
        return None

    digits = value.strip()
    if digits.startswith("+"):
        digits = digits[1:]
    if not digits or not digits.isascii() or not digits.isdecimal():
        return None

    significant_digits = digits.lstrip("0") or "0"
    maximum_text = str(maximum)
    if (
        len(significant_digits) > len(maximum_text)
        or (
            len(significant_digits) == len(maximum_text)
            and significant_digits > maximum_text
        )
    ):
        return None

    parsed = int(significant_digits)
    return parsed if parsed >= minimum else None


def _normalize_ooxml_integer(value, max_digits: int = 128) -> str | None:
    """规范化 XSD integer，不构造无界大整数。"""
    if not isinstance(value, str):
        return None
    candidate = value.strip()
    if not candidate:
        return None
    sign = ""
    if candidate[0] in {"+", "-"}:
        sign, candidate = candidate[0], candidate[1:]
    if (
        not candidate
        or len(candidate) > max_digits
        or not candidate.isascii()
        or not candidate.isdecimal()
    ):
        return None
    magnitude = candidate.lstrip("0") or "0"
    if magnitude == "0":
        return "0"
    return f"-{magnitude}" if sign == "-" else magnitude


_XSD_DOUBLE_LEXICAL_RE = re.compile(
    r"(?:[+-]?(?:(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)"
    r"(?:[eE][+-]?[0-9]+)?)|INF|-INF|NaN)",
    flags=re.ASCII,
)


def _parse_xsd_int32(value) -> int | None:
    """解析 XSD Int32 词法，同时避免把超长整数交给 ``int()``。"""
    if not isinstance(value, str):
        return None
    candidate = value.strip(" \t\r\n")
    if not candidate:
        return None
    negative = candidate.startswith("-")
    if candidate[0] in {"+", "-"}:
        candidate = candidate[1:]
    if not candidate or not candidate.isascii() or not candidate.isdecimal():
        return None
    magnitude = candidate.lstrip("0") or "0"
    if len(magnitude) > 10:
        return None
    parsed = int(magnitude)
    if negative:
        parsed = -parsed
    if -0x80000000 <= parsed <= 0x7FFFFFFF:
        return parsed
    return None


def _parse_xsd_uint32(value) -> int | None:
    """解析 XSD UInt32 词法，允许任意数量的前导零。"""
    if not isinstance(value, str):
        return None
    candidate = value.strip(" \t\r\n")
    negative = candidate.startswith("-")
    if candidate.startswith(("+", "-")):
        candidate = candidate[1:]
    if not candidate or not candidate.isascii() or not candidate.isdecimal():
        return None
    magnitude = candidate.lstrip("0") or "0"
    if len(magnitude) > 10:
        return None
    parsed = int(magnitude)
    if negative and parsed:
        return None
    return parsed if parsed <= 0xFFFFFFFF else None


def _parse_xsd_int64(value) -> int | None:
    """解析 XSD Int64 词法，允许 XML whitespace 和前导零。"""
    if not isinstance(value, str):
        return None
    candidate = value.strip(" \t\r\n")
    if not candidate:
        return None
    negative = candidate.startswith("-")
    if candidate[0] in {"+", "-"}:
        candidate = candidate[1:]
    if not candidate or not candidate.isascii() or not candidate.isdecimal():
        return None
    magnitude = candidate.lstrip("0") or "0"
    if len(magnitude) > 19:
        return None
    parsed = int(magnitude)
    if negative:
        parsed = -parsed
    if -0x8000000000000000 <= parsed <= 0x7FFFFFFFFFFFFFFF:
        return parsed
    return None


def _is_xsd_double_lexical(value) -> bool:
    """判断文本是否属于 XSD double 的标准词法空间。"""
    if not isinstance(value, str):
        return False
    candidate = value.strip(" \t\r\n")
    return bool(candidate and _XSD_DOUBLE_LEXICAL_RE.fullmatch(candidate))


def _has_xsd_list_items(value) -> bool:
    return isinstance(value, str) and bool(
        value.strip(" \t\r\n")
    )


def _is_xsd_boolean(value) -> bool:
    return (
        isinstance(value, str)
        and value.strip(" \t\r\n") in _XSD_BOOLEAN_VALUES
    )


def _parse_vml_dimension(value) -> float | None:
    """解析 VML 尺寸数字，避免超长或非有限浮点值进入缩放计算。"""
    if not isinstance(value, str) or len(value) > 64:
        return None

    candidate = value.strip()
    if not re.fullmatch(r"(?:\d+(?:\.\d*)?|\.\d+)", candidate, flags=re.ASCII):
        return None

    try:
        parsed = float(candidate)
    except (OverflowError, ValueError):
        return None
    return parsed if math.isfinite(parsed) and parsed > 0 else None


def _is_negative_ooxml_decimal(value) -> bool:
    """识别 OOXML 中的负十进制数，无需构造任意精度整数。"""
    if not isinstance(value, str):
        return False

    candidate = value.strip()
    if not candidate.startswith("-"):
        return False

    magnitude = candidate[1:]
    return (
        bool(magnitude)
        and magnitude.isascii()
        and magnitude.isdecimal()
        and any(digit != "0" for digit in magnitude)
    )


def extract_heading_numbering(text: str, para_type: str) -> tuple[str, tuple[int, ...]]:
    """
    从标题文本中提取编号前缀和纯标题文本。

    例如：
      - "1 引言"    -> ("引言", (1,))
      - "1.1 背景"  -> ("背景", (1, 1))
      - "1.1.1 假设" -> ("假设", (1, 1, 1))
    """
    normalized = normalize_text_for_matching(text)
    pattern = HEADING_NUMBER_PATTERNS.get(para_type)
    if pattern is None:
        return normalized, ()

    match = pattern.match(normalized)
    if match is None:
        return normalized, ()

    number_text = (match.group("number") or "").rstrip(".")
    title_text = normalize_text_for_matching(match.group("title"))
    if not number_text or not title_text:
        return normalized, ()

    number_parts = []
    for part in number_text.split("."):
        parsed_part = _parse_bounded_decimal(part, MAX_HEADING_NUMBER_COMPONENT)
        if parsed_part is None:
            return normalized, ()
        number_parts.append(parsed_part)

    return title_text, tuple(number_parts)


def _read_outline_level(p_pr) -> int | None:
    """从段落或样式的 pPr 中读取 outline level。"""
    if p_pr is None:
        return None

    outline_level = p_pr.find(qn("w:outlineLvl"))
    if outline_level is None:
        return None

    raw_value = outline_level.get(qn("w:val"))
    if raw_value is None:
        return None

    return _parse_bounded_decimal(raw_value, MAX_OOXML_DECIMAL_NUMBER)


def _build_paragraph_style_metadata_index(document_part):
    """一次性索引段落样式，避免每个段落都 XPath 扫描 styles.xml。"""
    styles_element = document_part.styles.element
    metadata_by_id = {}
    default_metadata = ("", None)

    for style_element in styles_element.findall(qn("w:style")):
        if style_element.get(qn("w:type")) != "paragraph":
            continue

        name_element = style_element.find(qn("w:name"))
        style_name = name_element.get(qn("w:val"), "") if name_element is not None else ""
        outline_level = _read_outline_level(style_element.find(qn("w:pPr")))
        metadata = (style_name, outline_level)
        style_id = style_element.get(qn("w:styleId"))
        if style_id and style_id not in metadata_by_id:
            metadata_by_id[style_id] = metadata

        if style_element.get(qn("w:default")) in {"1", "true", "on"}:
            default_metadata = metadata

    return metadata_by_id, default_metadata


def _get_paragraph_style_metadata(paragraph) -> tuple[str, int | None]:
    part = paragraph.part
    document_part = getattr(part, "_document_part", part)
    cached = getattr(document_part, "_academic_paragraph_style_metadata", None)
    if cached is None:
        cached = _build_paragraph_style_metadata_index(document_part)
        setattr(document_part, "_academic_paragraph_style_metadata", cached)

    metadata_by_id, default_metadata = cached
    p_pr = paragraph._element.pPr
    p_style = p_pr.find(qn("w:pStyle")) if p_pr is not None else None
    style_id = p_style.get(qn("w:val")) if p_style is not None else None
    return metadata_by_id.get(style_id, default_metadata)


def _ensure_normal_paragraph_style(doc):
    """确保保留样式 Normal 存在且为段落样式。"""
    if hasattr(doc.part, "_academic_paragraph_style_metadata"):
        delattr(doc.part, "_academic_paragraph_style_metadata")
    styles_element = doc.styles.element
    normal_style_element = next(
        (
            style_element
            for style_element in styles_element.findall(qn("w:style"))
            if style_element.get(qn("w:styleId")) == "Normal"
        ),
        None,
    )

    if normal_style_element is None:
        normal_style = doc.styles.add_style("Normal", WD_STYLE_TYPE.PARAGRAPH)
    else:
        normal_style_element.set(qn("w:type"), "paragraph")
        normal_style = doc.styles["Normal"]

    # 后续每个段落都会重置为 Normal。缓存已验证的 styleId，
    # 避免 python-docx 为数千个段落重复按名称扫描 styles.xml。
    doc.part._academic_normal_paragraph_style_id = normal_style.style_id
    return normal_style


def _get_paragraph_outline_level_hint(paragraph) -> int | None:
    """
    获取段落自带的标题层级提示。

    优先读取段落直接设置的 outline level；如果没有，再回退到样式中配置的
    outline level 或常见的 Heading 样式名称。
    """
    direct_level = _read_outline_level(paragraph._element.pPr)
    if direct_level in {0, 1, 2}:
        return direct_level

    style_name, style_level = _get_paragraph_style_metadata(paragraph)
    if style_level in {0, 1, 2}:
        return style_level

    style_name = re.sub(r"\s+", "", style_name or "")
    return HEADING_STYLE_NAME_TO_LEVEL.get(style_name.lower(), HEADING_STYLE_NAME_TO_LEVEL.get(style_name))


def looks_like_unnumbered_heading(text: str, normalized_text: str | None = None) -> bool:
    """保守判断一段文字是否像“缺失编号的标题”。"""
    normalized = normalized_text if normalized_text is not None else normalize_text_for_matching(text)
    if not normalized or len(normalized) > UNNUMBERED_HEADING_MAX_LENGTH:
        return False

    if normalized.endswith(("。", "！", "？", "!", "?", "；", ";", "，", ",", "：", ":")):
        return False

    if is_reference_entry_text(normalized, normalized_text=normalized):
        return False

    if match_caption_in_normalized_text(normalized):
        return False

    return True


def infer_heading_type_from_paragraph(
    paragraph,
    text: str,
    normalized_text: str | None = None,
) -> str | None:
    """根据段落原始样式/outline level，推断未编号标题的层级。"""
    normalized = normalized_text if normalized_text is not None else normalize_text_for_matching(text)
    if not looks_like_unnumbered_heading(text, normalized_text=normalized):
        return None

    level = _get_paragraph_outline_level_hint(paragraph)
    return {
        0: ParagraphType.HEADING_L1,
        1: ParagraphType.HEADING_L2,
        2: ParagraphType.HEADING_L3,
    }.get(level)


def _build_paragraph_analyses(paragraphs) -> list[ParagraphAnalysis]:
    """预分析所有正文段落，复用规范化文本、分类结果与公式探测。"""
    analyses = []
    for index, paragraph in enumerate(paragraphs):
        raw_text = _paragraph_text_for_matching(paragraph)
        normalized_text = normalize_text_for_matching(raw_text)
        caption_match = match_caption_in_normalized_text(normalized_text)
        classified_type = classify_normalized_paragraph(
            normalized_text,
            caption_match=caption_match,
        )
        has_drawing, has_equation = _scan_paragraph_content(paragraph)
        inferred_heading_type = None
        if classified_type == ParagraphType.BODY:
            inferred_heading_type = infer_heading_type_from_paragraph(
                paragraph,
                raw_text,
                normalized_text=normalized_text,
            )

        analyses.append(
            ParagraphAnalysis(
                index=index,
                normalized_text=normalized_text,
                classified_type=classified_type,
                caption_match=caption_match,
                inferred_heading_type=inferred_heading_type,
                has_drawing=has_drawing,
                has_equation=has_equation,
                has_rewrite_sensitive_content=(
                    _has_rewrite_sensitive_inline_content(paragraph)
                ),
                is_reference_entry_candidate=is_reference_entry_text(
                    normalized_text,
                    normalized_text=normalized_text,
                ),
                is_caption_note_candidate=is_caption_note_candidate(
                    normalized_text,
                    normalized_text=normalized_text,
                ),
            )
        )

    return analyses


def _log_detected_paragraph(label: str, index: int, text: str) -> None:
    """逐段调试日志仅在 DEBUG 下输出，避免大文档时产生过多日志 I/O。"""
    if not logger.isEnabledFor(logging.DEBUG):
        return

    preview = text if len(text) <= 30 else f"{text[:30]}..."
    logger.debug(f'  [{label}] 第{index + 1}段: "{preview}"')


def resolve_heading_numbering_parts(
    para_type: str,
    explicit_parts: tuple[int, ...],
    numbering_state: list[int],
    *,
    allow_auto_numbering: bool = False,
) -> tuple[int, ...]:
    """综合显式编号和推断层级，返回当前标题应使用的原生编号。"""
    level = HEADING_LEVEL_BY_TYPE.get(para_type)
    if level is None:
        return ()

    if explicit_parts:
        for index in range(len(numbering_state)):
            numbering_state[index] = explicit_parts[index] if index < len(explicit_parts) else 0
        return explicit_parts

    if not allow_auto_numbering:
        return ()

    if level > 0 and any(numbering_state[index] <= 0 for index in range(level)):
        return ()

    numbering_state[level] = numbering_state[level] + 1 if numbering_state[level] > 0 else 1
    for index in range(level + 1, len(numbering_state)):
        numbering_state[index] = 0

    return tuple(numbering_state[: level + 1])


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
    else:
        r_pr.remove(r_fonts)
    r_pr._insert_rFonts(r_fonts)

    r_fonts.set(qn("w:eastAsia"), cn_font)
    r_fonts.set(qn("w:ascii"), en_font)
    r_fonts.set(qn("w:hAnsi"), en_font)
    r_fonts.set(qn("w:cs"), en_font)  # 复杂脚本字体也设为西文字体


def _remove_all_runs(paragraph):
    """删除段落中全部 run（含超链接等容器内的 run），便于安全重建文本。"""
    for run_element in list(paragraph._element.iter(qn("w:r"))):
        parent = run_element.getparent()
        if parent is not None:
            parent.remove(run_element)

    # 标题重建后不应留下空的超链接容器（及其无效的外部链接引用）。
    for hyperlink in list(paragraph._element.iter(qn("w:hyperlink"))):
        if not any(child.tag == qn("w:r") for child in hyperlink.iter()):
            parent = hyperlink.getparent()
            if parent is not None:
                parent.remove(hyperlink)


def _replace_paragraph_text(paragraph, text: str):
    """安全重建纯文本段落内容。仅用于标题、图表标题等纯文本段落。"""
    _remove_all_runs(paragraph)
    if text:
        paragraph.add_run(text)


def _iter_paragraph_runs(paragraph):
    """遍历段落直接拥有的 run，包括超链接/字段容器中的 run。"""
    paragraph_element = paragraph._element
    for run_element in paragraph_element.iter(qn("w:r")):
        nearest_paragraph = run_element
        while nearest_paragraph is not None and nearest_paragraph.tag != qn("w:p"):
            nearest_paragraph = nearest_paragraph.getparent()
        if nearest_paragraph is paragraph_element:
            yield Run(run_element, paragraph)


def _apply_run_fonts(paragraph, cn_font: str, en_font: str, size_pt: float, bold: bool | None = None):
    """遍历段落中所有 run，统一设置中西文字体。"""
    for run in _iter_paragraph_runs(paragraph):
        if bold is not None:
            current_bold = bold
        else:
            try:
                current_bold = run.font.bold
            except (InvalidXmlError, OxmlInvalidXmlError, ValueError):
                # 非法 w:b@w:val 不应让整份文档排版失败。
                r_pr = run._element.rPr
                bold_element = r_pr.find(qn("w:b")) if r_pr is not None else None
                if bold_element is not None:
                    r_pr.remove(bold_element)
                current_bold = None
        _set_run_font(run, cn_font=cn_font, en_font=en_font, size_pt=size_pt, bold=current_bold)


def _is_list_paragraph(paragraph) -> bool:
    """判断段落是否为项目符号/编号列表，尽量保留其原有列表样式和缩进。"""
    style_name = ""
    style_name, _ = _get_paragraph_style_metadata(paragraph)
    style_name = (style_name or "").lower()

    if style_name.startswith("list"):
        return True

    p_pr = paragraph._element.pPr
    return p_pr is not None and p_pr.numPr is not None


def _clear_paragraph_numbering(paragraph):
    """移除段落原有的编号定义，避免旧模板列表样式残留。"""
    p_pr = paragraph._element.pPr
    if p_pr is None:
        return

    num_pr = p_pr.find(qn("w:numPr"))
    if num_pr is not None:
        p_pr.remove(num_pr)


def iter_unique_row_cells(row):
    """逐个包装真实 w:tc，避免 row.cells 按 gridSpan 展开并递归解析纵向合并。"""
    for cell_element in row._tr.tc_lst:
        yield _Cell(cell_element, row.table)


def iter_table_paragraphs(tables):
    """递归遍历所有表格单元格内的段落。"""
    for table in tables:
        for row in table.rows:
            for cell in iter_unique_row_cells(row):
                for paragraph in cell.paragraphs:
                    yield paragraph
                yield from iter_table_paragraphs(cell.tables)


def iter_all_tables(tables):
    """递归遍历文档中的所有表格（含嵌套表格）。"""
    for table in tables:
        yield table
        for row in table.rows:
            for cell in iter_unique_row_cells(row):
                yield from iter_all_tables(cell.tables)


def _has_drawing(paragraph) -> bool:
    """判断段落中是否包含图片等 drawing 元素。"""
    return _element_contains_any_tag(paragraph._element, DRAWING_XML_TAGS)


def _has_equation_content(paragraph) -> bool:
    """判断段落中是否包含 Word 公式或嵌入式公式对象。"""
    return _element_contains_any_tag(paragraph._element, EQUATION_XML_TAGS)


def _element_contains_any_tag(element, tags: frozenset[str]) -> bool:
    """判断 XML 子树中是否包含任一目标标签。"""
    return any(getattr(node, "tag", None) in tags for node in element.iter())


def _scan_paragraph_content(paragraph) -> tuple[bool, bool]:
    """单次扫描段落 XML，同时探测图片/对象与公式内容。"""
    has_drawing = False
    has_equation = False

    for node in paragraph._element.iter():
        tag = getattr(node, "tag", None)
        if tag in DRAWING_XML_TAGS:
            has_drawing = True
        if tag in EQUATION_XML_TAGS:
            has_equation = True
        if has_drawing and has_equation:
            break

    return has_drawing, has_equation


def _run_has_rewrite_sensitive_content(run_element) -> bool:
    """判断 run 是否含字段、引用、对象等不能靠纯文本等价重建的节点。"""
    for child in run_element:
        tag = getattr(child, "tag", None)
        if tag not in REWRITE_SAFE_RUN_CHILD_TAGS:
            return True
        # A plain line break can be normalized, but page/column breaks and
        # wrapping-clear directives cannot be represented by plain text.
        if tag == qn("w:br") and (
            child.get(qn("w:type"), "textWrapping") != "textWrapping"
            or child.get(qn("w:clear"), "none") != "none"
        ):
            return True
        if tag == qn("w:rPr") and any(
            getattr(property_child, "tag", None)
            not in REWRITE_SAFE_RUN_PROPERTY_TAGS
            for property_child in child
        ):
            return True
    return False


def _has_rewrite_sensitive_inline_content(paragraph) -> bool:
    """判断段落是否含重建纯文本时会丢失或错位的行内语义。

    纯文本 run、换行/制表符及普通超链接仍沿用既有规范化行为；书签、
    字段、脚注/批注引用、内容控件、修订内容和未知扩展节点则一律保留。
    """
    paragraph_element = paragraph._element
    for child in paragraph_element:
        tag = getattr(child, "tag", None)
        if tag in {qn("w:pPr"), qn("w:proofErr")}:
            continue
        if tag == qn("w:r"):
            if _run_has_rewrite_sensitive_content(child):
                return True
            continue
        if tag == qn("w:hyperlink"):
            for hyperlink_child in child:
                hyperlink_tag = getattr(hyperlink_child, "tag", None)
                if hyperlink_tag == qn("w:proofErr"):
                    continue
                if (
                    hyperlink_tag != qn("w:r")
                    or _run_has_rewrite_sensitive_content(hyperlink_child)
                ):
                    return True
            continue
        return True
    return False


def _insert_xml_child_before(parent, child, successor_names=()):
    """在普通 lxml 元素中按 OOXML 子节点次序插入。"""
    successor_tags = {qn(name) for name in successor_names}
    successor = next((item for item in parent if item.tag in successor_tags), None)
    if successor is None:
        parent.append(child)
    else:
        successor.addprevious(child)
    return child


def _ensure_xml_child(parent, tag_name: str, *, prepend: bool = False, successors=()):
    """确保底层 XML 节点存在且位于其 schema 后继节点之前。"""
    matching_children = list(parent.findall(qn(tag_name)))
    child = matching_children[0] if matching_children else None
    changed = False
    if child is None:
        child = OxmlElement(tag_name)
        if prepend:
            parent.insert(0, child)
        elif successors:
            _insert_xml_child_before(parent, child, successors)
        else:
            parent.append(child)
        return child, True

    for duplicate in matching_children[1:]:
        parent.remove(duplicate)
        changed = True

    if prepend and parent.index(child) != 0:
        parent.remove(child)
        parent.insert(0, child)
        changed = True
    elif successors:
        successor_tags = {qn(name) for name in successors}
        child_index = parent.index(child)
        successor = next(
            (
                item
                for index, item in enumerate(parent)
                if index < child_index and item.tag in successor_tags
            ),
            None,
        )
        if successor is not None:
            parent.remove(child)
            successor.addprevious(child)
            changed = True

    return child, changed


def _set_xml_attribute(element, attr_name: str, value: str) -> bool:
    """仅在属性值变化时写入，便于统计 XML 是否被修改。"""
    attr = qn(attr_name)
    if element.get(attr) == value:
        return False

    element.set(attr, value)
    return True


def _format_footnote_run_xml(run) -> bool:
    """统一脚注 run 的中西文字体和字号。"""
    r_pr, changed = _ensure_xml_child(run, "w:rPr", prepend=True)

    r_fonts, node_changed = _ensure_xml_child(
        r_pr,
        "w:rFonts",
        successors=(
            "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps",
            "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss",
            "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden",
            "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz",
            "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
            "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
            "w:eastAsianLayout", "w:specVanish", "w:oMath",
        ),
    )
    changed |= node_changed

    changed |= _set_xml_attribute(r_fonts, "w:eastAsia", "宋体")
    changed |= _set_xml_attribute(r_fonts, "w:ascii", "Times New Roman")
    changed |= _set_xml_attribute(r_fonts, "w:hAnsi", "Times New Roman")
    changed |= _set_xml_attribute(r_fonts, "w:cs", "Times New Roman")

    size_val = str(int(FOOTNOTE_FONT_SIZE_PT * 2))
    size, node_changed = _ensure_xml_child(
        r_pr,
        "w:sz",
        successors=(
            "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
            "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
            "w:eastAsianLayout", "w:specVanish", "w:oMath",
        ),
    )
    changed |= node_changed
    size_cs, node_changed = _ensure_xml_child(
        r_pr,
        "w:szCs",
        successors=(
            "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText",
            "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
            "w:eastAsianLayout", "w:specVanish", "w:oMath",
        ),
    )
    changed |= node_changed
    changed |= _set_xml_attribute(size, "w:val", size_val)
    changed |= _set_xml_attribute(size_cs, "w:val", size_val)

    if run.find(qn("w:footnoteRef")) is not None or run.find(qn("w:footnoteReference")) is not None:
        r_style, node_changed = _ensure_xml_child(r_pr, "w:rStyle", prepend=True)
        changed |= node_changed
        vert_align, node_changed = _ensure_xml_child(
            r_pr,
            "w:vertAlign",
            successors=(
                "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
                "w:specVanish", "w:oMath",
            ),
        )
        changed |= node_changed
        changed |= _set_xml_attribute(r_style, "w:val", "FootnoteReference")
        changed |= _set_xml_attribute(vert_align, "w:val", "superscript")

    return changed


def _format_footnote_paragraph_xml(paragraph) -> bool:
    """统一脚注段落的行距、对齐和缩进。"""
    p_pr, changed = _ensure_xml_child(paragraph, "w:pPr", prepend=True)

    spacing, node_changed = _ensure_xml_child(
        p_pr,
        "w:spacing",
        successors=(
            "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
            "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
            "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
        ),
    )
    changed |= node_changed
    changed |= _set_xml_attribute(spacing, "w:before", "0")
    changed |= _set_xml_attribute(spacing, "w:after", "0")
    changed |= _set_xml_attribute(spacing, "w:line", "240")
    changed |= _set_xml_attribute(spacing, "w:lineRule", "auto")

    ind, node_changed = _ensure_xml_child(
        p_pr,
        "w:ind",
        successors=(
            "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
            "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
            "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
        ),
    )
    changed |= node_changed
    changed |= _set_xml_attribute(ind, "w:left", "0")
    changed |= _set_xml_attribute(ind, "w:right", "0")
    changed |= _set_xml_attribute(ind, "w:firstLine", "0")
    if ind.get(qn("w:hanging")) is not None:
        del ind.attrib[qn("w:hanging")]
        changed = True

    jc, node_changed = _ensure_xml_child(
        p_pr,
        "w:jc",
        successors=(
            "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
            "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
        ),
    )
    changed |= node_changed
    changed |= _set_xml_attribute(jc, "w:val", "left")

    widow_control, node_changed = _ensure_xml_child(
        p_pr,
        "w:widowControl",
        successors=(
            "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs",
            "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
            "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi",
            "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
            "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
            "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
            "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
        ),
    )
    changed |= node_changed
    changed |= _set_xml_attribute(widow_control, "w:val", "true")

    return changed


def _rewrite_docx_part(
    docx_path: str | Path,
    part_name: str,
    transform,
    max_output_bytes=None,
    *,
    atomic_write: bool = True,
    source_stream=None,
    output_validation_limits: DocxValidationLimits | None = None,
) -> int:
    """
    重写 docx 中指定部件。

    transform 回调返回 `(new_bytes, count, changed)`。
    """
    docx_file = Path(docx_path)
    output_limit = _normalize_output_limit(max_output_bytes)
    if source_stream is None:
        if not docx_file.exists():
            return 0
        source_size = docx_file.stat().st_size
    else:
        source_stream.seek(0, 2)
        source_size = source_stream.tell()
        source_stream.seek(0)
    if output_limit is not None and source_size > output_limit:
        raise OutputSizeLimitExceeded(output_limit)

    rewrite_path = None
    try:
        source_target = docx_file if source_stream is None else source_stream
        with ZipFile(source_target, "r") as source:
            try:
                original_bytes = source.read(part_name)
            except KeyError:
                return 0

            new_bytes, count, changed = transform(original_bytes)
            if not changed:
                return count

            entries = source.infolist()
            archive_comment = source.comment
            rewrite_path = _create_output_staging_path(docx_file)

            def write_archive(target_stream_or_path):
                with ZipFile(target_stream_or_path, "w") as target:
                    target.comment = archive_comment
                    for entry in entries:
                        target_entry = deepcopy(entry)
                        if entry.filename == part_name:
                            target.writestr(target_entry, new_bytes)
                            continue
                        with (
                            source.open(entry, "r") as source_part,
                            target.open(target_entry, "w") as target_part,
                        ):
                            while chunk := source_part.read(1024 * 1024):
                                target_part.write(chunk)

            # 即使调用方已在外层暂存（atomic_write=False），也不能在读取
            # 原 ZIP 时原地截断它；始终先写同目录私有文件，关闭源包后再替换。
            _write_with_output_limit(
                write_archive,
                rewrite_path,
                output_limit,
            )
            _ensure_private_staging_file(rewrite_path)
            if output_validation_limits is not None:
                with rewrite_path.open("rb") as rewritten_stream:
                    if not is_valid_docx_stream(
                        rewritten_stream,
                        limits=output_validation_limits,
                    ):
                        raise InvalidGeneratedDocxError(
                            "rewritten document failed DOCX safety validation"
                        )

        # Keep the final rename guarded even when validation succeeds: the
        # validator opens an attacker-controlled path and can race its inode.
        _ensure_private_staging_file(rewrite_path)
        rewrite_path.replace(docx_file)
        return count
    finally:
        if rewrite_path is not None:
            try:
                _cleanup_staging_path(rewrite_path)
            finally:
                release_active_temp_path(rewrite_path)


def _related_part_uri_from_archive(archive, source_uri: PackURI, relationship_type: str):
    """按 OPC relationship 返回目标 Part URI；缺失或歧义时返回 None。"""
    relationship_part_name = str(source_uri.rels_uri).lstrip("/")
    try:
        relationship_root = _parse_untrusted_ooxml_part(
            archive.read(relationship_part_name)
        )
    except KeyError:
        return None

    relationships_tag = (
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationships"
    )
    relationship_tag = (
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    )
    if relationship_root.tag != relationships_tag:
        return None

    matches = [
        relationship
        for relationship in relationship_root.findall(relationship_tag)
        if relationship.get("Type") == relationship_type
        and relationship.get("TargetMode") != "External"
    ]
    if len(matches) != 1:
        return None

    target = matches[0].get("Target")
    if not target:
        return None
    target_uri = PackURI.from_rel_ref(source_uri.baseURI, target)
    target_name = str(target_uri).lstrip("/")
    try:
        if archive.getinfo(target_name).is_dir():
            return None
    except KeyError:
        return None
    return target_uri


def _resolve_footnote_part_name_from_archive(archive) -> str | None:
    """从已打开的 package relationships 解析实际脚注 PartName。"""
    main_document_uri = _related_part_uri_from_archive(
        archive,
        PackURI("/"),
        RT.OFFICE_DOCUMENT,
    )
    if main_document_uri is None:
        return None
    footnote_uri = _related_part_uri_from_archive(
        archive,
        main_document_uri,
        RT.FOOTNOTES,
    )
    return str(footnote_uri).lstrip("/") if footnote_uri is not None else None


def _resolve_footnote_part_name(docx_path: str | Path) -> str | None:
    """从 package relationships 解析实际脚注 PartName，兼容重定位部件。"""
    if not Path(docx_path).exists():
        return None
    with ZipFile(docx_path, "r") as archive:
        return _resolve_footnote_part_name_from_archive(archive)


def _resolve_configured_footnote_special_ids_from_archive(archive) -> set[str]:
    """从已打开 package 的 settings.xml 读取正数脚注分隔符 ID。"""
    main_document_uri = _related_part_uri_from_archive(
        archive,
        PackURI("/"),
        RT.OFFICE_DOCUMENT,
    )
    if main_document_uri is None:
        return set()
    settings_uri = _related_part_uri_from_archive(
        archive,
        main_document_uri,
        RT.SETTINGS,
    )
    if settings_uri is None:
        return set()
    settings_root = _parse_untrusted_ooxml_part(
        archive.read(str(settings_uri).lstrip("/"))
    )

    properties = settings_root.find(qn("w:footnotePr"))
    if properties is None:
        return set()
    special_ids = set()
    for reference in properties.findall(qn("w:footnote")):
        normalized_id = _normalize_ooxml_integer(reference.get(qn("w:id")))
        if normalized_id is None:
            raise ValueError("footnote settings contain an invalid special-note ID")
        special_ids.add(normalized_id)
    return special_ids


def _resolve_configured_footnote_special_ids(
    docx_path: str | Path,
) -> set[str]:
    """从 settings.xml 读取正数脚注分隔符 ID。"""
    if not Path(docx_path).exists():
        return set()
    with ZipFile(docx_path, "r") as archive:
        return _resolve_configured_footnote_special_ids_from_archive(archive)


def format_docx_footnotes(
    docx_path: str | Path,
    max_output_bytes=None,
    *,
    atomic_write: bool = True,
    raise_on_error: bool = False,
    limits: DocxValidationLimits = INPUT_DOCX_LIMITS,
) -> int:
    """
    统一脚注正文的字体与字号。

    `python-docx` 目前缺少稳定的脚注公开 API，因此这里在文档保存后
    直接修正脚注 relationship 实际指向部件中的 run 属性，尽量以最小
    改动覆盖真实论文里常见的“脚注字号/字体不统一”问题。

    外部路径默认使用输入文档预算；内部生成文档由调用方显式传入
    `GENERATED_DOCX_LIMITS`。校验和重写复用同一已打开的文件流。
    """
    configured_special_ids = set()

    def transform(xml_bytes: bytes):
        root = _parse_untrusted_ooxml_part(xml_bytes)
        formatted_count = 0
        changed = False

        for footnote in root.findall(qn("w:footnote")):
            footnote_type = footnote.get(qn("w:type"))
            footnote_id = footnote.get(qn("w:id"))
            if footnote_type in FOOTNOTE_SKIP_TYPES:
                continue

            if _is_negative_ooxml_decimal(footnote_id):
                continue

            normalized_id = _normalize_ooxml_integer(footnote_id)
            if (
                normalized_id in configured_special_ids
                or _note_item_has_special_content(footnote)
            ):
                continue

            footnote_changed = False
            for paragraph in footnote.findall("./" + qn("w:p")):
                footnote_changed |= _format_footnote_paragraph_xml(paragraph)
            for run in footnote.findall(".//" + qn("w:r")):
                footnote_changed |= _format_footnote_run_xml(run)

            if footnote_changed:
                formatted_count += 1
                changed = True

        xml_output = etree.tostring(
            root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )
        return xml_output, formatted_count, changed

    try:
        with Path(docx_path).open("rb") as source_stream:
            if not is_valid_docx_stream(source_stream, limits=limits):
                raise InvalidInputDocxError(
                    "document failed DOCX safety validation"
                )
            with ZipFile(source_stream, "r") as archive:
                configured_special_ids = (
                    _resolve_configured_footnote_special_ids_from_archive(
                        archive
                    )
                )
                footnote_part_name = (
                    _resolve_footnote_part_name_from_archive(archive)
                    or FOOTNOTE_XML_PATH
                )
            source_stream.seek(0)
            return _rewrite_docx_part(
                docx_path,
                footnote_part_name,
                transform,
                max_output_bytes=max_output_bytes,
                atomic_write=atomic_write,
                source_stream=source_stream,
                output_validation_limits=GENERATED_DOCX_LIMITS,
            )
    except OutputSizeLimitExceeded:
        raise
    except Exception as exc:
        logger.warning(f"统一脚注格式时出现警告：{format_log_exception(exc, docx_path)}")
        if raise_on_error:
            raise
        return 0


def _save_docx_with_footnote_postprocessing(
    save_callable,
    output_path: str | Path,
    max_output_bytes=None,
) -> int:
    """在同目录暂存完整 DOCX，脚注后处理成功后再一次性发布。"""
    output_file = Path(output_path)
    staging_path = _create_output_staging_path(output_file)
    try:
        _write_with_output_limit(save_callable, staging_path, max_output_bytes)
        _ensure_private_staging_file(staging_path)
        formatted_footnote_count = format_docx_footnotes(
            staging_path,
            max_output_bytes=max_output_bytes,
            atomic_write=False,
            raise_on_error=True,
            limits=GENERATED_DOCX_LIMITS,
        )
        # 脚注后处理会直接改写暂存包；在最终发布前再次验证，避免
        # 后处理器或兼容实现返回成功但留下结构损坏的 DOCX。
        if not is_valid_generated_docx(
            staging_path,
            limits=GENERATED_DOCX_LIMITS,
        ):
            raise InvalidGeneratedDocxError(
                "generated document failed output validation after footnote processing"
            )
        _ensure_private_staging_file(staging_path)
        staging_path.replace(output_file)
        return formatted_footnote_count
    finally:
        try:
            _cleanup_staging_path(staging_path)
        finally:
            release_active_temp_path(staging_path)


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


def _set_paragraph_outline_level(paragraph, level: int | None):
    """为段落设置目录级别，便于 Word TOC 字段收录。"""
    p_pr = paragraph._element.get_or_add_pPr()
    outline_level = p_pr.find(qn("w:outlineLvl"))

    if level is None:
        if outline_level is not None:
            p_pr.remove(outline_level)
        return

    if outline_level is None:
        outline_level = OxmlElement("w:outlineLvl")
    else:
        p_pr.remove(outline_level)
    p_pr._insert_outlineLvl(outline_level)

    outline_level.set(qn("w:val"), str(level))


def _set_paragraph_on_off_flag(paragraph, flag: str, enabled: bool):
    """设置段落级 on/off 标志，如 keepNext、keepLines。"""
    p_pr = paragraph._element.get_or_add_pPr()
    flag_element = p_pr.find(qn(f"w:{flag}"))

    if not enabled:
        if flag_element is not None:
            p_pr.remove(flag_element)
        return

    if flag_element is None:
        flag_element = OxmlElement(f"w:{flag}")
    else:
        p_pr.remove(flag_element)
    getattr(p_pr, f"_insert_{flag}")(flag_element)

    flag_element.set(qn("w:val"), "true")


def _set_paragraph_pagination_flags(
    paragraph,
    *,
    keep_next: bool | None = None,
    keep_lines: bool | None = None,
    widow_control: bool | None = None,
):
    """统一设置段落分页相关控制项。"""
    if keep_next is not None:
        _set_paragraph_on_off_flag(paragraph, "keepNext", keep_next)

    if keep_lines is not None:
        _set_paragraph_on_off_flag(paragraph, "keepLines", keep_lines)

    if widow_control is not None:
        _set_paragraph_on_off_flag(paragraph, "widowControl", widow_control)


def _clear_paragraph_style(paragraph, preserve_list_style: bool = False):
    """
    清除段落的已有样式设置，防止模板样式干扰排版。
    将段落样式重置为 Normal。
    """
    if preserve_list_style and _is_list_paragraph(paragraph):
        return

    try:
        part = paragraph.part
        document_part = getattr(part, "_document_part", part)
        normal_style_id = getattr(
            document_part,
            "_academic_normal_paragraph_style_id",
            None,
        )
        if normal_style_id:
            p_style = paragraph._element.get_or_add_pPr().get_or_add_pStyle()
            p_style.val = normal_style_id
        else:
            # 独立调用者可能没有经过文档初始化，保留原有回退路径。
            paragraph.style = "Normal"
    except (InvalidXmlError, OxmlInvalidXmlError, KeyError, ValueError):
        # 模板可能缺少 Normal，或把它错定义为字符样式。
        # 移除显式 pStyle 后由 Word 回退到包内可用的默认段落样式。
        p_pr = paragraph._element.pPr
        p_style = p_pr.find(qn("w:pStyle")) if p_pr is not None else None
        if p_style is not None:
            p_pr.remove(p_style)
    _clear_paragraph_numbering(paragraph)


def _get_primary_paragraph(container):
    """把页眉/页脚重置为单个主段落，并清除全部旧块级/行内内容。"""
    paragraphs = list(container.paragraphs)
    paragraph = paragraphs[0] if paragraphs else container.add_paragraph()

    container_element = container._element
    for child in list(container_element):
        if child is not paragraph._element:
            container_element.remove(child)

    paragraph_properties = paragraph._element.pPr
    for child in list(paragraph._element):
        if child is not paragraph_properties:
            paragraph._element.remove(child)

    # 该 story 的所有内容已被明确重置，旧图片、图表、OLE 及
    # 外部链接关系不再有合法引用。删除 Part 关系后由
    # python-docx 的关系图写出器自动裁掉已不可达的媒体部件。
    for relationship_id in list(container.part.rels):
        container.part.drop_rel(relationship_id)
    return paragraph


def _ensure_independent_first_page_part(
    section,
    *,
    reference_tag: str,
    expected_part_type,
    add_part,
) -> None:
    """确保首页页眉/页脚不与其它引用共享后再清空其内容。

    一些模板会显式写出 ``first`` 引用，却让它与 ``default`` 页眉/页脚
    复用同一个 relationship 或同一个目标 Part。直接通过首页代理修改 XML
    会因此连带修改正文运行页眉或页码。这里采用写时复制：只有引用缺失、
    类型异常或目标 Part 仍被其它 XML 引用时才新建部件；替换引用后，旧
    关系仅在确实没有其它引用时删除。
    """
    sect_pr = section._sectPr
    reference = next(
        (
            candidate
            for candidate in sect_pr.findall(qn(reference_tag))
            if candidate.get(qn("w:type")) == "first"
        ),
        None,
    )
    if reference is None:
        # 没有显式 first 定义时，python-docx 访问代理会创建独立部件。
        return

    document_part = section._document_part
    old_rid = reference.get(qn("r:id"))
    relationship = document_part.rels.get(old_rid) if old_rid else None
    shared = bool(old_rid) and document_part._rel_ref_count(old_rid) > 1
    valid_existing_part = (
        relationship is not None
        and not relationship.is_external
        and isinstance(relationship.target_part, expected_part_type)
    )
    shared_target_part = valid_existing_part and any(
        candidate_rid != old_rid
        and not candidate_relationship.is_external
        and candidate_relationship.target_part is relationship.target_part
        and document_part._rel_ref_count(candidate_rid) > 0
        for candidate_rid, candidate_relationship in document_part.rels.items()
    )
    shared = shared or shared_target_part
    if valid_existing_part and not shared:
        return

    _new_part, new_rid = add_part()
    reference.set(qn("r:id"), new_rid)

    if old_rid and document_part.rels.get(old_rid) is not None:
        if document_part._rel_ref_count(old_rid) == 0:
            document_part.drop_rel(old_rid)


def _get_primary_cell_paragraph(cell):
    """获取单元格中的主段落，并删除多余的默认空段落。"""
    paragraphs = list(cell.paragraphs)
    paragraph = paragraphs[0] if paragraphs else cell.add_paragraph()

    for extra in paragraphs[1:]:
        extra._element.getparent().remove(extra._element)

    _remove_all_runs(paragraph)
    return paragraph


def _allocate_package_partname(
    package,
    preferred_partname: PackURI,
    indexed_template: str,
) -> PackURI:
    """按 OPC 大小写不敏感规则分配全包唯一 Part 名。"""
    existing_names = [str(part.partname) for part in package.iter_parts()]
    used_names = {name.casefold() for name in existing_names}
    if len(used_names) != len(existing_names):
        raise ValueError("package has case-insensitive Part name collisions")
    if str(preferred_partname).casefold() not in used_names:
        return preferred_partname

    for index in range(1, len(existing_names) + 2):
        candidate = PackURI(indexed_template % index)
        if str(candidate).casefold() not in used_names:
            return candidate
    raise ValueError("package has no available Part name")


def _get_or_create_numbering_part(doc) -> NumberingPart:
    """返回文档的编号 Part；缺失时创建可由 python-docx 保存的最小 Part。

    ``python-docx 1.2`` 的 ``DocumentPart.numbering_part`` 在关系缺失时会
    调用尚未实现的 ``NumberingPart.new()``。numbering Part 本身在未使用
    原生编号的合法 DOCX 中是可选的，因此这里显式创建空的 ``w:numbering``
    根节点，后续再按需加入标题编号定义。
    """
    try:
        numbering_part = doc.part.part_related_by(RT.NUMBERING)
    except KeyError:
        package = doc.part.package
        if package is None:  # pragma: no cover - 已打开 Document 始终隶属 Package
            raise ValueError("document has no package")

        partname = _allocate_package_partname(
            package,
            PackURI("/word/numbering.xml"),
            "/word/numbering%d.xml",
        )

        numbering_part = NumberingPart(
            partname,
            CT.WML_NUMBERING,
            OxmlElement("w:numbering"),
            package,
        )
        doc.part.relate_to(numbering_part, RT.NUMBERING)

    if not isinstance(numbering_part, NumberingPart):
        raise TypeError("numbering relationship does not target a NumberingPart")
    return numbering_part


def _get_next_abstract_num_id(numbering_root) -> int:
    """返回 numbering.xml 中下一个可用的 abstractNumId。"""
    abstract_ids = set()
    for abstract_num in numbering_root.findall("./" + qn("w:abstractNum")):
        abstract_num_id = _parse_bounded_decimal(
            abstract_num.get(qn("w:abstractNumId")),
            MAX_OOXML_DECIMAL_NUMBER,
        )
        if abstract_num_id is not None:
            abstract_ids.add(abstract_num_id)

    candidate = 0
    while candidate in abstract_ids:
        candidate += 1
    if candidate > MAX_OOXML_DECIMAL_NUMBER:  # pragma: no cover - 受 XML 大小上限保护
        raise ValueError("numbering.xml has no available abstractNumId")
    return candidate


def _get_next_num_id(numbering_root, cache_owner=None) -> int:
    """返回 numbering.xml 中下一个可用的 numId，忽略非法属性并可按文档缓存。"""
    cache_key = id(numbering_root)
    cached_key = getattr(cache_owner, "_academic_numbering_root_key", None)
    num_ids = getattr(cache_owner, "_academic_used_num_ids", None)
    candidate = getattr(cache_owner, "_academic_next_num_id", None)

    if cached_key != cache_key or not isinstance(num_ids, set) or not isinstance(candidate, int):
        num_ids = set()
        for num in numbering_root.findall("./" + qn("w:num")):
            num_id = _parse_bounded_decimal(
                num.get(qn("w:numId")),
                MAX_OOXML_DECIMAL_NUMBER,
                minimum=1,
            )
            if num_id is not None:
                num_ids.add(num_id)

        candidate = 1
        while candidate in num_ids:
            candidate += 1

    if candidate > MAX_OOXML_DECIMAL_NUMBER:  # pragma: no cover - 受 XML 大小上限保护
        raise ValueError("numbering.xml has no available numId")

    num_ids.add(candidate)
    if cache_owner is not None:
        next_candidate = candidate + 1
        while next_candidate in num_ids:
            next_candidate += 1
        setattr(cache_owner, "_academic_numbering_root_key", cache_key)
        setattr(cache_owner, "_academic_used_num_ids", num_ids)
        setattr(cache_owner, "_academic_next_num_id", next_candidate)
    return candidate


def _is_compatible_heading_numbering_abstract(abstract_num) -> bool:
    """仅复用层级、格式与编号文本都符合本工具约定的定义。"""
    multi_level_type = abstract_num.find(qn("w:multiLevelType"))
    if multi_level_type is None or multi_level_type.get(qn("w:val")) != "multilevel":
        return False

    level_elements = abstract_num.findall(qn("w:lvl"))
    if len(level_elements) != len(HEADING_NUMBERING_LEVEL_TEXTS):
        return False

    levels = {}
    for level in level_elements:
        ilvl = _parse_bounded_decimal(
            level.get(qn("w:ilvl")),
            len(HEADING_NUMBERING_LEVEL_TEXTS) - 1,
        )
        if ilvl is None or ilvl in levels:
            return False
        levels[ilvl] = level

    if set(levels) != set(range(len(HEADING_NUMBERING_LEVEL_TEXTS))):
        return False

    for ilvl, expected_text in enumerate(HEADING_NUMBERING_LEVEL_TEXTS):
        level = levels[ilvl]
        start = level.find(qn("w:start"))
        num_fmt = level.find(qn("w:numFmt"))
        suffix = level.find(qn("w:suff"))
        level_text = level.find(qn("w:lvlText"))
        if (
            start is None
            or start.get(qn("w:val")) != "1"
            or num_fmt is None
            or num_fmt.get(qn("w:val")) != "decimal"
            or suffix is None
            or suffix.get(qn("w:val")) != "space"
            or level_text is None
            or level_text.get(qn("w:val")) != expected_text
        ):
            return False

    return True


def _get_or_create_heading_numbering_abstract_id(doc) -> int:
    """获取学术论文标题专用的多级编号定义。"""
    cached_id = getattr(doc, "_academic_heading_abstract_num_id", None)
    if cached_id is not None:
        return cached_id

    numbering_root = _get_or_create_numbering_part(doc).numbering_definitions._numbering

    for abstract_num in numbering_root.findall("./" + qn("w:abstractNum")):
        nsid = abstract_num.find(qn("w:nsid"))
        if nsid is not None and nsid.get(qn("w:val")) == HEADING_NUMBERING_NSID:
            abstract_num_id = _parse_bounded_decimal(
                abstract_num.get(qn("w:abstractNumId")),
                MAX_OOXML_DECIMAL_NUMBER,
            )
            if abstract_num_id is not None and _is_compatible_heading_numbering_abstract(abstract_num):
                setattr(doc, "_academic_heading_abstract_num_id", abstract_num_id)
                return abstract_num_id

    abstract_num_id = _get_next_abstract_num_id(numbering_root)
    abstract_num = OxmlElement("w:abstractNum")
    abstract_num.set(qn("w:abstractNumId"), str(abstract_num_id))

    nsid = OxmlElement("w:nsid")
    nsid.set(qn("w:val"), HEADING_NUMBERING_NSID)
    abstract_num.append(nsid)

    multi_level_type = OxmlElement("w:multiLevelType")
    multi_level_type.set(qn("w:val"), "multilevel")
    abstract_num.append(multi_level_type)

    template = OxmlElement("w:tmpl")
    template.set(qn("w:val"), HEADING_NUMBERING_TEMPLATE)
    abstract_num.append(template)

    for ilvl, level_text in enumerate(HEADING_NUMBERING_LEVEL_TEXTS):
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(ilvl))

        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        lvl.append(start)

        num_fmt = OxmlElement("w:numFmt")
        num_fmt.set(qn("w:val"), "decimal")
        lvl.append(num_fmt)

        suffix = OxmlElement("w:suff")
        suffix.set(qn("w:val"), "space")
        lvl.append(suffix)

        lvl_text = OxmlElement("w:lvlText")
        lvl_text.set(qn("w:val"), level_text)
        lvl.append(lvl_text)

        lvl_jc = OxmlElement("w:lvlJc")
        lvl_jc.set(qn("w:val"), "left")
        lvl.append(lvl_jc)

        abstract_num.append(lvl)

    children = list(numbering_root)
    insert_at = next(
        (
            index
            for index, child in enumerate(children)
            if child.tag in {qn("w:num"), qn("w:numIdMacAtCleanup")}
        ),
        len(children),
    )
    numbering_root.insert(insert_at, abstract_num)
    setattr(doc, "_academic_heading_abstract_num_id", abstract_num_id)
    return abstract_num_id


def _create_heading_numbering_instance(doc, numbering_parts: tuple[int, ...]) -> int | None:
    """为当前标题创建一个 concrete numbering 实例，并保留原始章节号。"""
    if not numbering_parts:
        return None

    abstract_num_id = _get_or_create_heading_numbering_abstract_id(doc)
    numbering_root = _get_or_create_numbering_part(doc).numbering_definitions._numbering
    num_id = _get_next_num_id(numbering_root, cache_owner=doc)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))

    abstract_num_id_element = OxmlElement("w:abstractNumId")
    abstract_num_id_element.set(qn("w:val"), str(abstract_num_id))
    num.append(abstract_num_id_element)

    for ilvl, start_value in enumerate(numbering_parts):
        lvl_override = OxmlElement("w:lvlOverride")
        lvl_override.set(qn("w:ilvl"), str(ilvl))

        start_override = OxmlElement("w:startOverride")
        start_override.set(qn("w:val"), str(start_value))
        lvl_override.append(start_override)
        num.append(lvl_override)

    # 使用 xmlchemy 的有序插入器，避免把 w:num 放到
    # w:numIdMacAtCleanup 之后；该路径不会扫描或解析既有 numId。
    numbering_root._insert_num(num)
    return num_id


def _apply_paragraph_numbering(paragraph, num_id: int, ilvl: int):
    """将 concrete numbering 绑定到段落。"""
    p_pr = paragraph._element.get_or_add_pPr()
    num_pr = p_pr.find(qn("w:numPr"))

    if num_pr is None:
        num_pr = OxmlElement("w:numPr")
    else:
        p_pr.remove(num_pr)
        for child in list(num_pr):
            num_pr.remove(child)
    p_pr._insert_numPr(num_pr)

    ilvl_element = OxmlElement("w:ilvl")
    ilvl_element.set(qn("w:val"), str(ilvl))
    num_pr.append(ilvl_element)

    num_id_element = OxmlElement("w:numId")
    num_id_element.set(qn("w:val"), str(num_id))
    num_pr.append(num_id_element)


def apply_native_heading_numbering(doc, paragraph, para_type: str, numbering_parts: tuple[int, ...]):
    """把识别出的章节号写成 Word 原生多级编号。"""
    level = HEADING_LEVEL_BY_TYPE.get(para_type)
    if level is None:
        return

    num_id = _create_heading_numbering_instance(doc, numbering_parts)
    if num_id is None:
        return

    _apply_paragraph_numbering(paragraph, num_id=num_id, ilvl=level)


def _set_xml_borders(border_container, border_map: dict[str, dict[str, str]]):
    """
    使用底层 XML 设置表格/单元格边框。

    python-docx 没有提供“只保留某一条边框”的高级接口，
    因此需要直接写 WordprocessingML 的边框节点属性。
    """
    border_order = (
        "top", "start", "left", "bottom", "end", "right",
        "insideH", "insideV", "between", "bar", "tl2br", "tr2bl",
    )
    border_positions = {edge: index for index, edge in enumerate(border_order)}

    for edge, attrs in border_map.items():
        edge_tag = qn(f"w:{edge}")
        border = border_container.find(edge_tag)
        if border is None:
            border = OxmlElement(f"w:{edge}")
        else:
            border_container.remove(border)

        edge_position = border_positions.get(edge, len(border_order))
        _insert_xml_child_before(
            border_container,
            border,
            tuple(f"w:{name}" for name in border_order[edge_position + 1:]),
        )

        for key, value in attrs.items():
            border.set(qn(f"w:{key}"), str(value))


def _hidden_border_attrs() -> dict[str, str]:
    """
    返回隐藏边框使用的属性。

    `none` 在 WPS 对表格/段落边框的兼容性通常比 `nil` 更稳定，
    更适合“隐藏大部分边框、只显示个别一条线”的场景。
    """
    return {
        "val": "none",
        "sz": "0",
        "space": "0",
        "color": "auto",
    }


def _hide_table_borders(table):
    """
    在表格级别隐藏所有边框。

    先把整张表的外框和内部横竖线全部置为 nil，
    后面再对“值列”的单元格单独开启 bottom 边框，
    这样就能得到类似“填写横线”的封面效果。
    """
    tbl_pr = table._tbl.tblPr
    tbl_borders = tbl_pr.find(qn("w:tblBorders"))
    if tbl_borders is None:
        tbl_borders = OxmlElement("w:tblBorders")
    else:
        tbl_pr.remove(tbl_borders)
    tbl_pr.insert_element_before(
        tbl_borders,
        "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
        "w:tblCaption", "w:tblDescription", "w:tblPrChange",
    )

    hidden = _hidden_border_attrs()
    _set_xml_borders(
        tbl_borders,
        {
            "top": hidden,
            "left": hidden,
            "bottom": hidden,
            "right": hidden,
            "insideH": hidden,
            "insideV": hidden,
        },
    )


def _hide_cell_borders(cell):
    """显式隐藏单元格四周边框，避免模板样式残留。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
    else:
        tc_pr.remove(tc_borders)
    tc_pr.insert_element_before(
        tc_borders,
        "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
        "w:vAlign", "w:hideMark", "w:headers", "w:cellIns", "w:cellDel",
        "w:cellMerge", "w:tcPrChange",
    )

    hidden = _hidden_border_attrs()
    _set_xml_borders(
        tc_borders,
        {
            "top": hidden,
            "left": hidden,
            "right": hidden,
            "bottom": hidden,
        },
    )


def _set_cell_only_bottom_border(cell, color: str = "000000", size: str = "8"):
    """
    只保留单元格的下边框，用来模拟封面信息栏的填写横线。

    实现步骤：
    1. 先把 top/left/right/bottom 全部清空为 nil
    2. 再把 bottom 单独改成 single

    这样可以确保无论模板原本是否带边框，最终都只有底部这一条线可见。
    """
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
    else:
        tc_pr.remove(tc_borders)
    tc_pr.insert_element_before(
        tc_borders,
        "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
        "w:vAlign", "w:hideMark", "w:headers", "w:cellIns", "w:cellDel",
        "w:cellMerge", "w:tcPrChange",
    )

    hidden = _hidden_border_attrs()
    _set_xml_borders(
        tc_borders,
        {
            "top": hidden,
            "left": hidden,
            "right": hidden,
            "bottom": hidden,
        },
    )
    _set_xml_borders(
        tc_borders,
        {
            "bottom": {
                "val": "single",
                "sz": size,
                "space": "0",
                "color": color,
            }
        },
    )


def _set_explicit_cell_borders(cell, top: bool = False, bottom: bool = False, color: str = "000000", size: str = "10"):
    """
    通过直接覆写单元格级别的四面边框，强行阻断并覆盖所有来自表格级别的样式继承。
    对于不需要显示的边，设置为 val="none"。
    这能保证 100% 出现需要的三线表线条，并消除表格自带格线。
    """
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is not None:
        tc_pr.remove(tc_borders)
        
    tc_borders = OxmlElement("w:tcBorders")
    tc_pr.insert_element_before(
        tc_borders,
        "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
        "w:vAlign", "w:hideMark", "w:headers", "w:cellIns", "w:cellDel",
        "w:cellMerge", "w:tcPrChange",
    )

    hidden = _hidden_border_attrs()
    visible = {
        "val": "single",
        "sz": size,
        "space": "0",
        "color": color,
    }
    
    _set_xml_borders(
        tc_borders,
        {
            "top": visible if top else hidden,
            "bottom": visible if bottom else hidden,
            "left": hidden,
            "right": hidden,
        },
    )

def _remove_table_borders(table):
    """移除表格级的边框定义，以防与单元格边框交织冲突"""
    tbl_pr = table._tbl.tblPr
    tbl_borders = tbl_pr.find(qn("w:tblBorders"))
    if tbl_borders is not None:
        tbl_pr.remove(tbl_borders)


def _set_row_repeat_as_header(row, enabled: bool = True):
    """将表格首行标记为跨页重复表头。"""
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = tr_pr.find(qn("w:tblHeader"))

    if not enabled:
        if tbl_header is not None:
            tr_pr.remove(tbl_header)
        return

    if tbl_header is None:
        tbl_header = OxmlElement("w:tblHeader")
    else:
        tr_pr.remove(tbl_header)
    tr_pr.insert_element_before(
        tbl_header,
        "w:tblCellSpacing", "w:jc", "w:hidden", "w:ins", "w:del", "w:trPrChange",
    )

    tbl_header.set(qn("w:val"), "true")


def _set_row_cant_split(row, enabled: bool = True):
    """禁止表格行在分页处被拆开，提升长表阅读连续性。"""
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = tr_pr.find(qn("w:cantSplit"))

    if not enabled:
        if cant_split is not None:
            tr_pr.remove(cant_split)
        return

    if cant_split is None:
        cant_split = OxmlElement("w:cantSplit")
    else:
        tr_pr.remove(cant_split)
    tr_pr.insert_element_before(
        cant_split,
        "w:trHeight", "w:tblHeader", "w:tblCellSpacing", "w:jc", "w:hidden",
        "w:ins", "w:del", "w:trPrChange",
    )

    cant_split.set(qn("w:val"), "true")


def format_three_line_table(table):
    """将普通表格处理为学术论文常见的三线表边框。"""
    rows = list(table.rows)
    if not rows:
        return

    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    _remove_table_borders(table)
    # 取消底层的表格样式以防底层隐藏线逻辑有奇怪的行为
    try:
        table.style = "Normal Table"
    except (InvalidXmlError, OxmlInvalidXmlError, KeyError, ValueError):
        tbl_style = table._tbl.tblPr.find(qn("w:tblStyle"))
        if tbl_style is not None:
            table._tbl.tblPr.remove(tbl_style)

    last_row_index = len(rows) - 1

    for row_index, row in enumerate(rows):
        is_header = (row_index == 0)
        is_last = (row_index == last_row_index)
        _set_row_repeat_as_header(row, enabled=is_header)
        _set_row_cant_split(row, enabled=True)
        
        for cell in iter_unique_row_cells(row):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

            for paragraph in cell.paragraphs:
                _set_paragraph_format(
                    paragraph,
                    alignment=WD_ALIGN_PARAGRAPH.CENTER if is_header else WD_ALIGN_PARAGRAPH.LEFT,
                    first_line_indent=Pt(0),
                    line_spacing=1.5,
                    line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
                )
                if is_header:
                    _apply_run_fonts(
                        paragraph,
                        cn_font="宋体",
                        en_font="Times New Roman",
                        size_pt=12,
                        bold=True,
                    )

            # 第一行：画顶线和底线（粗一点或者普通的单线条）
            # 最后一行：画底线
            # 其他行：全部 none
            _set_explicit_cell_borders(
                cell,
                top=is_header,
                bottom=is_header or is_last,
                size="10" if is_header or is_last else "0",
            )


def _set_paragraph_only_bottom_border(paragraph, color: str = "000000", size: str = "10"):
    """
    给段落仅保留一条下边框。

    WPS 对“单元格边框 + 整表隐藏边框”的组合并不总是稳定，
    所以这里在值列段落上再补一条段落下边框，作为更稳的可视化横线。
    """
    p_pr = paragraph._element.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
    else:
        p_pr.remove(p_bdr)
    p_pr.insert_element_before(
        p_bdr,
        "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
        "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
        "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
        "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
        "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
        "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
    )

    hidden = _hidden_border_attrs()
    _set_xml_borders(
        p_bdr,
        {
            "top": hidden,
            "left": hidden,
            "right": hidden,
            "bottom": {
                "val": "single",
                "sz": size,
                "space": "1",
                "color": color,
            },
        },
    )


def _prepend_block_elements(doc, elements):
    """
    将一组已创建好的段落/表格 XML 元素移动到文档最前方。

    python-docx 没有“在开头插入 block”的公开 API，
    所以这里采用“先追加、再搬到开头”的方式实现。
    """
    body = doc._body._element

    for insert_index, element in enumerate(elements):
        body.remove(element)
        body.insert(insert_index, element)


def _insert_block_elements_after(paragraph, elements):
    """将一组 block 元素插入到指定段落之后。"""
    parent = paragraph._element.getparent()
    insert_index = parent.index(paragraph._element) + 1

    for offset, element in enumerate(elements):
        element.getparent().remove(element)
        parent.insert(insert_index + offset, element)


def _resolve_cover_asset_path(explicit_path, asset_key: str) -> Path | None:
    """
    解析封面素材路径。

    优先使用显式传入的绝对/相对路径；如果没传，则自动在 `static/` 下
    查找约定好的默认文件名，减少调用方配置成本。
    """
    if explicit_path:
        candidate = Path(str(explicit_path)).expanduser()
        if candidate.exists():
            return candidate

    static_dir = Path(__file__).resolve().parent / "static"
    for filename in COVER_IMAGE_CANDIDATES.get(asset_key, ()):
        candidate = static_dir / filename
        if candidate.exists():
            return candidate

    return None


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


def _remove_invalid_header_footer_references(section) -> None:
    """移除缺失或类型错误的页眉/页脚关系，随后可安全重建。"""
    relationships = section._document_part.rels
    allowed_reference_types = {"default", "even", "first"}
    reference_specs = (
        ("w:headerReference", RT.HEADER, HeaderPart, qn("w:hdr")),
        ("w:footerReference", RT.FOOTER, FooterPart, qn("w:ftr")),
    )

    for tag, expected_reltype, expected_part_type, expected_root_tag in reference_specs:
        retained_types = set()
        for reference in list(section._sectPr.findall(qn(tag))):
            relationship_id = reference.get(qn("r:id"))
            reference_type = reference.get(qn("w:type"))
            relationship = relationships.get(relationship_id)
            if (
                reference_type not in allowed_reference_types
                or reference_type in retained_types
                or relationship is None
                or relationship.is_external
                or relationship.reltype != expected_reltype
                or not isinstance(relationship.target_part, expected_part_type)
                or relationship.target_part.element.tag != expected_root_tag
            ):
                section._sectPr.remove(reference)
                if (
                    relationship_id in relationships
                    and section._document_part._rel_ref_count(relationship_id) == 0
                ):
                    section._document_part.drop_rel(relationship_id)
                continue

            retained_types.add(reference_type)


def _link_default_header_footer_to_previous(section) -> None:
    """移除当前节的默认页眉/页脚引用，但不误删其他节共享的关系。"""
    document_part = section._document_part
    for tag in ("w:headerReference", "w:footerReference"):
        for reference in list(section._sectPr.findall(qn(tag))):
            if reference.get(qn("w:type")) != "default":
                continue
            relationship_id = reference.get(qn("r:id"))
            section._sectPr.remove(reference)
            if (
                relationship_id in document_part.rels
                and document_part._rel_ref_count(relationship_id) == 0
            ):
                document_part.drop_rel(relationship_id)


def _remove_header_footer_reference_types(section, reference_types) -> None:
    """删除指定类型的页眉/页脚引用，并回收无引用关系。"""
    document_part = section._document_part
    for tag in ("w:headerReference", "w:footerReference"):
        for reference in list(section._sectPr.findall(qn(tag))):
            if reference.get(qn("w:type")) not in reference_types:
                continue
            relationship_id = reference.get(qn("r:id"))
            section._sectPr.remove(reference)
            if (
                relationship_id in document_part.rels
                and document_part._rel_ref_count(relationship_id) == 0
            ):
                document_part.drop_rel(relationship_id)


def _paragraph_contains_page_break(paragraph) -> bool:
    """仅当段落最后一项可见/语义内容是显式分页符时返回 True。"""
    last_content_is_page_break = False
    text_content_tags = {
        qn("w:t"),
        qn("w:delText"),
        qn("w:instrText"),
    }
    non_text_content_tags = {
        qn("w:tab"),
        qn("w:ptab"),
        qn("w:cr"),
        qn("w:noBreakHyphen"),
        qn("w:softHyphen"),
        qn("w:sym"),
        qn("w:fldChar"),
        qn("w:footnoteReference"),
        qn("w:endnoteReference"),
        qn("w:commentReference"),
        qn("w:separator"),
        qn("w:continuationSeparator"),
        qn("w:pgNum"),
        qn("w:ruby"),
    }

    for node in paragraph._element.iter():
        tag = getattr(node, "tag", None)
        if tag == qn("w:br"):
            last_content_is_page_break = node.get(qn("w:type")) == "page"
        elif tag in text_content_tags:
            if node.text:
                last_content_is_page_break = False
        elif tag in non_text_content_tags or tag in DRAWING_XML_TAGS or tag in EQUATION_XML_TAGS:
            last_content_is_page_break = False

    return last_content_is_page_break


def _iter_document_body_blocks(doc):
    """按真实顺序遍历正文顶层块，排除包尾的 sectPr。"""
    for child in doc._element.body:
        if child.tag != qn("w:sectPr"):
            yield child


def _append_forced_page_break(doc):
    """在文档末尾追加一个显式分页符。"""
    page_break = doc.add_paragraph()
    _clear_paragraph_style(page_break)
    _set_paragraph_format(
        page_break,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Pt(0),
        line_spacing=1.0,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )
    page_break.add_run().add_break(WD_BREAK.PAGE)
    return page_break


def ensure_document_ends_with_page_break(doc):
    """确保当前文档末尾带有分页符，便于正文从下一页开始。"""
    blocks = list(_iter_document_body_blocks(doc))
    if blocks and blocks[-1].tag == qn("w:p"):
        paragraph_by_element = {
            paragraph._element: paragraph
            for paragraph in doc.paragraphs
        }
        last_paragraph = paragraph_by_element.get(blocks[-1])
        if last_paragraph is not None and _paragraph_contains_page_break(last_paragraph):
            return

    _append_forced_page_break(doc)


def get_max_printable_width(doc) -> int:
    """返回文档各节中可打印区域的最小宽度（EMU）。"""
    widths = []
    for section in doc.sections:
        printable_width = int(section.page_width) - int(section.left_margin) - int(section.right_margin)
        if printable_width > 0:
            widths.append(printable_width)

    return min(widths) if widths else 0


def _set_drawing_extent(container, extent, cx: int, cy: int) -> None:
    """更新绘图外框及实际变换尺寸，兼容缺失 graphic 子树的畸形图形。"""
    extent.set("cx", str(cx))
    extent.set("cy", str(cy))
    for drawing_extent in container.findall(".//" + qn("a:ext")):
        if drawing_extent.get("cx") is None or drawing_extent.get("cy") is None:
            continue
        drawing_extent.set("cx", str(cx))
        drawing_extent.set("cy", str(cy))


def constrain_inline_images(doc, progress_callback=None) -> int:
    """
    检测文档中的内嵌图片，若超出页面可打印宽度则按比例缩小。

    这里只处理 inline_shapes：
    - 对当前项目已经覆盖的“插入式图片”最有效
    - 与 python-docx 的能力边界一致，兼容性相对稳定
    """
    max_width = get_max_printable_width(doc)
    if max_width <= 0:
        return 0

    inline_shapes = list(doc.inline_shapes)
    resized_count = 0
    total_images = len(inline_shapes)

    for index, shape in enumerate(inline_shapes, start=1):
        emit_progress(
            progress_callback,
            3,
            f"正在检查第 {index}/{total_images} 张图片尺寸",
        )
        extent = shape._inline.find(qn("wp:extent"))
        if extent is None:
            continue
        original_width = _parse_bounded_decimal(extent.get("cx"), MAX_OOXML_COORDINATE, minimum=1)
        original_height = _parse_bounded_decimal(extent.get("cy"), MAX_OOXML_COORDINATE, minimum=1)
        if original_width is None or original_height is None or original_width <= max_width:
            continue

        scale = max_width / original_width
        new_height = max(1, int(round(original_height * scale)))
        _set_drawing_extent(shape._inline, extent, max_width, new_height)
        resized_count += 1
        emit_progress(
            progress_callback,
            3,
            f"已缩小第 {index}/{total_images} 张图片",
            "嵌入型图片宽度已限制到页面可打印区域内",
        )

    # 兼容各种由于直接粘贴导致的非标准图片（如浮动图形 wp:anchor 或旧版 v:shape）
    for anchor in doc._element.findall(".//" + qn("wp:anchor")):
        extent = anchor.find(".//" + qn("wp:extent"))
        if extent is not None:
            cx = _parse_bounded_decimal(extent.get("cx"), MAX_OOXML_COORDINATE, minimum=1)
            cy = _parse_bounded_decimal(extent.get("cy"), MAX_OOXML_COORDINATE, minimum=1)
            if cx is not None and cy is not None and cx > max_width:
                scale = max_width / cx
                new_cx = max_width
                new_cy = max(1, int(round(cy * scale)))
                _set_drawing_extent(anchor, extent, new_cx, new_cy)
                resized_count += 1

    # 兼容直接从部分网页/Excel 粘贴带来的 VML 图形 (v:shape)
    for v_shape in doc._element.findall(".//{urn:schemas-microsoft-com:vml}shape"):
        style_str = v_shape.get("style", "")
        if "width:" in style_str and "height:" in style_str:
            w_match = re.search(r"width:([0-9.]+)([a-zA-Z]+)", style_str)
            h_match = re.search(r"height:([0-9.]+)([a-zA-Z]+)", style_str)
            if w_match and h_match:
                w_val = _parse_vml_dimension(w_match.group(1))
                h_val = _parse_vml_dimension(h_match.group(1))
                w_unit = w_match.group(2).lower()
                h_unit = h_match.group(2).lower()
                unit_to_emu = {"pt": 12700, "in": 914400, "cm": 360000, "mm": 36000, "px": 9525}
                if (
                    w_val is None
                    or h_val is None
                    or w_unit not in unit_to_emu
                    or h_unit not in unit_to_emu
                    or w_val > MAX_OOXML_COORDINATE / unit_to_emu[w_unit]
                    or h_val > MAX_OOXML_COORDINATE / unit_to_emu[h_unit]
                ):
                    continue
                w_emu = int(w_val * unit_to_emu[w_unit])
                if w_emu > max_width:
                    scale = max_width / w_emu
                    new_style = re.sub(r"width:[0-9.]+[a-zA-Z]+", f"width:{w_val * scale:.2f}{w_unit}", style_str)
                    new_style = re.sub(r"height:[0-9.]+[a-zA-Z]+", f"height:{h_val * scale:.2f}{h_unit}", new_style)
                    v_shape.set("style", new_style)
                    resized_count += 1

    return resized_count


def _append_toc_field(paragraph):
    """插入 Word 目录域，打开文档更新域后即可生成目录。"""
    size_pt = 12

    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    run_begin = paragraph.add_run()
    _set_run_font(run_begin, cn_font="宋体", en_font="Times New Roman", size_pt=size_pt)
    run_begin._r.append(fld_char_begin)

    instr_text = OxmlElement("w:instrText")
    instr_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr_text.text = 'TOC \\o "1-3" \\h \\z \\u'
    run_instr = paragraph.add_run()
    _set_run_font(run_instr, cn_font="宋体", en_font="Times New Roman", size_pt=size_pt)
    run_instr._r.append(instr_text)

    fld_char_separate = OxmlElement("w:fldChar")
    fld_char_separate.set(qn("w:fldCharType"), "separate")
    run_separate = paragraph.add_run()
    _set_run_font(run_separate, cn_font="宋体", en_font="Times New Roman", size_pt=size_pt)
    run_separate._r.append(fld_char_separate)

    placeholder = paragraph.add_run("目录将在打开文档后自动生成，可右键更新目录立即刷新。")
    _set_run_font(placeholder, cn_font="宋体", en_font="Times New Roman", size_pt=size_pt)

    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")
    run_end = paragraph.add_run()
    _set_run_font(run_end, cn_font="宋体", en_font="Times New Roman", size_pt=size_pt)
    run_end._r.append(fld_char_end)


def _enable_field_updates_on_open(doc):
    """提示 Word/WPS 在打开文档时更新域代码。"""
    settings_element = doc.settings.element
    existing_update_fields = list(settings_element.findall(qn("w:updateFields")))
    update_fields = (
        existing_update_fields[0]
        if existing_update_fields
        else OxmlElement("w:updateFields")
    )
    for existing in existing_update_fields:
        settings_element.remove(existing)
    settings_element.insert_element_before(
        update_fields,
        "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat",
        "w:docVars", "w:rsids", "m:mathPr", "w:attachedSchema", "w:themeFontLang",
        "w:clrSchemeMapping", "w:doNotIncludeSubdocsInStats",
        "w:doNotAutoCompressPictures", "w:forceUpgrade", "w:captions",
        "w:readModeInkLockDown", "w:smartTagType", "sl:schemaLibrary",
        "w:shapeDefaults", "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
    )

    update_fields.set(qn("w:val"), "true")


def apply_document_layout(doc, header_text: str) -> dict:
    """统一设置纸张、页边距、页眉和页码。"""
    raw_header = (header_text or "").strip()
    if not raw_header:
        header_value = DEFAULT_HEADER_TEXT
    elif len(raw_header) <= RUNNING_HEADER_MAX_LENGTH:
        header_value = raw_header
    else:
        header_value = f"{raw_header[:RUNNING_HEADER_MAX_LENGTH - 3].rstrip()}..."

    # 排版器生成一套所有页通用的运行页眉/页码。若模板开启
    # 奇偶页或首页分离，仅重建 default Part 会让旧 Part 继续显示。
    # 新封面会在后续 generate_cover_page 中仅为第一节恢复首页空白模式。
    doc.settings.odd_and_even_pages_header_footer = False

    for section_index, section in enumerate(doc.sections):
        _remove_invalid_header_footer_references(section)
        _remove_header_footer_reference_types(section, {"even", "first"})
        section.different_first_page_header_footer = False
        section.page_width = Cm(PAGE_LAYOUT["page_width_cm"])
        section.page_height = Cm(PAGE_LAYOUT["page_height_cm"])
        section.top_margin = Cm(PAGE_LAYOUT["margins_cm"]["top"])
        section.bottom_margin = Cm(PAGE_LAYOUT["margins_cm"]["bottom"])
        section.left_margin = Cm(PAGE_LAYOUT["margins_cm"]["left"])
        section.right_margin = Cm(PAGE_LAYOUT["margins_cm"]["right"])
        section.header_distance = Cm(PAGE_LAYOUT["header_distance_cm"])
        section.footer_distance = Cm(PAGE_LAYOUT["footer_distance_cm"])
        if section_index > 0:
            # 所有节使用同一套运行页眉与页码，避免为每个分节
            # 新建两个部件并在大量分节时引发 O(S²) 包扫描。
            _link_default_header_footer_to_previous(section)
            continue

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
        if header_value:
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


def generate_cover_page(doc, info_dict):
    """
    在文档最前方插入一页标准课程论文封面，并在封面末尾补分页符。

    设计为“正文排版完成后再调用”，这样封面标题不会被正文识别逻辑再次改写。
    如果缺少 title，则视为没有封面信息，直接返回 False。
    """
    if doc is None:
        raise ValueError("doc 不能为空")

    info_dict = info_dict or {}
    title_text = normalize_text_for_matching(str(info_dict.get("title", "")))
    cover_title = normalize_text_for_matching(
        str(info_dict.get("cover_title") or info_dict.get("course_title") or title_text)
    )
    if not title_text:
        return False

    new_blocks = []

    school_name = normalize_text_for_matching(str(info_dict.get("school_name", "浙江工商大学")))
    logo_file = _resolve_cover_asset_path(info_dict.get("logo_path"), "logo")
    school_name_image_file = _resolve_cover_asset_path(info_dict.get("school_name_image_path"), "school_name")

    # ---------- 1. 顶部校徽 ----------
    logo_paragraph = doc.add_paragraph()
    _clear_paragraph_style(logo_paragraph)
    _set_paragraph_format(
        logo_paragraph,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),
        space_before=Pt(COVER_LAYOUT["logo_space_before_pt"]),
        space_after=Pt(COVER_LAYOUT["logo_space_after_pt"]),
        line_spacing=1.0,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )
    if logo_file is not None:
        logo_paragraph.add_run().add_picture(str(logo_file), width=Cm(COVER_LAYOUT["logo_width_cm"]))
    new_blocks.append(logo_paragraph._element)

    # ---------- 2. 校名 ----------
    school_paragraph = doc.add_paragraph()
    _clear_paragraph_style(school_paragraph)
    _set_paragraph_format(
        school_paragraph,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),
        space_after=Pt(COVER_LAYOUT["school_name_space_after_pt"]),
        line_spacing=1.0,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )
    if school_name_image_file is not None:
        school_paragraph.add_run().add_picture(str(school_name_image_file), width=Cm(COVER_LAYOUT["school_name_width_cm"]))
    else:
        school_run = school_paragraph.add_run(school_name)
        _set_run_font(school_run, cn_font="华文行楷", en_font="Times New Roman", size_pt=28, bold=False)
    new_blocks.append(school_paragraph._element)

    # ---------- 3. 封面大标题 ----------
    title_paragraph = doc.add_paragraph()
    _clear_paragraph_style(title_paragraph)
    _set_paragraph_format(
        title_paragraph,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),
        space_after=Pt(COVER_LAYOUT["title_space_after_pt"]),
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )
    title_run = title_paragraph.add_run(cover_title)
    _set_run_font(
        title_run,
        cn_font="宋体",
        en_font="Times New Roman",
        size_pt=COVER_LAYOUT["title_size_pt"],
        bold=True,
    )
    new_blocks.append(title_paragraph._element)

    info_spacer = doc.add_paragraph()
    _clear_paragraph_style(info_spacer)
    _set_paragraph_format(
        info_spacer,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),
        space_after=Pt(COVER_LAYOUT["info_spacer_after_pt"]),
        line_spacing=1.0,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )
    new_blocks.append(info_spacer._element)

    # ---------- 4. 个人信息栏 ----------
    # 绝不使用空格/下划线硬凑对齐，而是借助 5x2 表格完成稳定布局。
    table = doc.add_table(rows=len(COVER_INFO_FIELDS), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _hide_table_borders(table)

    for row_index, (label, key) in enumerate(COVER_INFO_FIELDS):
        label_cell = table.cell(row_index, 0)
        value_cell = table.cell(row_index, 1)

        label_cell.width = Cm(COVER_LAYOUT["label_width_cm"])
        value_cell.width = Cm(COVER_LAYOUT["value_width_cm"])
        label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        value_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        label_paragraph = _get_primary_cell_paragraph(label_cell)
        _clear_paragraph_style(label_paragraph)
        _set_paragraph_format(
            label_paragraph,
            alignment=WD_ALIGN_PARAGRAPH.RIGHT,
            first_line_indent=Pt(0),
            space_before=Pt(6),
            space_after=Pt(6),
            line_spacing=1.25,
            line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
        )
        label_run = label_paragraph.add_run(label)
        _set_run_font(
            label_run,
            cn_font="宋体",
            en_font="Times New Roman",
            size_pt=COVER_LAYOUT["info_font_pt"],
            bold=True,
        )

        value_paragraph = _get_primary_cell_paragraph(value_cell)
        _clear_paragraph_style(value_paragraph)
        _set_paragraph_format(
            value_paragraph,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            first_line_indent=Pt(0),
            space_before=Pt(6),
            space_after=Pt(6),
            line_spacing=1.25,
            line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
        )
        value_text = normalize_text_for_matching(str(info_dict.get(key, "")))
        value_run = value_paragraph.add_run(value_text)
        # 统一设置中西文字体对，中文保持宋体，数字学号自动显示为 Times New Roman。
        _set_run_font(
            value_run,
            cn_font="宋体",
            en_font="Times New Roman",
            size_pt=COVER_LAYOUT["info_font_pt"],
            bold=False,
        )

        _hide_cell_borders(label_cell)
        _set_cell_only_bottom_border(value_cell, size="10")
        _set_paragraph_only_bottom_border(value_paragraph, size="10")

    new_blocks.append(table._element)

    # ---------- 5. 封面结束后分页 ----------
    page_break = doc.add_paragraph()
    _clear_paragraph_style(page_break)
    _set_paragraph_format(
        page_break,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Pt(0),
        line_spacing=1.0,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )
    page_break.add_run().add_break(WD_BREAK.PAGE)
    new_blocks.append(page_break._element)

    _prepend_block_elements(doc, new_blocks)

    # 使用“首页不同”模式，让封面页不显示页眉和页码，正文从第二页开始承接常规页眉页码。
    first_section = doc.sections[0]
    first_section.different_first_page_header_footer = True

    _ensure_independent_first_page_part(
        first_section,
        reference_tag="w:headerReference",
        expected_part_type=HeaderPart,
        add_part=first_section._document_part.add_header_part,
    )
    first_page_header = _get_primary_paragraph(first_section.first_page_header)
    _clear_paragraph_style(first_page_header)
    _set_paragraph_format(
        first_page_header,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),
        line_spacing=1.0,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )

    _ensure_independent_first_page_part(
        first_section,
        reference_tag="w:footerReference",
        expected_part_type=FooterPart,
        add_part=first_section._document_part.add_footer_part,
    )
    first_page_footer = _get_primary_paragraph(first_section.first_page_footer)
    _clear_paragraph_style(first_page_footer)
    _set_paragraph_format(
        first_page_footer,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),
        line_spacing=1.0,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )

    return True


def prepare_cover_info(cover_info, detected_title: str) -> dict | None:
    """整理自动封面所需信息，并为缺省标题补齐回退值。"""
    if not isinstance(cover_info, dict):
        return None

    resolved = {}
    for key in (
        "title",
        "cover_title",
        "course_title",
        "college",
        "teacher",
        "class_name",
        "student_name",
        "student_id",
        "school_name",
        "logo_path",
        "school_name_image_path",
    ):
        value = cover_info.get(key)
        if value is None:
            continue
        resolved[key] = str(value).strip() if isinstance(value, str) else value

    fallback_title = normalize_text_for_matching(
        str(
            resolved.get("title")
            or detected_title
            or resolved.get("cover_title")
            or resolved.get("course_title")
            or ""
        )
    )
    if not fallback_title:
        return None

    resolved["title"] = fallback_title
    return resolved


def insert_table_of_contents(doc, title_index: int | None, heading_count: int) -> bool:
    """在标题后插入目录标题、TOC 域和分页符。"""
    if title_index is None or heading_count < MIN_TOC_HEADING_COUNT:
        return False

    title_paragraph = doc.paragraphs[title_index]
    new_blocks = []

    toc_heading = doc.add_paragraph()
    _clear_paragraph_style(toc_heading)
    _replace_paragraph_text(toc_heading, "目录")
    _set_paragraph_format(
        toc_heading,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),
        space_before=Pt(16),
        space_after=Pt(12),
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )
    _set_paragraph_outline_level(toc_heading, None)
    _apply_run_fonts(toc_heading, cn_font="黑体", en_font="Times New Roman", size_pt=16, bold=True)
    new_blocks.append(toc_heading._element)

    toc_paragraph = doc.add_paragraph()
    _clear_paragraph_style(toc_paragraph)
    _set_paragraph_format(
        toc_paragraph,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Pt(0),
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )
    _set_paragraph_outline_level(toc_paragraph, None)
    _append_toc_field(toc_paragraph)
    new_blocks.append(toc_paragraph._element)

    page_break = doc.add_paragraph()
    _clear_paragraph_style(page_break)
    _set_paragraph_format(
        page_break,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Pt(0),
        line_spacing=1.0,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )
    _set_paragraph_outline_level(page_break, None)
    page_break.add_run().add_break(WD_BREAK.PAGE)
    new_blocks.append(page_break._element)

    _insert_block_elements_after(title_paragraph, new_blocks)
    _enable_field_updates_on_open(doc)
    return True


# ============================================================
# 各类型段落的格式化函数
# ============================================================
def format_body(
    paragraph,
    in_table: bool = False,
    *,
    normalized_text: str | None = None,
    has_equation: bool | None = None,
    has_drawing: bool | None = None,
):
    """
    正文格式：
      - 中文字体：宋体
      - 西文字体：Times New Roman
      - 字号：小四（12pt）
      - 首行缩进：2 个中文字符（约 0.74cm × 2 ≈ 对于小四号字约 24pt）
      - 行距：1.5 倍行距
    """
    if normalized_text is None:
        normalized_text = normalize_text_for_matching(paragraph.text)
    _set_paragraph_outline_level(paragraph, None)
    if has_equation is None:
        has_equation = _has_equation_content(paragraph)

    if has_equation and not normalized_text:
        _clear_paragraph_style(paragraph)
        _set_paragraph_format(
            paragraph,
            alignment=WD_ALIGN_PARAGRAPH.LEFT if in_table else WD_ALIGN_PARAGRAPH.CENTER,
            first_line_indent=Pt(0),
            line_spacing=1.5,
            line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
        )
        _set_paragraph_pagination_flags(
            paragraph,
            keep_next=False,
            keep_lines=True,
            widow_control=True,
        )
        _apply_run_fonts(paragraph, cn_font="宋体", en_font="Times New Roman", size_pt=12)
        return

    if has_drawing is None:
        has_drawing = _has_drawing(paragraph)

    if has_drawing and not normalized_text:
        try:
            drawing_alignment = paragraph.paragraph_format.alignment
        except (InvalidXmlError, OxmlInvalidXmlError, ValueError):
            drawing_alignment = None
        _set_paragraph_format(
            paragraph,
            alignment=drawing_alignment or WD_ALIGN_PARAGRAPH.CENTER,
            first_line_indent=Pt(0),
        )
        _set_paragraph_pagination_flags(
            paragraph,
            keep_next=not in_table,
            keep_lines=True,
            widow_control=True,
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

    _set_paragraph_pagination_flags(paragraph, widow_control=True)
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
    _set_paragraph_outline_level(paragraph, None)
    _set_paragraph_pagination_flags(paragraph, keep_next=True, keep_lines=True)

    _apply_run_fonts(paragraph, cn_font="黑体", en_font="Times New Roman", size_pt=18, bold=True)


def format_heading_l1(paragraph, text_override: str | None = None, outline_level: int | None = 0):
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
    _set_paragraph_outline_level(paragraph, outline_level)
    _set_paragraph_pagination_flags(paragraph, keep_next=True, keep_lines=True)

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
    _set_paragraph_outline_level(paragraph, 1)
    _set_paragraph_pagination_flags(paragraph, keep_next=True, keep_lines=True)

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
    _set_paragraph_outline_level(paragraph, 2)
    _set_paragraph_pagination_flags(paragraph, keep_next=True, keep_lines=True)

    _apply_run_fonts(paragraph, cn_font="宋体", en_font="Times New Roman", size_pt=12, bold=True)


def format_figure_table(paragraph, text_override: str | None = None, *, keep_next: bool = False):
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
    _set_paragraph_outline_level(paragraph, None)
    _set_paragraph_pagination_flags(paragraph, keep_next=keep_next, keep_lines=True)

    _apply_run_fonts(paragraph, cn_font="黑体", en_font="Times New Roman", size_pt=10.5)


def format_references_heading(paragraph, text_override: str | None = None):
    """
    参考文献标题格式：
      - 居中
      - 宋体加粗
      - 比一级标题更克制，贴近参考样文的紧凑风格
      - 自动补全为“参考文献：”
    """
    _clear_paragraph_style(paragraph)
    if text_override is not None:
        heading_text = normalize_text_for_matching(text_override)
        if heading_text:
            _replace_paragraph_text(paragraph, "参考文献：")

    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),
        space_before=Pt(12),
        space_after=Pt(12),
        line_spacing=1.0,
        line_spacing_rule=WD_LINE_SPACING.SINGLE,
    )
    _set_paragraph_outline_level(paragraph, None)
    _set_paragraph_pagination_flags(paragraph, keep_next=True, keep_lines=True)

    _apply_run_fonts(paragraph, cn_font="宋体", en_font="Times New Roman", size_pt=12, bold=True)


def _format_labeled_paragraph(
    paragraph,
    label_pattern: re.Pattern,
    *,
    cn_font: str,
    en_font: str,
    size_pt: float,
    alignment,
    first_line_indent,
    label_text_override: str | None = None,
    preserve_runs: bool = False,
):
    """按“标签 + 正文”结构重建段落，并统一字体与加粗规则。"""
    _clear_paragraph_style(paragraph)
    _set_paragraph_outline_level(paragraph, None)
    _set_paragraph_format(
        paragraph,
        alignment=alignment,
        first_line_indent=first_line_indent,
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )
    _set_paragraph_pagination_flags(paragraph, widow_control=True)

    if preserve_runs:
        _apply_run_fonts(paragraph, cn_font=cn_font, en_font=en_font, size_pt=size_pt)
        return

    full_text = paragraph.text
    match = label_pattern.match(full_text)

    if match:
        label_end = match.end()
        label_text = label_text_override if label_text_override is not None else full_text[:label_end]
        body_text = full_text[label_end:]

        _remove_all_runs(paragraph)

        run_label = paragraph.add_run(label_text)
        _set_run_font(run_label, cn_font=cn_font, en_font=en_font, size_pt=size_pt, bold=True)

        if body_text:
            run_body = paragraph.add_run(body_text)
            _set_run_font(run_body, cn_font=cn_font, en_font=en_font, size_pt=size_pt, bold=False)
        return

    _apply_run_fonts(paragraph, cn_font=cn_font, en_font=en_font, size_pt=size_pt)


def format_abstract_or_keywords(
    paragraph,
    label_pattern: re.Pattern,
    *,
    preserve_runs: bool = False,
):
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
    _format_labeled_paragraph(
        paragraph,
        label_pattern,
        cn_font="宋体",
        en_font="Times New Roman",
        size_pt=12,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_indent=Pt(24),
        preserve_runs=preserve_runs,
    )

def format_english_abstract_heading(paragraph, *, preserve_runs: bool = False):
    """英文摘要标题格式：居中、Times New Roman、12pt、加粗。"""
    _clear_paragraph_style(paragraph)
    if not preserve_runs:
        _replace_paragraph_text(paragraph, "Abstract")
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Pt(0),
        space_before=Pt(12),
        space_after=Pt(12),
        line_spacing=1.0,
        line_spacing_rule=WD_LINE_SPACING.SINGLE,
    )
    _set_paragraph_outline_level(paragraph, None)
    _set_paragraph_pagination_flags(paragraph, widow_control=True)
    _apply_run_fonts(paragraph, cn_font="宋体", en_font="Times New Roman", size_pt=12, bold=True)


def format_english_abstract(
    paragraph,
    label_pattern: re.Pattern | None = None,
    *,
    preserve_runs: bool = False,
):
    """英文摘要正文：Times New Roman、12pt、1.5 倍行距，不额外首行缩进。"""
    if label_pattern is not None:
        _format_labeled_paragraph(
            paragraph,
            label_pattern,
            cn_font="宋体",
            en_font="Times New Roman",
            size_pt=12,
            alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
            first_line_indent=Pt(0),
            label_text_override="Abstract: ",
            preserve_runs=preserve_runs,
        )
        return

    _clear_paragraph_style(paragraph)
    _set_paragraph_outline_level(paragraph, None)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_indent=Pt(0),
        line_spacing=1.5,
        line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    )
    _set_paragraph_pagination_flags(paragraph, widow_control=True)
    _apply_run_fonts(paragraph, cn_font="宋体", en_font="Times New Roman", size_pt=12)


def format_english_keywords(paragraph, *, preserve_runs: bool = False):
    """英文关键词：标签加粗、正文常规，整体左对齐。"""
    _format_labeled_paragraph(
        paragraph,
        RE_ENGLISH_KEYWORDS,
        cn_font="宋体",
        en_font="Times New Roman",
        size_pt=12,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Pt(0),
        label_text_override="Keywords: ",
        preserve_runs=preserve_runs,
    )


def format_caption_note(paragraph, *, preserve_runs: bool = False):
    """
    图表附注/来源格式：
      - 左对齐
      - 宋体 / Times New Roman，五号（10.5pt）
      - 单倍行距
      - 标签部分加粗，正文常规
    """
    _clear_paragraph_style(paragraph)
    _set_paragraph_outline_level(paragraph, None)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Pt(0),
        space_before=Pt(0),
        space_after=Pt(0),
        line_spacing=1.0,
        line_spacing_rule=WD_LINE_SPACING.SINGLE,
    )
    _set_paragraph_pagination_flags(paragraph, widow_control=True)

    if preserve_runs:
        _apply_run_fonts(paragraph, cn_font="宋体", en_font="Times New Roman", size_pt=10.5)
        return

    full_text = paragraph.text
    match = RE_CAPTION_NOTE.match(full_text)
    if match:
        label_text = full_text[:match.end()]
        body_text = full_text[match.end():]

        _remove_all_runs(paragraph)

        run_label = paragraph.add_run(label_text)
        _set_run_font(run_label, cn_font="宋体", en_font="Times New Roman", size_pt=10.5, bold=True)

        if body_text:
            run_body = paragraph.add_run(body_text)
            _set_run_font(run_body, cn_font="宋体", en_font="Times New Roman", size_pt=10.5, bold=False)
        return

    _apply_run_fonts(paragraph, cn_font="宋体", en_font="Times New Roman", size_pt=10.5)


def format_reference_entry(paragraph):
    """
    参考文献条目格式：
      - 左对齐
      - 宋体 / Times New Roman，五号（10pt）
      - 单倍行距
      - 参考样文使用轻微首行缩进，而不是正文式悬挂缩进
    """
    _clear_paragraph_style(paragraph)
    _set_paragraph_format(
        paragraph,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Pt(21),
        space_before=Pt(0),
        space_after=Pt(0),
        line_spacing=1.0,
        line_spacing_rule=WD_LINE_SPACING.SINGLE,
    )
    _set_paragraph_outline_level(paragraph, None)
    _set_paragraph_pagination_flags(paragraph, widow_control=True)
    paragraph.paragraph_format.left_indent = Pt(0)
    paragraph.paragraph_format.right_indent = Pt(0)

    _apply_run_fonts(paragraph, cn_font="宋体", en_font="Times New Roman", size_pt=10)


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
def _run_document_processing(
    doc,
    output_path: str,
    *,
    progress_callback=None,
    cover_info=None,
    max_output_bytes=None,
) -> dict | bool:
    """执行文档处理并将模板/部件异常转换为稳定的 False。

    上层 Web 路径会再根据 False 返回用户可理解的错误；保留
    OutputSizeLimitExceeded 向上传递，以便调用方正确返回 413。
    """
    try:
        return _process_document(
            doc,
            output_path,
            progress_callback=progress_callback,
            cover_info=cover_info,
            max_output_bytes=max_output_bytes,
        )
    except OutputSizeLimitExceeded:
        raise
    except Exception as exc:
        logger.error(
            f"文档处理失败：{format_log_exception(exc, output_path)}"
        )
        return False


def format_academic_paper(
    input_path: str,
    output_path: str,
    progress_callback=None,
    cover_info=None,
    max_output_bytes=None,
) -> dict | bool:
    """
    读取未排版的 .docx 文档，根据学术论文排版规则进行格式重构，保存为新文档。

    Args:
        input_path:  输入文档路径（.docx）
        output_path: 输出文档路径（.docx）
        max_output_bytes: 可选的输出 ZIP 硬上限，写入过程中立即执行

    Returns:
        True 表示排版成功，False 表示处理失败
    """
    # ---------- 1. 文件校验与读取 ----------
    input_file = Path(input_path)

    if not input_file.exists():
        logger.error(f"输入文件不存在：{format_log_path(input_path)}")
        return False

    if not input_file.suffix.lower() == ".docx":
        logger.error(f"不支持的文件格式（仅支持 .docx）：{input_file.suffix}")
        return False

    try:
        doc = _load_validated_document(input_path)
        logger.info(f"成功读取文档：{format_log_path(input_path)}（共 {len(doc.paragraphs)} 个段落）")
        emit_progress(
            progress_callback,
            1,
            "文档读取完成，正在解析结构",
            f"共 {len(doc.paragraphs)} 个段落",
        )
    except Exception as e:
        logger.error(f"无法读取文档 {format_log_path(input_path)}：{format_log_exception(e, input_path)}")
        return False

    return _run_document_processing(
        doc,
        output_path,
        progress_callback=progress_callback,
        cover_info=cover_info,
        max_output_bytes=max_output_bytes,
    )


def format_academic_paper_from_text(
    text: str,
    output_path: str,
    progress_callback=None,
    cover_info=None,
    max_output_bytes=None,
) -> dict | bool:
    """
    将纯文本内容转换为符合学术论文排版规则的 .docx 文档。

    Args:
        text:        输入的纯文本内容（多行）
        output_path: 输出文档路径（.docx）
        max_output_bytes: 可选的输出 ZIP 硬上限，写入过程中立即执行

    Returns:
        True 表示排版成功，False 表示处理失败
    """
    try:
        lines = split_text_to_paragraphs(text)
        doc = Document()
        for line in lines:
            doc.add_paragraph(line)
        logger.info(f"成功从文本创建文档（共 {len(doc.paragraphs)} 个段落）")
        emit_progress(
            progress_callback,
            1,
            "文本读取完成，正在生成文档结构",
            f"共 {len(doc.paragraphs)} 个段落",
        )
    except Exception as e:
        logger.error(f"无法从文本创建文档：{format_log_exception(e, output_path)}")
        return False

    return _run_document_processing(
        doc,
        output_path,
        progress_callback=progress_callback,
        cover_info=cover_info,
        max_output_bytes=max_output_bytes,
    )


def _process_document(
    doc,
    output_path: str,
    progress_callback=None,
    cover_info=None,
    max_output_bytes=None,
) -> dict | bool:
    """内部处理逻辑，将 Document 对象排版并保存。"""
    # ---------- 2. 设置默认文档级字体 ----------
    emit_progress(progress_callback, 1, "正在初始化页面设置与默认样式")
    try:
        style = _ensure_normal_paragraph_style(doc)
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
        logger.warning(f"设置默认样式时出现警告：{format_log_exception(e, output_path)}")

    # ---------- 3. 统计信息 ----------
    stats = {
        ParagraphType.TITLE: 0,
        ParagraphType.HEADING_L1: 0,
        ParagraphType.HEADING_L2: 0,
        ParagraphType.HEADING_L3: 0,
        ParagraphType.FIGURE_CAPTION: 0,
        ParagraphType.TABLE_CAPTION: 0,
        ParagraphType.CAPTION_NOTE: 0,
        ParagraphType.SECTION_HEADING: 0,
        ParagraphType.REFERENCES_HEADING: 0,
        ParagraphType.REFERENCE_ENTRY: 0,
        ParagraphType.ABSTRACT: 0,
        ParagraphType.KEYWORDS: 0,
        ParagraphType.ENGLISH_ABSTRACT_HEADING: 0,
        ParagraphType.ENGLISH_ABSTRACT: 0,
        ParagraphType.ENGLISH_KEYWORDS: 0,
        ParagraphType.BODY: 0,
    }

    paragraphs = list(doc.paragraphs)
    analyses = _build_paragraph_analyses(paragraphs)
    title_index = find_title_paragraph_index(paragraphs, analyses)
    title_text = ""
    if title_index is not None:
        title_text = analyses[title_index].normalized_text

    page_setup = apply_document_layout(doc, title_text)
    emit_progress(
        progress_callback,
        1,
        "文档结构解析完成",
        f"共 {len(paragraphs)} 个段落，准备识别标题与摘要",
    )
    outline = []
    in_references = False
    in_english_abstract = False
    previous_nonempty_para_type = None
    figure_counter = 0
    table_counter = 0
    equation_paragraph_count = 0
    heading_number_state = [0, 0, 0]
    total_paragraphs = len(paragraphs)

    # ---------- 4. 遍历并格式化每个段落 ----------
    emit_progress(progress_callback, 2, "正在识别标题层级与摘要结构")
    for paragraph, analysis in zip(paragraphs, analyses):
        i = analysis.index
        text = analysis.normalized_text
        para_type = ParagraphType.TITLE if i == title_index else analysis.classified_type
        inferred_heading = False
        if para_type == ParagraphType.BODY:
            inferred_para_type = analysis.inferred_heading_type
            if inferred_para_type is not None:
                para_type = inferred_para_type
                inferred_heading = True
        if analysis.has_equation:
            equation_paragraph_count += 1
        preserve_embedded_content = analysis.has_rewrite_sensitive_content

        if total_paragraphs and (
            i == 0
            or i == total_paragraphs - 1
            or (i + 1) % max(1, total_paragraphs // 4 or 1) == 0
        ):
            emit_progress(
                progress_callback,
                2,
                f"正在识别第 {i + 1}/{total_paragraphs} 段的结构",
            )

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
            elif analysis.is_reference_entry_candidate:
                para_type = ParagraphType.REFERENCE_ENTRY
            else:
                in_references = False

        if para_type in {ParagraphType.ENGLISH_ABSTRACT_HEADING, ParagraphType.ENGLISH_ABSTRACT}:
            in_english_abstract = True
        elif in_english_abstract and text:
            if para_type == ParagraphType.ENGLISH_KEYWORDS:
                in_english_abstract = False
            elif para_type in {
                ParagraphType.TITLE,
                ParagraphType.HEADING_L1,
                ParagraphType.HEADING_L2,
                ParagraphType.HEADING_L3,
                ParagraphType.FIGURE_CAPTION,
                ParagraphType.TABLE_CAPTION,
                ParagraphType.SECTION_HEADING,
                ParagraphType.REFERENCES_HEADING,
                ParagraphType.ABSTRACT,
                ParagraphType.KEYWORDS,
            }:
                in_english_abstract = False
            else:
                para_type = ParagraphType.ENGLISH_ABSTRACT

        if (
            para_type == ParagraphType.BODY
            and text
            and analysis.is_caption_note_candidate
            and previous_nonempty_para_type in {
                ParagraphType.FIGURE_CAPTION,
                ParagraphType.TABLE_CAPTION,
                ParagraphType.CAPTION_NOTE,
            }
        ):
            para_type = ParagraphType.CAPTION_NOTE

        stats[para_type] += 1
        _append_outline_entry(outline, para_type, text)

        if para_type == ParagraphType.TITLE:
            _log_detected_paragraph("论文标题", i, text)
            format_title(
                paragraph,
                text_override=None if preserve_embedded_content else text,
            )

        elif para_type == ParagraphType.HEADING_L1:
            heading_text, explicit_parts = extract_heading_numbering(text, para_type)
            numbering_parts = resolve_heading_numbering_parts(
                para_type,
                explicit_parts,
                heading_number_state,
                allow_auto_numbering=inferred_heading,
            )
            heading_label = "推断一级标题" if inferred_heading else "一级标题"
            _log_detected_paragraph(heading_label, i, text)
            format_heading_l1(
                paragraph,
                text_override=None if preserve_embedded_content else heading_text,
            )
            if not (preserve_embedded_content and explicit_parts):
                apply_native_heading_numbering(doc, paragraph, para_type, numbering_parts)

        elif para_type == ParagraphType.HEADING_L2:
            heading_text, explicit_parts = extract_heading_numbering(text, para_type)
            numbering_parts = resolve_heading_numbering_parts(
                para_type,
                explicit_parts,
                heading_number_state,
                allow_auto_numbering=inferred_heading,
            )
            heading_label = "推断二级标题" if inferred_heading else "二级标题"
            _log_detected_paragraph(heading_label, i, text)
            format_heading_l2(
                paragraph,
                text_override=None if preserve_embedded_content else heading_text,
            )
            if not (preserve_embedded_content and explicit_parts):
                apply_native_heading_numbering(doc, paragraph, para_type, numbering_parts)

        elif para_type == ParagraphType.HEADING_L3:
            heading_text, explicit_parts = extract_heading_numbering(text, para_type)
            numbering_parts = resolve_heading_numbering_parts(
                para_type,
                explicit_parts,
                heading_number_state,
                allow_auto_numbering=inferred_heading,
            )
            heading_label = "推断三级标题" if inferred_heading else "三级标题"
            _log_detected_paragraph(heading_label, i, text)
            format_heading_l3(
                paragraph,
                text_override=None if preserve_embedded_content else heading_text,
            )
            if not (preserve_embedded_content and explicit_parts):
                apply_native_heading_numbering(doc, paragraph, para_type, numbering_parts)

        elif para_type == ParagraphType.FIGURE_CAPTION:
            figure_counter += 1
            caption_match = analysis.caption_match
            caption_text = rebuild_caption_text(ParagraphType.FIGURE_CAPTION, figure_counter, caption_match[1])
            _log_detected_paragraph("图标题", i, caption_text)
            emit_progress(progress_callback, 2, f"识别到第 {figure_counter} 张图片标题")
            format_figure_table(
                paragraph,
                text_override=None if preserve_embedded_content else caption_text,
                keep_next=False,
            )

        elif para_type == ParagraphType.TABLE_CAPTION:
            table_counter += 1
            caption_match = analysis.caption_match
            caption_text = rebuild_caption_text(ParagraphType.TABLE_CAPTION, table_counter, caption_match[1])
            _log_detected_paragraph("表标题", i, caption_text)
            emit_progress(progress_callback, 2, f"识别到第 {table_counter} 张表格标题")
            format_figure_table(
                paragraph,
                text_override=None if preserve_embedded_content else caption_text,
                keep_next=True,
            )

        elif para_type == ParagraphType.SECTION_HEADING:
            _log_detected_paragraph("非编号章节标题", i, text)
            format_heading_l1(
                paragraph,
                text_override=None if preserve_embedded_content else text,
                outline_level=None,
            )

        elif para_type == ParagraphType.REFERENCES_HEADING:
            _log_detected_paragraph("参考文献标题", i, text)
            emit_progress(progress_callback, 2, "识别到参考文献区域")
            format_references_heading(
                paragraph,
                text_override=None if preserve_embedded_content else text,
            )

        elif para_type == ParagraphType.REFERENCE_ENTRY:
            _log_detected_paragraph("参考文献条目", i, text)
            format_reference_entry(paragraph)

        elif para_type == ParagraphType.ABSTRACT:
            _log_detected_paragraph("摘要段落", i, text)
            format_abstract_or_keywords(
                paragraph,
                RE_ABSTRACT,
                preserve_runs=preserve_embedded_content,
            )

        elif para_type == ParagraphType.KEYWORDS:
            _log_detected_paragraph("关键词段", i, text)
            format_abstract_or_keywords(
                paragraph,
                RE_KEYWORDS,
                preserve_runs=preserve_embedded_content,
            )

        elif para_type == ParagraphType.ENGLISH_ABSTRACT_HEADING:
            _log_detected_paragraph("英文摘要标题", i, text)
            format_english_abstract_heading(
                paragraph,
                preserve_runs=preserve_embedded_content,
            )

        elif para_type == ParagraphType.ENGLISH_ABSTRACT:
            _log_detected_paragraph("英文摘要", i, text)
            if RE_ENGLISH_ABSTRACT.match(text):
                format_english_abstract(
                    paragraph,
                    RE_ENGLISH_ABSTRACT,
                    preserve_runs=preserve_embedded_content,
                )
            else:
                format_english_abstract(
                    paragraph,
                    preserve_runs=preserve_embedded_content,
                )

        elif para_type == ParagraphType.ENGLISH_KEYWORDS:
            _log_detected_paragraph("英文关键词", i, text)
            format_english_keywords(
                paragraph,
                preserve_runs=preserve_embedded_content,
            )

        elif para_type == ParagraphType.CAPTION_NOTE:
            _log_detected_paragraph("图表附注", i, text)
            format_caption_note(
                paragraph,
                preserve_runs=preserve_embedded_content,
            )

        else:
            # 正文段落（含空段落）
            format_body(
                paragraph,
                normalized_text=text,
                has_equation=analysis.has_equation,
                has_drawing=analysis.has_drawing,
            )

        if text:
            previous_nonempty_para_type = para_type

    # ---------- 5. 处理表格内段落 ----------
    emit_progress(progress_callback, 3, "正在应用排版规则")
    table_paragraph_count = 0
    for paragraph in iter_table_paragraphs(doc.tables):
        table_paragraph_count += 1
        if _has_equation_content(paragraph):
            equation_paragraph_count += 1
        format_body(paragraph, in_table=True)

    all_tables = list(iter_all_tables(doc.tables))
    for table_index, table in enumerate(all_tables, start=1):
        emit_progress(
            progress_callback,
            3,
            f"正在排版第 {table_index}/{len(all_tables)} 张表格",
            "应用三线表边框与表内正文格式",
        )
        format_three_line_table(table)

    heading_count = (
        stats[ParagraphType.HEADING_L1]
        + stats[ParagraphType.HEADING_L2]
        + stats[ParagraphType.HEADING_L3]
    )
    if heading_count >= MIN_TOC_HEADING_COUNT:
        emit_progress(progress_callback, 3, "正在插入自动目录字段")
    insert_table_of_contents(doc, title_index, heading_count)
    resized_image_count = constrain_inline_images(doc, progress_callback=progress_callback)
    cover_generated = False
    formatted_footnote_count = 0
    resolved_cover_info = prepare_cover_info(cover_info, title_text)
    if resolved_cover_info is not None:
        emit_progress(progress_callback, 3, "正在生成课程论文封面")
        cover_generated = generate_cover_page(doc, resolved_cover_info)
        if cover_generated:
            emit_progress(progress_callback, 3, "封面模板已插入文档首页")

    # ---------- 6. 保存输出文档 ----------
    emit_progress(progress_callback, 4, "正在生成输出文档")
    try:
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        emit_progress(progress_callback, 4, "正在统一脚注字体与字号")
        formatted_footnote_count = _save_docx_with_footnote_postprocessing(
            doc.save,
            output_file,
            max_output_bytes,
        )
        logger.info(f"排版完成！已保存至：{format_log_path(output_path)}")
    except OutputSizeLimitExceeded:
        raise
    except Exception as e:
        logger.error(f"无法保存文档 {format_log_path(output_path)}：{format_log_exception(e, output_path)}")
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
    logger.info(f"  图表附注：{stats[ParagraphType.CAPTION_NOTE]} 个")
    logger.info(f"  非编号章节标题：{stats[ParagraphType.SECTION_HEADING]} 个")
    logger.info(f"  参考文献标题：{stats[ParagraphType.REFERENCES_HEADING]} 个")
    logger.info(f"  参考文献条目：{stats[ParagraphType.REFERENCE_ENTRY]} 条")
    logger.info(f"  摘要段落：{stats[ParagraphType.ABSTRACT]} 个")
    logger.info(f"  关键词段：{stats[ParagraphType.KEYWORDS]} 个")
    logger.info(f"  英文摘要标题：{stats[ParagraphType.ENGLISH_ABSTRACT_HEADING]} 个")
    logger.info(f"  英文摘要段落：{stats[ParagraphType.ENGLISH_ABSTRACT]} 个")
    logger.info(f"  英文关键词：{stats[ParagraphType.ENGLISH_KEYWORDS]} 个")
    logger.info(f"  正文段落：{stats[ParagraphType.BODY]} 个")
    logger.info(f"  表格内段落：{table_paragraph_count} 个")
    logger.info(f"  公式段落：{equation_paragraph_count} 个")
    logger.info(f"  已统一脚注：{formatted_footnote_count} 条")
    logger.info(f"  自动缩放图片：{resized_image_count} 张")
    logger.info(f"  自动封面：{'已生成' if cover_generated else '未生成'}")
    logger.info("=" * 50)

    return {
        "stats": stats,
        "title_text": title_text,
        "page_setup": page_setup,
        "table_paragraphs": table_paragraph_count,
        "equation_paragraphs": equation_paragraph_count,
        "formatted_footnotes": formatted_footnote_count,
        "resized_images": resized_image_count,
        "cover_generated": cover_generated,
        "outline": outline,
    }


# ============================================================
# 封面+正文合并
# ============================================================
def merge_cover_and_body(
    cover_path: str,
    body_path: str,
    output_path: str,
    progress_callback=None,
    max_output_bytes=None,
):
    """
    合并封面文档和正文文档。

    封面保持原样不做排版处理，正文按学术论文规范排版，
    合并后正文部分页码从 1 开始。

    Returns:
        排版结果 dict（成功）或 False（失败）

    Raises:
        OutputSizeLimitExceeded: 中间正文或最终文档超过配置的单文件上限
    """
    import tempfile
    from contextlib import suppress

    cover_file = Path(cover_path)
    body_file = Path(body_path)

    if not cover_file.exists():
        logger.error(f"封面文件不存在：{format_log_path(cover_path)}")
        return False

    if not body_file.exists():
        logger.error(f"正文文件不存在：{format_log_path(body_path)}")
        return False

    formatted_body_path = None
    try:
        emit_progress(progress_callback, 1, "正在读取封面与正文文档")
        # 封面是外部输入；先完成校验，避免非法封面触发正文排版工作。
        cover_doc = _load_validated_document(cover_file)
        # 放进受 TTL 管理的上传目录；即使进程被强杀，遗留中间件也能被后续请求回收。
        with _ACTIVE_TEMP_PATHS_LOCK:
            with tempfile.NamedTemporaryFile(
                suffix=".docx",
                dir=body_file.parent,
                delete=False,
            ) as tmp:
                formatted_body_path = _register_active_temp_path_locked(tmp.name)

        body_result = format_academic_paper(
            body_path,
            formatted_body_path,
            progress_callback=progress_callback,
            max_output_bytes=max_output_bytes,
        )
        if not body_result:
            return False

        from docxcompose.composer import Composer

        emit_progress(progress_callback, 4, "正在合并封面与排版后的正文")
        formatted_body_doc = _load_validated_document(
            formatted_body_path,
            limits=GENERATED_DOCX_LIMITS,
        )
        body_section_count = len(formatted_body_doc.sections)

        # docxcompose 会忽略“被插入文档”末尾 body sectPr。因此
        # 必须以已排版正文为 master，把封面包装成独立节后前插，
        # 才能保留正文的页面、页眉页脚与页码属性。
        _strip_trailing_blank_paragraphs(cover_doc)
        cover_doc.add_section(WD_SECTION.NEW_PAGE)

        cover_uses_even_stories = (
            cover_doc.settings.odd_and_even_pages_header_footer
        )
        body_uses_even_stories = (
            formatted_body_doc.settings.odd_and_even_pages_header_footer
        )
        if cover_uses_even_stories or body_uses_even_stories:
            if not cover_uses_even_stories:
                _materialize_even_header_footer_from_defaults(cover_doc)
            if not body_uses_even_stories:
                _materialize_even_header_footer_from_defaults(
                    formatted_body_doc
                )
            formatted_body_doc.settings.odd_and_even_pages_header_footer = True

        _merge_inserted_custom_properties(cover_doc, formatted_body_doc)
        composer = _create_structure_preserving_composer(
            formatted_body_doc,
            composer_class=Composer,
        )
        composer.insert(0, cover_doc, remove_property_fields=False)

        # 正文各节仍在 master 末尾；据此定位首节并从 1 编页。
        merged_sections = list(formatted_body_doc.sections)
        body_first_index = len(merged_sections) - body_section_count
        _restore_inserted_header_footer_parts(
            composer,
            cover_doc,
            formatted_body_doc,
            body_first_index,
        )
        if 0 <= body_first_index < len(merged_sections):
            _set_section_page_number_start(merged_sections[body_first_index], 1)

        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        emit_progress(progress_callback, 4, "正在写入合并后的输出文档")
        emit_progress(progress_callback, 4, "正在统一合并文档中的脚注格式")
        formatted_footnote_count = _save_docx_with_footnote_postprocessing(
            composer.save,
            output_file,
            max_output_bytes,
        )
        if isinstance(body_result, dict):
            body_result["formatted_footnotes"] = formatted_footnote_count

        logger.info(f"合并完成！封面 + 排版正文 → {format_log_path(output_path)}")
        return body_result

    except OutputSizeLimitExceeded:
        raise
    except Exception as e:
        logger.error(
            f"合并文档失败: {format_log_exception(e, cover_path, body_path, output_path, formatted_body_path)}",
            exc_info=logger.isEnabledFor(logging.DEBUG),
        )
        return False
    finally:
        if formatted_body_path:
            try:
                with suppress(OSError):
                    Path(formatted_body_path).unlink()
            finally:
                release_active_temp_path(formatted_body_path)


class DocumentConcatError(Exception):
    """拼接过程中可直接展示给用户的错误（如文档损坏 / 非有效 .docx）。"""


def _is_blank_paragraph(paragraph) -> bool:
    """判断段落是否为可安全删除的空段落（无文字、无图片/对象/换行符）。"""
    if _paragraph_text_for_matching(paragraph).strip():
        return False
    if _has_rewrite_sensitive_inline_content(paragraph):
        return False
    if paragraph._element.xpath(".//w:drawing | .//w:pict | .//w:object | .//w:br | .//w:tab"):
        return False
    p_pr = paragraph._element.find(qn("w:pPr"))
    if p_pr is not None and p_pr.find(qn("w:sectPr")) is not None:
        return False
    return True


def _strip_trailing_blank_paragraphs(doc) -> int:
    """删除文档末尾连续的空段落，降低拼接后多出空白页的概率。"""
    removed = 0
    paragraph_by_element = {
        paragraph._element: paragraph
        for paragraph in doc.paragraphs
    }
    for block in reversed(list(_iter_document_body_blocks(doc))):
        if block.tag != qn("w:p"):
            break
        paragraph = paragraph_by_element.get(block)
        if paragraph is None or not _is_blank_paragraph(paragraph):
            break
        paragraph._element.getparent().remove(paragraph._element)
        removed += 1
    return removed


def _set_section_page_number_start(section, start_value: int) -> None:
    """设置某一节的页码起始值（w:pgNumType@w:start）。"""
    sect_pr = section._sectPr
    pg_num_type = sect_pr.find(qn("w:pgNumType"))
    if pg_num_type is None:
        pg_num_type = OxmlElement("w:pgNumType")
        sect_pr.insert_element_before(
            pg_num_type,
            "w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg",
            "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid",
            "w:printerSettings", "w:sectPrChange",
        )
    pg_num_type.set(qn("w:start"), str(start_value))


def _strip_header_footer_references(section) -> None:
    """移除某一节的页眉/页脚引用，使该节（如封面）不显示页眉与页码。

    只删除该节 sectPr 中的引用元素，不动被引用的页眉/页脚部件本身，
    因此与其它节共享的页脚（含页码字段）不受影响。
    """
    sect_pr = section._sectPr
    for tag in ("w:headerReference", "w:footerReference"):
        for ref in sect_pr.findall(qn(tag)):
            sect_pr.remove(ref)


def _serialize_composer_xml_root(root) -> bytes:
    """序列化合并过程中的 XML Part，供通用 ``Part`` 回写。"""
    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )


def _composer_xml_part_root(part):
    """返回 XmlPart 的活元素，或解析通用 Part 的 blob。"""
    element = getattr(part, "element", None)
    if element is not None:
        return element

    cached_root = getattr(part, "_academic_composer_xml_root", None)
    if cached_root is not None:
        return cached_root

    # 先经禁用 DTD/实体/网络的解析器消毒，再转成
    # python-docx xmlchemy 元素。docxcompose 的样式/编号迁移
    # 会读取 w:pStyle/w:rStyle 等节点的 ``.val`` 属性，
    # 普通 lxml 元素不具备该接口。
    safe_root = _parse_untrusted_ooxml_part(part.blob)
    typed_root = parse_xml(etree.tostring(safe_root, encoding="UTF-8"))
    setattr(part, "_academic_composer_xml_root", typed_root)
    return typed_root


def _commit_composer_xml_part_root(part, root) -> None:
    """通用 Part 需显式回写 blob；XmlPart 的活元素会自行序列化。"""
    if getattr(part, "element", None) is root:
        return
    part._blob = _serialize_composer_xml_root(root)
    setattr(part, "_academic_composer_xml_root", root)


def _copy_composer_element_relationships(
    composer,
    source_part,
    target_part,
    element,
) -> None:
    """复制元素中的关系目标，并把关系 ID 改写为新 Part 中的 ID。"""
    relationship_namespace = qn("r:id").partition("}")[0] + "}"
    legacy_relationship_attribute = qn("o:relid")
    relationship_ids = {}

    for node in element.iter():
        for attribute_name, source_rid in list(node.attrib.items()):
            if not (
                attribute_name.startswith(relationship_namespace)
                or attribute_name == legacy_relationship_attribute
            ):
                continue

            relationship = source_part.rels.get(source_rid)
            if relationship is None:
                raise ValueError("inserted story element has a missing relationship")

            target_rid = relationship_ids.get(source_rid)
            if target_rid is None:
                target_relationship = composer.add_relationship(
                    source_part,
                    target_part,
                    relationship,
                )
                target_rid = target_relationship.rId
                relationship_ids[source_rid] = target_rid
            node.set(attribute_name, target_rid)


def _rewrite_composer_custom_xml_bindings_in_part(composer, part) -> None:
    """在复制独立 story Part 前改写其跨 Part 语义引用。"""
    root = _composer_xml_part_root(part)
    composer._rewrite_custom_xml_bindings(root)
    composer._rewrite_bookmark_names_and_references(root)
    _commit_composer_xml_part_root(part, root)


def _find_story_item_by_id(root, item_tag, item_id: int):
    """按数值 ID 查找唯一的批注/脚注/尾注元素。"""
    matches = []
    for item in root.findall(qn(item_tag)):
        parsed_id = _parse_bounded_decimal(
            item.get(qn("w:id")),
            MAX_OOXML_DECIMAL_NUMBER,
        )
        if parsed_id == item_id:
            matches.append(item)
    if len(matches) != 1:
        raise ValueError("inserted story reference does not resolve uniquely")
    return matches[0]


_SPECIAL_NOTE_CONTENT_TAGS = frozenset(
    {
        qn("w:separator"),
        qn("w:continuationSeparator"),
        qn("w:continuationNotice"),
    }
)


def _note_settings_spec(relationship_type: str):
    if relationship_type == RT.FOOTNOTES:
        return "w:footnotePr", "w:footnote"
    if relationship_type == RT.ENDNOTES:
        return "w:endnotePr", "w:endnote"
    return None


def _configured_special_note_ids(source_doc, relationship_type: str) -> set[str]:
    """读取 settings.xml 中 WPS/Word 使用的分隔符 ID 引用。"""
    spec = _note_settings_spec(relationship_type)
    if spec is None:
        return set()
    properties_tag, reference_tag = spec
    properties = source_doc.settings.element.find(qn(properties_tag))
    if properties is None:
        return set()

    special_ids = set()
    for reference in properties.findall(qn(reference_tag)):
        normalized_id = _normalize_ooxml_integer(reference.get(qn("w:id")))
        if normalized_id is None:
            raise ValueError("note settings contain an invalid special-note ID")
        if normalized_id in special_ids:
            raise ValueError("note settings contain duplicate special-note IDs")
        special_ids.add(normalized_id)
    return special_ids


def _note_item_has_special_content(item) -> bool:
    return any(
        node.tag in _SPECIAL_NOTE_CONTENT_TAGS
        for node in item.iter()
    )


def _is_special_note_item(item, special_ids: set[str]) -> bool:
    normalized_id = _normalize_ooxml_integer(item.get(qn("w:id")))
    return (
        item.get(qn("w:type")) in NOTE_SPECIAL_TYPES
        or (normalized_id is not None and normalized_id in special_ids)
        or _is_negative_ooxml_decimal(item.get(qn("w:id")))
        or _note_item_has_special_content(item)
    )


def _special_note_item_ids(
    source_doc,
    source_root,
    relationship_type: str,
    item_tag: str,
) -> set[str]:
    """合并 settings、w:type、负 ID 与内容标记得到特殊注释 ID。"""
    configured_ids = _configured_special_note_ids(
        source_doc,
        relationship_type,
    )
    special_ids = set(configured_ids)
    item_counts = {}
    for item in source_root.findall(qn(item_tag)):
        normalized_id = _normalize_ooxml_integer(item.get(qn("w:id")))
        if normalized_id is None:
            if _is_special_note_item(item, special_ids):
                raise ValueError("special note has an invalid ID")
            continue
        item_counts[normalized_id] = item_counts.get(normalized_id, 0) + 1
        if _is_special_note_item(item, special_ids):
            special_ids.add(normalized_id)

    for configured_id in configured_ids:
        if item_counts.get(configured_id) != 1:
            raise ValueError("special-note settings reference does not resolve uniquely")
    return special_ids


def _copy_special_note_settings_references(
    source_doc,
    target_doc,
    relationship_type: str,
    copied_special_ids: set[str],
) -> None:
    """把正数分隔符在 settings.xml 中的引用同步到 master。"""
    spec = _note_settings_spec(relationship_type)
    if spec is None or not copied_special_ids:
        return
    properties_tag, reference_tag = spec
    source_properties = source_doc.settings.element.find(qn(properties_tag))
    if source_properties is None:
        return

    references = []
    for source_reference in source_properties.findall(qn(reference_tag)):
        normalized_id = _normalize_ooxml_integer(
            source_reference.get(qn("w:id"))
        )
        if normalized_id in copied_special_ids:
            references.append(source_reference)
    if not references:
        return

    target_settings = target_doc.settings.element
    target_properties = target_settings.find(qn(properties_tag))
    if target_properties is None:
        target_properties = OxmlElement(properties_tag)
        target_settings.insert_element_before(
            target_properties,
            "w:endnotePr" if properties_tag == "w:footnotePr" else "w:compat",
            "w:compat",
            "w:docVars",
            "w:rsids",
            "m:mathPr",
            "w:attachedSchema",
            "w:themeFontLang",
        )

    target_ids = {
        _normalize_ooxml_integer(reference.get(qn("w:id")))
        for reference in target_properties.findall(qn(reference_tag))
    }
    for source_reference in references:
        normalized_id = _normalize_ooxml_integer(
            source_reference.get(qn("w:id"))
        )
        if normalized_id in target_ids:
            continue
        target_properties.append(deepcopy(source_reference))
        target_ids.add(normalized_id)


def _index_story_items(
    root,
    item_tag,
    special_ids: set[str] | None = None,
) -> dict[int, object]:
    """一次建立 story 子项索引，避免大量脚注/批注时重复扫描 XML。"""
    special_ids = special_ids or set()
    index = {}
    for item in root.findall(qn(item_tag)):
        normalized_id = _normalize_ooxml_integer(item.get(qn("w:id")))
        if normalized_id in special_ids:
            continue
        parsed_id = _parse_bounded_decimal(
            item.get(qn("w:id")),
            MAX_OOXML_DECIMAL_NUMBER,
        )
        if parsed_id is None:
            if (
                _is_special_note_item(item, special_ids)
            ):
                continue
            raise ValueError("story Part contains an invalid item ID")
        if parsed_id in index:
            raise ValueError("story Part contains duplicate or invalid item IDs")
        index[parsed_id] = item
    return index


def _used_story_item_ids(root, item_tag) -> set[int]:
    """返回目标 story Part 中已使用的非负数 ID。"""
    used_ids = set()
    for item in root.findall(qn(item_tag)):
        parsed_id = _parse_bounded_decimal(
            item.get(qn("w:id")),
            MAX_OOXML_DECIMAL_NUMBER,
        )
        if parsed_id is not None:
            used_ids.add(parsed_id)
    return used_ids


def _allocate_story_item_id(
    root,
    item_tag,
    preferred_id: int,
    used_ids: set[int] | None = None,
) -> int:
    """优先保留源 ID，冲突时选取最小可用正整数。"""
    if used_ids is None:
        used_ids = _used_story_item_ids(root, item_tag)
    if preferred_id not in used_ids:
        return preferred_id

    candidate = 1
    while candidate in used_ids:
        candidate += 1
    if candidate > MAX_OOXML_DECIMAL_NUMBER:  # pragma: no cover - 受 XML 大小限制
        raise ValueError("story Part has no available item ID")
    return candidate


def _get_or_create_composer_story_part(
    composer,
    source_doc,
    source_part,
    *,
    relationship_type: str,
    content_type: str,
    root_tag: str,
    item_tag: str,
    default_partname: str,
    next_partname_template: str,
    copy_special_notes: bool,
    special_ids: set[str],
):
    """获取 master 的批注/脚注/尾注 Part，缺失时创建最小合法 Part。"""
    try:
        target_part = composer.doc.part.part_related_by(relationship_type)
    except KeyError:
        package = composer.doc.part.package
        if package is None:  # pragma: no cover - 已打开 Document 始终有 Package
            raise ValueError("document has no package")

        partname = composer._allocate_copied_partname(
            PackURI(default_partname)
        )

        target_root = OxmlElement(root_tag)
        target_part = Part(
            partname,
            content_type,
            _serialize_composer_xml_root(target_root),
            package,
        )
        composer.doc.part.relate_to(target_part, relationship_type)

        if copy_special_notes:
            source_root = _composer_xml_part_root(source_part)
            copied_special_ids = set()
            for source_item in source_root.findall(qn(item_tag)):
                if not _is_special_note_item(source_item, special_ids):
                    continue
                normalized_id = _normalize_ooxml_integer(
                    source_item.get(qn("w:id"))
                )
                if normalized_id is None:
                    raise ValueError("special note has an invalid ID")
                copied_item = deepcopy(source_item)
                composer._retain_inserted_default_formatting(
                    source_doc,
                    copied_item,
                )
                composer._materialize_inserted_theme_references(
                    source_doc,
                    copied_item,
                )
                composer._rewrite_custom_xml_bindings(copied_item)
                composer._rewrite_bookmark_names_and_references(copied_item)
                _copy_composer_element_relationships(
                    composer,
                    source_part,
                    target_part,
                    copied_item,
                )
                composer.add_styles(source_doc, copied_item)
                composer.add_numberings(source_doc, copied_item)
                target_root.append(copied_item)
                copied_special_ids.add(normalized_id)
            _commit_composer_xml_part_root(target_part, target_root)
            _copy_special_note_settings_references(
                source_doc,
                composer.doc,
                relationship_type,
                copied_special_ids,
            )

    target_root = _composer_xml_part_root(target_part)
    if target_root.tag != qn(root_tag):
        raise ValueError("story relationship targets a Part with the wrong root")
    return target_part, target_root


def _copy_composer_story_references(
    composer,
    source_doc,
    element,
    *,
    reference_tags: tuple[str, ...],
    relationship_type: str,
    content_type: str,
    root_tag: str,
    item_tag: str,
    default_partname: str,
    next_partname_template: str,
    copy_special_notes: bool = False,
) -> None:
    """复制行内 story 引用指向的内容，并对冲突 ID 做稳定重映射。"""
    references = [
        reference
        for reference_tag in reference_tags
        for reference in element.findall(".//" + qn(reference_tag))
    ]
    if not references:
        return

    if (
        relationship_type == RT.COMMENTS
        and _has_modern_comment_sidecars(composer.doc)
    ):
        raise DocumentConcatError(
            "目标文档包含现代线程批注（回复/已解决状态），"
            "当前无法安全插入普通批注；请先在 Word 中删除这些批注"
            "或转为普通批注。"
        )

    try:
        source_part = source_doc.part.part_related_by(relationship_type)
    except KeyError as exc:
        raise ValueError("inserted story references a missing Part") from exc

    source_root = _composer_xml_part_root(source_part)
    if source_root.tag != qn(root_tag):
        raise ValueError("inserted story Part has the wrong root")

    special_ids = (
        _special_note_item_ids(
            source_doc,
            source_root,
            relationship_type,
            item_tag,
        )
        if copy_special_notes
        else set()
    )

    target_part, target_root = _get_or_create_composer_story_part(
        composer,
        source_doc,
        source_part,
        relationship_type=relationship_type,
        content_type=content_type,
        root_tag=root_tag,
        item_tag=item_tag,
        default_partname=default_partname,
        next_partname_template=next_partname_template,
        copy_special_notes=copy_special_notes,
        special_ids=special_ids,
    )
    id_mapping = getattr(composer, "_academic_story_id_mapping", None)
    if not isinstance(id_mapping, dict):
        id_mapping = {}
        setattr(composer, "_academic_story_id_mapping", id_mapping)
    source_item_indexes = getattr(composer, "_academic_story_item_indexes", None)
    if not isinstance(source_item_indexes, dict):
        source_item_indexes = {}
        setattr(composer, "_academic_story_item_indexes", source_item_indexes)
    source_index_key = (id(source_part), item_tag)
    source_item_index = source_item_indexes.get(source_index_key)
    if source_item_index is None:
        source_item_index = _index_story_items(
            source_root,
            item_tag,
            special_ids,
        )
        source_item_indexes[source_index_key] = source_item_index

    target_used_indexes = getattr(composer, "_academic_story_used_ids", None)
    if not isinstance(target_used_indexes, dict):
        target_used_indexes = {}
        setattr(composer, "_academic_story_used_ids", target_used_indexes)
    target_index_key = (id(target_part), item_tag)
    target_used_ids = target_used_indexes.get(target_index_key)
    if target_used_ids is None:
        target_used_ids = _used_story_item_ids(target_root, item_tag)
        target_used_indexes[target_index_key] = target_used_ids

    target_changed = False
    for reference in references:
        source_id = _parse_bounded_decimal(
            reference.get(qn("w:id")),
            MAX_OOXML_DECIMAL_NUMBER,
        )
        if source_id is None:
            raise ValueError("inserted story reference has an invalid ID")
        if str(source_id) in special_ids:
            raise ValueError("document content references a special note")

        mapping_key = (id(source_part), relationship_type, source_id)
        target_id = id_mapping.get(mapping_key)
        if target_id is None:
            source_item = source_item_index.get(source_id)
            if source_item is None:
                raise ValueError("inserted story reference does not resolve uniquely")
            if copy_special_notes and _is_special_note_item(
                source_item,
                special_ids,
            ):
                raise ValueError("document content references a special note")

            target_id = _allocate_story_item_id(
                target_root,
                item_tag,
                source_id,
                used_ids=target_used_ids,
            )
            copied_item = deepcopy(source_item)
            composer._retain_inserted_default_formatting(
                source_doc,
                copied_item,
            )
            composer._materialize_inserted_theme_references(
                source_doc,
                copied_item,
            )
            copied_item.set(qn("w:id"), str(target_id))
            composer._rewrite_custom_xml_bindings(copied_item)
            composer._rewrite_bookmark_names_and_references(copied_item)
            _copy_composer_element_relationships(
                composer,
                source_part,
                target_part,
                copied_item,
            )
            composer.add_styles(source_doc, copied_item)
            composer.add_numberings(source_doc, copied_item)
            target_root.append(copied_item)
            id_mapping[mapping_key] = target_id
            target_used_ids.add(target_id)
            target_changed = True

        reference.set(qn("w:id"), str(target_id))

    if target_changed:
        _commit_composer_xml_part_root(target_part, target_root)


_CUSTOM_XML_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
)
_CUSTOM_XML_DATASTORE_TAG = f"{{{_CUSTOM_XML_NAMESPACE}}}datastoreItem"
_CUSTOM_XML_ITEM_ID_ATTRIBUTE = f"{{{_CUSTOM_XML_NAMESPACE}}}itemID"
_CUSTOM_XML_ITEM_ID_RE = re.compile(
    r"^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-"
    r"[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$"
)
_WORD_2012_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml"
_WORD_2012_DATA_BINDING_TAG = f"{{{_WORD_2012_NAMESPACE}}}dataBinding"
_CUSTOM_XML_DATA_BINDING_TAGS = (
    qn("w:dataBinding"),
    _WORD_2012_DATA_BINDING_TAG,
)
_DRAWINGML_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main"
_DRAWING_2010_NAMESPACE = "http://schemas.microsoft.com/office/drawing/2010/main"
_MARKUP_COMPATIBILITY_NAMESPACE = (
    "http://schemas.openxmlformats.org/markup-compatibility/2006"
)
_XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace"
_DRAWING_2016_SVG_NAMESPACE = (
    "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
)
_DRAWING_2016_11_NAMESPACE = (
    "http://schemas.microsoft.com/office/drawing/2016/11/main"
)
_DRAWING_2017_MODEL3D_NAMESPACE = (
    "http://schemas.microsoft.com/office/drawing/2017/model3d"
)
_WEB_EXTENSIONS_2010_NAMESPACE = (
    "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
)
_MODEL3D_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2017/06/relationships/model3d"
)
_MODEL3D_CONTENT_TYPE = "model/gltf-binary"
_RELATIONSHIP_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
_RELATIONSHIP_ID_ATTRIBUTE = f"{{{_RELATIONSHIP_NAMESPACE}}}id"
_RELATIONSHIP_EMBED_ATTRIBUTE = f"{{{_RELATIONSHIP_NAMESPACE}}}embed"
_RELATIONSHIP_LINK_ATTRIBUTE = f"{{{_RELATIONSHIP_NAMESPACE}}}link"
_DRAWINGML_IMAGE_TAGS = frozenset(
    {
        f"{{{_DRAWINGML_NAMESPACE}}}blip",
        f"{{{_DRAWING_2016_SVG_NAMESPACE}}}svgBlip",
        f"{{{_DRAWING_2017_MODEL3D_NAMESPACE}}}blip",
        f"{{{_WEB_EXTENSIONS_2010_NAMESPACE}}}snapshot",
    }
)
_DRAWINGML_EMBEDDED_IMAGE_ONLY_TAGS = frozenset(
    {f"{{{_DRAWING_2010_NAMESPACE}}}imgLayer"}
)
_DRAWINGML_HYPERLINK_TAGS = frozenset(
    f"{{{_DRAWINGML_NAMESPACE}}}{name}"
    for name in ("hlinkClick", "hlinkHover", "hlinkMouseOver")
)
_DRAWINGML_MEDIA_RELATIONSHIP_SPECS = (
    (
        f"{{{_DRAWINGML_NAMESPACE}}}snd",
        _RELATIONSHIP_EMBED_ATTRIBUTE,
        RT.AUDIO,
        True,
    ),
    (
        f"{{{_DRAWINGML_NAMESPACE}}}wavAudioFile",
        _RELATIONSHIP_EMBED_ATTRIBUTE,
        RT.AUDIO,
        True,
    ),
    (
        f"{{{_DRAWINGML_NAMESPACE}}}audioFile",
        _RELATIONSHIP_LINK_ATTRIBUTE,
        RT.AUDIO,
        False,
    ),
    (
        f"{{{_DRAWINGML_NAMESPACE}}}videoFile",
        _RELATIONSHIP_LINK_ATTRIBUTE,
        RT.VIDEO,
        False,
    ),
    (
        f"{{{_DRAWINGML_NAMESPACE}}}quickTimeFile",
        _RELATIONSHIP_LINK_ATTRIBUTE,
        RT.VIDEO,
        False,
    ),
)
_DRAWINGML_REQUIRED_RELATIONSHIP_ATTRIBUTES = {
    **{
        tag: attribute_name
        for tag, attribute_name, _relationship_type, _must_be_internal
        in _DRAWINGML_MEDIA_RELATIONSHIP_SPECS
    },
    f"{{{_DRAWING_2016_11_NAMESPACE}}}picAttrSrcUrl": (
        _RELATIONSHIP_ID_ATTRIBUTE
    ),
    f"{{{_DRAWING_2017_MODEL3D_NAMESPACE}}}attrSrcUrl": (
        _RELATIONSHIP_ID_ATTRIBUTE
    ),
}
_EMBEDDED_FONT_TAG_NAMES = (
    "w:embedRegular",
    "w:embedBold",
    "w:embedItalic",
    "w:embedBoldItalic",
)
_MODERN_COMMENT_RELATIONSHIP_TYPES = frozenset(
    {
        "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
        "http://schemas.microsoft.com/office/2011/relationships/people",
        "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
        "http://schemas.microsoft.com/office/2016/09/relationships/people",
        "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
    }
)

# A Word package has one document-level theme. ``docxcompose`` keeps the
# master's theme when inserting another document, so theme-only references in
# inserted XML must be materialized from the source theme first.
_THEME_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main"
_THEME_TAG = f"{{{_THEME_NAMESPACE}}}theme"
_THEME_FONT_SCHEME_TAG = f"{{{_THEME_NAMESPACE}}}fontScheme"
_THEME_COLOR_SCHEME_TAG = f"{{{_THEME_NAMESPACE}}}clrScheme"
_THEME_FORMAT_SCHEME_TAG = f"{{{_THEME_NAMESPACE}}}fmtScheme"
_THEME_SCHEME_COLOR_TAG = f"{{{_THEME_NAMESPACE}}}schemeClr"
_THEME_SRGB_COLOR_TAG = f"{{{_THEME_NAMESPACE}}}srgbClr"
_THEME_SYS_COLOR_TAG = f"{{{_THEME_NAMESPACE}}}sysClr"
_THEME_FONT_TAG = f"{{{_THEME_NAMESPACE}}}font"
_THEME_FONT_SLOT_TAGS = {
    "latin": f"{{{_THEME_NAMESPACE}}}latin",
    "ea": f"{{{_THEME_NAMESPACE}}}ea",
    "cs": f"{{{_THEME_NAMESPACE}}}cs",
}
_THEME_EXTENSION_LIST_TAG = f"{{{_THEME_NAMESPACE}}}extLst"
_THEME_COLOR_MODEL_TAGS = frozenset(
    f"{{{_THEME_NAMESPACE}}}{name}"
    for name in (
        "scrgbClr",
        "srgbClr",
        "hslClr",
        "sysClr",
        "schemeClr",
        "prstClr",
    )
)
_THEME_FORMAT_SCHEME_LIST_SPECS = (
    (
        f"{{{_THEME_NAMESPACE}}}fillStyleLst",
        frozenset(
            f"{{{_THEME_NAMESPACE}}}{name}"
            for name in (
                "noFill",
                "solidFill",
                "gradFill",
                "blipFill",
                "pattFill",
                "grpFill",
            )
        ),
    ),
    (
        f"{{{_THEME_NAMESPACE}}}lnStyleLst",
        frozenset({f"{{{_THEME_NAMESPACE}}}ln"}),
    ),
    (
        f"{{{_THEME_NAMESPACE}}}effectStyleLst",
        frozenset({f"{{{_THEME_NAMESPACE}}}effectStyle"}),
    ),
    (
        f"{{{_THEME_NAMESPACE}}}bgFillStyleLst",
        frozenset(
            f"{{{_THEME_NAMESPACE}}}{name}"
            for name in (
                "noFill",
                "solidFill",
                "gradFill",
                "blipFill",
                "pattFill",
                "grpFill",
            )
        ),
    ),
)
_THEME_COLOR_CHILDREN = (
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
)
_THEME_FONT_REFERENCE_SPECS = (
    (qn("w:asciiTheme"), qn("w:ascii"), "latin"),
    (qn("w:hAnsiTheme"), qn("w:hAnsi"), "latin"),
    (qn("w:eastAsiaTheme"), qn("w:eastAsia"), "eastAsia"),
    (qn("w:cstheme"), qn("w:cs"), "bidi"),
)
_THEME_COLOR_ALIAS = {
    "dark1": "dk1",
    "light1": "lt1",
    "dark2": "dk2",
    "light2": "lt2",
    "background1": "lt1",
    "text1": "dk1",
    "background2": "lt2",
    "text2": "dk2",
    "t1": "dk1",
    "t2": "dk2",
    "hyperlink": "hlink",
    "followedhyperlink": "folHlink",
    "folhlink": "folHlink",
}
_DRAWING_THEME_LOGICAL_ALIASES = {
    "bg1": "bg1",
    "tx1": "tx1",
    "t1": "tx1",
    "bg2": "bg2",
    "tx2": "tx2",
    "t2": "tx2",
    "accent1": "accent1",
    "accent2": "accent2",
    "accent3": "accent3",
    "accent4": "accent4",
    "accent5": "accent5",
    "accent6": "accent6",
    "hlink": "hlink",
    "hyperlink": "hlink",
    "folhlink": "folhlink",
    "followedhyperlink": "folhlink",
}
_DRAWING_THEME_DEFAULT_MAPPING = {
    "bg1": "lt1",
    "tx1": "dk1",
    "bg2": "lt2",
    "tx2": "dk2",
    "accent1": "accent1",
    "accent2": "accent2",
    "accent3": "accent3",
    "accent4": "accent4",
    "accent5": "accent5",
    "accent6": "accent6",
    "hlink": "hlink",
    "folhlink": "folHlink",
}
_THEME_FONT_TOKEN_RE = re.compile(
    r"^(?P<major>major|minor)(?P<slot>Ascii|HAnsi|EastAsia|Bidi)$",
    re.IGNORECASE,
)
_HEX_COLOR_RE = re.compile(r"^[0-9A-Fa-f]{6}$")
_HEX_BYTE_RE = re.compile(r"^[0-9A-Fa-f]{2}$")
_THEME_FORMAT_REFERENCE_TAGS = frozenset(
    f"{{{_THEME_NAMESPACE}}}{name}"
    for name in ("fillRef", "lnRef", "effectRef")
)
_THEME_FONT_REFERENCE_TAG = f"{{{_THEME_NAMESPACE}}}fontRef"
_THEME_OVERRIDE_TAG = f"{{{_THEME_NAMESPACE}}}themeOverride"
_THEME_COLOR_MAPPING_OVERRIDE_TAGS = frozenset(
    f"{{{_THEME_NAMESPACE}}}{name}"
    for name in ("clrMap", "overrideClrMapping")
)
_CHART_NAMESPACE = (
    "http://schemas.openxmlformats.org/drawingml/2006/chart"
)
_CHARTEX_NAMESPACE = (
    "http://schemas.microsoft.com/office/drawing/2014/chartex"
)
_CHARTEX_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2014/relationships/chartEx"
)
_CHARTEX_CONTENT_TYPE = "application/vnd.ms-office.chartex+xml"
_CHART_STYLE_NAMESPACE = (
    "http://schemas.microsoft.com/office/drawing/2012/chartStyle"
)
_CHART_STYLE_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2011/relationships/chartStyle"
)
_CHART_COLOR_STYLE_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"
)
_CHART_STYLE_CONTENT_TYPE = "application/vnd.ms-office.chartstyle+xml"
_CHART_COLOR_STYLE_CONTENT_TYPE = (
    "application/vnd.ms-office.chartcolorstyle+xml"
)
_CHART_STYLE_ROOT_TAG = f"{{{_CHART_STYLE_NAMESPACE}}}chartStyle"
_CHART_COLOR_STYLE_ROOT_TAG = f"{{{_CHART_STYLE_NAMESPACE}}}colorStyle"
_LEGACY_SPREADSHEET_COLOR_INDEX_ATTRIBUTE = (
    f"{{{_DRAWING_2010_NAMESPACE}}}legacySpreadsheetColorIndex"
)
_CHART_MC_ATTRIBUTE_LOCAL_NAMES = frozenset(
    {
        "Ignorable",
        "ProcessContent",
        "PreserveAttributes",
        "PreserveElements",
        "MustUnderstand",
    }
)
_CHART_STYLE_UNDERSTOOD_ELEMENT_NAMESPACES = frozenset(
    """
    http://schemas.openxmlformats.org/drawingml/2006/main
    http://schemas.microsoft.com/office/drawing/2010/main
    http://schemas.microsoft.com/office/drawing/2012/main
    http://schemas.openxmlformats.org/officeDocument/2006/characteristics
    http://schemas.openxmlformats.org/officeDocument/2006/extended-properties
    http://schemas.microsoft.com/office/2006/activeX
    http://schemas.openxmlformats.org/officeDocument/2006/bibliography
    http://schemas.openxmlformats.org/drawingml/2006/chart
    http://schemas.microsoft.com/office/drawing/2007/8/2/chart
    http://schemas.microsoft.com/office/drawing/2012/chart
    http://schemas.microsoft.com/office/2006/customDocumentInformationPanel
    http://schemas.openxmlformats.org/drawingml/2006/chartDrawing
    http://schemas.microsoft.com/office/drawing/2010/chartDrawing
    http://schemas.microsoft.com/office/drawing/2010/compatibility
    http://schemas.openxmlformats.org/drawingml/2006/compatibility
    http://schemas.openxmlformats.org/package/2006/metadata/core-properties
    http://schemas.microsoft.com/office/2006/coverPageProps
    http://schemas.microsoft.com/office/drawing/2012/chartStyle
    http://schemas.microsoft.com/office/2006/metadata/contentType
    http://purl.org/dc/elements/1.1/
    http://purl.org/dc/terms/
    http://schemas.openxmlformats.org/drawingml/2006/diagram
    http://schemas.microsoft.com/office/drawing/2010/diagram
    http://schemas.openxmlformats.org/officeDocument/2006/customXml
    http://schemas.microsoft.com/office/drawing/2008/diagram
    http://www.w3.org/2003/04/emma
    http://www.w3.org/2003/InkML
    http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas
    http://schemas.microsoft.com/office/2006/metadata/longProperties
    http://schemas.openxmlformats.org/officeDocument/2006/math
    http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes
    http://schemas.openxmlformats.org/markup-compatibility/2006
    http://schemas.microsoft.com/office/office/2011/9/metroDictionary
    http://schemas.microsoft.com/ink/2010/main
    http://schemas.microsoft.com/office/2006/01/customui
    http://schemas.microsoft.com/office/2009/07/customui
    http://schemas.microsoft.com/office/2006/metadata/customXsn
    urn:schemas-microsoft-com:office:office
    http://schemas.openxmlformats.org/officeDocument/2006/custom-properties
    http://schemas.openxmlformats.org/presentationml/2006/main
    http://schemas.microsoft.com/office/powerpoint/2010/main
    http://schemas.microsoft.com/office/powerpoint/2012/main
    http://schemas.microsoft.com/office/internal/2007/ofapi/packaging
    http://schemas.microsoft.com/office/2007/6/19/audiovideo
    http://schemas.openxmlformats.org/drawingml/2006/picture
    http://schemas.microsoft.com/office/drawing/2010/picture
    http://schemas.microsoft.com/projectml/2012/main
    http://schemas.microsoft.com/office/powerpoint/2012/roamingSettings
    urn:schemas-microsoft-com:office:powerpoint
    http://schemas.openxmlformats.org/officeDocument/2006/relationships
    http://schemas.openxmlformats.org/schemaLibrary/2006/main
    http://schemas.microsoft.com/office/drawing/2010/slicer
    http://schemas.microsoft.com/office/thememl/2012/main
    http://schemas.microsoft.com/office/drawing/2012/timeslicer
    urn:schemas-microsoft-com:vml
    http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes
    http://schemas.openxmlformats.org/wordprocessingml/2006/main
    urn:schemas-microsoft-com:office:word
    http://schemas.microsoft.com/office/word/2010/wordml
    http://schemas.microsoft.com/office/word/2012/wordml
    http://schemas.microsoft.com/office/webextensions/webextension/2010/11
    http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11
    http://schemas.microsoft.com/office/word/2006/wordml
    http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing
    http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing
    http://schemas.microsoft.com/office/word/2012/wordprocessingDrawing
    http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas
    http://schemas.microsoft.com/office/word/2010/wordprocessingGroup
    http://schemas.microsoft.com/office/word/2010/wordprocessingShape
    http://schemas.openxmlformats.org/spreadsheetml/2006/main
    http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac
    http://schemas.microsoft.com/office/spreadsheetml/2009/9/main
    http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac
    http://schemas.microsoft.com/office/spreadsheetml/2010/11/main
    http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac
    http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing
    http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing
    http://www.w3.org/XML/1998/namespace
    http://schemas.microsoft.com/office/excel/2006/main
    http://www.w3.org/2001/XMLSchema
    urn:schemas-microsoft-com:office:excel
    """.split()
)
_CHART_STYLE_UNDERSTOOD_ATTRIBUTE_NAMESPACES = (
    _CHART_STYLE_UNDERSTOOD_ELEMENT_NAMESPACES
    | frozenset({_RELATIONSHIP_NAMESPACE})
)
_DRAWING_EXTENSION_TAG = f"{{{_THEME_NAMESPACE}}}ext"
_CHART_STYLE_CHILDREN = (
    "axisTitle",
    "categoryAxis",
    "chartArea",
    "dataLabel",
    "dataLabelCallout",
    "dataPoint",
    "dataPoint3D",
    "dataPointLine",
    "dataPointMarker",
    "dataPointMarkerLayout",
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
    "extLst",
)
_OPTIONAL_CHART_STYLE_CHILDREN = frozenset(
    {"dataLabelCallout", "dataPointMarkerLayout", "extLst"}
)
_CHART_STYLE_ENTRY_CHILDREN = (
    "lnRef",
    "lineWidthScale",
    "fillRef",
    "effectRef",
    "fontRef",
    "spPr",
    "defRPr",
    "bodyPr",
    "extLst",
)
_CHART_STYLE_ENTRY_OPTIONAL_CHILDREN = frozenset(
    {"lineWidthScale", "spPr", "defRPr", "bodyPr", "extLst"}
)
_CHART_COLOR_TRANSFORM_NAMES = frozenset(
    {
        "tint", "shade", "comp", "inv", "gray", "alpha", "alphaOff",
        "alphaMod", "hue", "hueOff", "hueMod", "sat", "satOff",
        "satMod", "lum", "lumOff", "lumMod", "red", "redOff",
        "redMod", "green", "greenOff", "greenMod", "blue", "blueOff",
        "blueMod", "gamma", "invGamma",
    }
)
_CHART_EMPTY_COLOR_TRANSFORMS = frozenset(
    {"comp", "inv", "gray", "gamma", "invGamma"}
)
_CHART_SYSTEM_COLOR_VALUES = frozenset(
    """
    scrollBar background activeCaption inactiveCaption menu window windowFrame
    menuText windowText captionText activeBorder inactiveBorder appWorkspace
    highlight highlightText btnFace btnShadow grayText btnText
    inactiveCaptionText btnHighlight 3dDkShadow 3dLight infoText infoBk hotLight
    gradientActiveCaption gradientInactiveCaption menuHighlight menuBar
    """.split()
)
_CHART_SCHEME_COLOR_VALUES = frozenset(
    """
    bg1 tx1 bg2 tx2 accent1 accent2 accent3 accent4 accent5 accent6 hlink
    folHlink phClr dk1 lt1 dk2 lt2
    """.split()
)
_CHART_PRESET_COLOR_VALUES = frozenset(
    """
    aliceBlue antiqueWhite aqua aquamarine azure beige bisque black
    blanchedAlmond blue blueViolet brown burlyWood cadetBlue chartreuse
    chocolate coral cornflowerBlue cornsilk crimson cyan dkBlue dkCyan
    dkGoldenrod dkGray dkGreen dkKhaki dkMagenta dkOliveGreen dkOrange dkOrchid
    dkRed dkSalmon dkSeaGreen dkSlateBlue dkSlateGray dkTurquoise dkViolet
    deepPink deepSkyBlue dimGray dodgerBlue firebrick floralWhite forestGreen
    fuchsia gainsboro ghostWhite gold goldenrod gray green greenYellow honeydew
    hotPink indianRed indigo ivory khaki lavender lavenderBlush lawnGreen
    lemonChiffon ltBlue ltCoral ltCyan ltGoldenrodYellow ltGray ltGreen ltPink
    ltSalmon ltSeaGreen ltSkyBlue ltSlateGray ltSteelBlue ltYellow lime
    limeGreen linen magenta maroon medAquamarine medBlue medOrchid medPurple
    medSeaGreen medSlateBlue medSpringGreen medTurquoise medVioletRed
    midnightBlue mintCream mistyRose moccasin navajoWhite navy oldLace olive
    oliveDrab orange orangeRed orchid paleGoldenrod paleGreen paleTurquoise
    paleVioletRed papayaWhip peachPuff peru pink plum powderBlue purple red
    rosyBrown royalBlue saddleBrown salmon sandyBrown seaGreen seaShell sienna
    silver skyBlue slateBlue slateGray snow springGreen steelBlue tan teal
    thistle tomato turquoise violet wheat white whiteSmoke yellow yellowGreen
    darkBlue darkCyan darkGoldenrod darkGray darkGrey darkGreen darkKhaki
    darkMagenta darkOliveGreen darkOrange darkOrchid darkRed darkSalmon
    darkSeaGreen darkSlateBlue darkSlateGray darkSlateGrey darkTurquoise
    darkViolet lightBlue lightCoral lightCyan lightGoldenrodYellow lightGray
    lightGrey lightGreen lightPink lightSalmon lightSeaGreen lightSkyBlue
    lightSlateGray lightSlateGrey lightSteelBlue lightYellow mediumAquamarine
    mediumBlue mediumOrchid mediumPurple mediumSeaGreen mediumSlateBlue
    mediumSpringGreen mediumTurquoise mediumVioletRed dkGrey dimGrey dkSlateGrey
    grey ltGrey ltSlateGrey slateGrey
    """.split()
)
_CHART_BLACK_WHITE_MODE_VALUES = frozenset(
    "clr auto gray ltGray invGray grayWhite blackGray blackWhite black white hidden".split()
)
_CHART_TEXT_UNDERLINE_VALUES = frozenset(
    """
    none words sng dbl heavy dotted dottedHeavy dash dashHeavy dashLong
    dashLongHeavy dotDash dotDashHeavy dotDotDash dotDotDashHeavy wavy
    wavyHeavy wavyDbl
    """.split()
)
_CHART_TEXT_STRIKE_VALUES = frozenset(
    {"noStrike", "sngStrike", "dblStrike"}
)
_CHART_TEXT_CAPS_VALUES = frozenset({"none", "small", "all"})
_CHART_TEXT_VERTICAL_OVERFLOW_VALUES = frozenset(
    {"overflow", "ellipsis", "clip"}
)
_CHART_TEXT_HORIZONTAL_OVERFLOW_VALUES = frozenset(
    {"overflow", "clip"}
)
_CHART_TEXT_VERTICAL_VALUES = frozenset(
    "horz vert vert270 wordArtVert eaVert mongolianVert wordArtVertRtl".split()
)
_CHART_TEXT_WRAPPING_VALUES = frozenset({"none", "square"})
_CHART_TEXT_ANCHORING_VALUES = frozenset({"t", "ctr", "b"})
_CHART_PRESET_LINE_DASH_VALUES = frozenset(
    """
    solid dot dash lgDash dashDot lgDashDot lgDashDotDot sysDash sysDot
    sysDashDot sysDashDotDot
    """.split()
)
_CHART_LINE_CAP_VALUES = frozenset({"rnd", "sq", "flat"})
_CHART_COMPOUND_LINE_VALUES = frozenset(
    {"sng", "dbl", "thickThin", "thinThick", "tri"}
)
_CHART_PEN_ALIGNMENT_VALUES = frozenset({"ctr", "in"})
_CHART_TILE_FLIP_VALUES = frozenset({"none", "x", "y", "xy"})
_CHART_PATH_SHADE_VALUES = frozenset({"shape", "circle", "rect"})
_CHART_PRESET_PATTERN_VALUES = frozenset(
    """
    pct5 pct10 pct20 pct25 pct30 pct40 pct50 pct60 pct70 pct75 pct80 pct90
    horz vert ltHorz ltVert dkHorz dkVert narHorz narVert dashHorz dashVert
    cross dnDiag upDiag ltDnDiag ltUpDiag dkDnDiag dkUpDiag wdDnDiag
    wdUpDiag dashDnDiag dashUpDiag diagCross smCheck lgCheck smGrid lgGrid
    dotGrid smConfetti lgConfetti horzBrick diagBrick solidDmnd openDmnd
    dotDmnd plaid sphere weave divot shingle wave trellis zigZag
    """.split()
)
_CHART_LINE_END_TYPE_VALUES = frozenset(
    {"none", "triangle", "stealth", "diamond", "oval", "arrow"}
)
_CHART_LINE_END_SIZE_VALUES = frozenset({"sm", "med", "lg"})
_CHART_TEXT_SHAPE_VALUES = frozenset(
    """
    textNoShape textPlain textStop textTriangle textTriangleInverted
    textChevron textChevronInverted textRingInside textRingOutside textArchUp
    textArchDown textCircle textButton textArchUpPour textArchDownPour
    textCirclePour textButtonPour textCurveUp textCurveDown textCanUp
    textCanDown textWave1 textWave2 textDoubleWave1 textWave4 textInflate
    textDeflate textInflateBottom textDeflateBottom textInflateTop
    textDeflateTop textDeflateInflate textDeflateInflateDeflate textFadeRight
    textFadeLeft textFadeUp textFadeDown textSlantUp textSlantDown
    textCascadeUp textCascadeDown
    """.split()
)
_CHART_SHAPE_TYPE_VALUES = frozenset(
    """
    line lineInv triangle rtTriangle rect diamond parallelogram trapezoid
    nonIsoscelesTrapezoid pentagon hexagon heptagon octagon decagon dodecagon
    star4 star5 star6 star7 star8 star10 star12 star16 star24 star32 roundRect
    round1Rect round2SameRect round2DiagRect snipRoundRect snip1Rect
    snip2SameRect snip2DiagRect plaque ellipse teardrop homePlate chevron
    pieWedge pie blockArc donut noSmoking rightArrow leftArrow upArrow
    downArrow stripedRightArrow notchedRightArrow bentUpArrow leftRightArrow
    upDownArrow leftUpArrow leftRightUpArrow quadArrow leftArrowCallout
    rightArrowCallout upArrowCallout downArrowCallout leftRightArrowCallout
    upDownArrowCallout quadArrowCallout bentArrow uturnArrow circularArrow
    leftCircularArrow leftRightCircularArrow curvedRightArrow curvedLeftArrow
    curvedUpArrow curvedDownArrow swooshArrow cube can lightningBolt heart sun
    moon smileyFace irregularSeal1 irregularSeal2 foldedCorner bevel frame
    halfFrame corner diagStripe chord arc leftBracket rightBracket leftBrace
    rightBrace bracketPair bracePair straightConnector1 bentConnector2
    bentConnector3 bentConnector4 bentConnector5 curvedConnector2
    curvedConnector3 curvedConnector4 curvedConnector5 callout1 callout2
    callout3 accentCallout1 accentCallout2 accentCallout3 borderCallout1
    borderCallout2 borderCallout3 accentBorderCallout1 accentBorderCallout2
    accentBorderCallout3 wedgeRectCallout wedgeRoundRectCallout
    wedgeEllipseCallout cloudCallout cloud ribbon ribbon2 ellipseRibbon
    ellipseRibbon2 leftRightRibbon verticalScroll horizontalScroll wave
    doubleWave plus flowChartProcess flowChartDecision flowChartInputOutput
    flowChartPredefinedProcess flowChartInternalStorage flowChartDocument
    flowChartMultidocument flowChartTerminator flowChartPreparation
    flowChartManualInput flowChartManualOperation flowChartConnector
    flowChartPunchedCard flowChartPunchedTape flowChartSummingJunction
    flowChartOr flowChartCollate flowChartSort flowChartExtract flowChartMerge
    flowChartOfflineStorage flowChartOnlineStorage flowChartMagneticTape
    flowChartMagneticDisk flowChartMagneticDrum flowChartDisplay flowChartDelay
    flowChartAlternateProcess flowChartOffpageConnector actionButtonBlank
    actionButtonHome actionButtonHelp actionButtonInformation
    actionButtonForwardNext actionButtonBackPrevious actionButtonEnd
    actionButtonBeginning actionButtonReturn actionButtonDocument
    actionButtonSound actionButtonMovie gear6 gear9 funnel mathPlus mathMinus
    mathMultiply mathDivide mathEqual mathNotEqual cornerTabs squareTabs
    plaqueTabs chartX chartStar chartPlus
    """.split()
)
_CHART_RECTANGLE_ALIGNMENT_VALUES = frozenset(
    {"tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"}
)
_CHART_BLEND_MODE_VALUES = frozenset(
    {"over", "mult", "screen", "darken", "lighten"}
)
_CHART_PRESET_SHADOW_VALUES = frozenset(
    {f"shdw{index}" for index in range(1, 21)}
)
_CHART_BLIP_COMPRESSION_VALUES = frozenset(
    {"email", "screen", "print", "hqprint", "none"}
)
_CHART_BLIP_EFFECT_NAMES = frozenset(
    {
        "alphaBiLevel", "alphaCeiling", "alphaFloor", "alphaInv",
        "alphaMod", "alphaModFix", "alphaRepl", "biLevel", "blur",
        "clrChange", "clrRepl", "duotone", "fillOverlay", "grayscl",
        "hsl", "lum", "tint",
    }
)
_CHART_PRESET_CAMERA_VALUES = frozenset(
    """
    legacyObliqueTopLeft legacyObliqueTop legacyObliqueTopRight
    legacyObliqueLeft legacyObliqueFront legacyObliqueRight
    legacyObliqueBottomLeft legacyObliqueBottom legacyObliqueBottomRight
    legacyPerspectiveTopLeft legacyPerspectiveTop legacyPerspectiveTopRight
    legacyPerspectiveLeft legacyPerspectiveFront legacyPerspectiveRight
    legacyPerspectiveBottomLeft legacyPerspectiveBottom
    legacyPerspectiveBottomRight orthographicFront isometricTopUp
    isometricTopDown isometricBottomUp isometricBottomDown isometricLeftUp
    isometricLeftDown isometricRightUp isometricRightDown isometricOffAxis1Left
    isometricOffAxis1Right isometricOffAxis1Top isometricOffAxis2Left
    isometricOffAxis2Right isometricOffAxis2Top isometricOffAxis3Left
    isometricOffAxis3Right isometricOffAxis3Bottom isometricOffAxis4Left
    isometricOffAxis4Right isometricOffAxis4Bottom obliqueTopLeft obliqueTop
    obliqueTopRight obliqueLeft obliqueRight obliqueBottomLeft obliqueBottom
    obliqueBottomRight perspectiveFront perspectiveLeft perspectiveRight
    perspectiveAbove perspectiveBelow perspectiveAboveLeftFacing
    perspectiveAboveRightFacing perspectiveContrastingLeftFacing
    perspectiveContrastingRightFacing perspectiveHeroicLeftFacing
    perspectiveHeroicRightFacing perspectiveHeroicExtremeLeftFacing
    perspectiveHeroicExtremeRightFacing perspectiveRelaxed
    perspectiveRelaxedModerately
    """.split()
)
_CHART_LIGHT_RIG_VALUES = frozenset(
    """
    legacyFlat1 legacyFlat2 legacyFlat3 legacyFlat4 legacyNormal1
    legacyNormal2 legacyNormal3 legacyNormal4 legacyHarsh1 legacyHarsh2
    legacyHarsh3 legacyHarsh4 threePt balanced soft harsh flood contrasting
    morning sunrise sunset chilly freezing flat twoPt glow brightRoom
    """.split()
)
_CHART_LIGHT_RIG_DIRECTION_VALUES = frozenset(
    {"tl", "t", "tr", "l", "r", "bl", "b", "br"}
)
_CHART_BEVEL_PRESET_VALUES = frozenset(
    """
    relaxedInset circle slope cross angle softRound convex coolSlant divot
    riblet hardEdge artDeco
    """.split()
)
_CHART_MATERIAL_VALUES = frozenset(
    """
    legacyMatte legacyPlastic legacyMetal legacyWireframe matte plastic metal
    warmMatte translucentPowder powder dkEdge softEdge clear flat softmetal
    """.split()
)
_CHART_PATH_FILL_VALUES = frozenset(
    {"none", "norm", "lighten", "lightenLess", "darken", "darkenLess"}
)
_XSD_BOOLEAN_VALUES = frozenset({"true", "false", "1", "0"})
_CHART_ROOT_TAG = f"{{{_CHART_NAMESPACE}}}chartSpace"
_CHARTEX_ROOT_TAG = f"{{{_CHARTEX_NAMESPACE}}}chartSpace"
_CHART_COLOR_MAPPING_TAG = f"{{{_CHART_NAMESPACE}}}clrMapOvr"
_CHARTEX_COLOR_MAPPING_TAG = f"{{{_CHARTEX_NAMESPACE}}}clrMapOvr"
_CHART_COLOR_MAPPING_ATTRIBUTES = (
    "bg1", "tx1", "bg2", "tx2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
)
_CHART_COLOR_MAPPING_VALUES = frozenset(_THEME_COLOR_CHILDREN)


def _theme_color_model_value(color_element, label: str) -> str:
    """读取主题色的实际 RGB 值；对无法无损解释的模型 fail-closed。"""
    models = [
        child
        for child in color_element
        if child.tag in {_THEME_SRGB_COLOR_TAG, _THEME_SYS_COLOR_TAG}
    ]
    if len(models) != 1:
        raise ValueError(f"{label} uses an unsupported theme color model")
    model = models[0]
    if len(model):
        raise ValueError(f"{label} has unsupported theme color transforms")
    value = (
        model.get("val")
        if model.tag == _THEME_SRGB_COLOR_TAG
        else model.get("lastClr") or model.get("val")
    )
    if not isinstance(value, str) or not _HEX_COLOR_RE.fullmatch(value):
        raise ValueError(f"{label} has an invalid RGB value")
    return value.upper()


def _apply_wml_theme_color_transform(
    rgb: str,
    tint_hex: str | None,
    shade_hex: str | None,
) -> str:
    """按 Word 使用的 0..240 HLS 规则应用 themeTint/themeShade。"""
    if tint_hex is not None and shade_hex is not None:
        raise ValueError("theme tint and shade cannot both be present")
    transform = tint_hex if tint_hex is not None else shade_hex
    if transform is None:
        return rgb
    if not _HEX_BYTE_RE.fullmatch(transform):
        raise ValueError("theme tint/shade must be a one-byte hexadecimal value")

    red, green, blue = (
        int(rgb[index : index + 2], 16) / 255.0
        for index in (0, 2, 4)
    )
    hue, luminance, saturation = colorsys.rgb_to_hls(red, green, blue)
    hls_max = 240

    def round_half_up(value: float) -> int:
        return math.floor(value + 0.5)

    hue = round_half_up(hue * hls_max)
    luminance = round_half_up(luminance * hls_max)
    saturation = round_half_up(saturation * hls_max)
    remaining = int(transform, 16) / 255.0
    if tint_hex is not None:
        luminance = round_half_up(
            luminance * remaining + hls_max * (1.0 - remaining)
        )
    else:
        luminance = round_half_up(luminance * remaining)

    transformed = colorsys.hls_to_rgb(
        hue / hls_max,
        luminance / hls_max,
        saturation / hls_max,
    )
    channels = tuple(
        min(255, max(0, math.floor(channel * 255.0 + 1e-9)))
        for channel in transformed
    )
    return "".join(f"{channel:02X}" for channel in channels)


def _theme_language_script(language: str | None) -> str | None:
    """将 Word 的 BCP-47 语言标记映射到 DrawingML supplemental script。"""
    if not language:
        return None
    normalized = language.replace("_", "-")
    parts = normalized.split("-")
    script = next(
        (part.title() for part in parts[1:] if len(part) == 4 and part.isalpha()),
        None,
    )
    if script is None:
        try:
            from babel import Locale
            from babel.core import get_global

            locale = Locale.parse(normalized, sep="-")
            script = locale.script
            if script is None:
                language_key = parts[0].lower()
                territory = next(
                    (part.upper() for part in parts[1:] if len(part) == 2),
                    None,
                )
                likely = get_global("likely_subtags")
                likely_value = (
                    likely.get(f"{language_key}_{territory}")
                    if territory
                    else None
                ) or likely.get(language_key)
                if likely_value:
                    script = next(
                        (
                            part.title()
                            for part in likely_value.split("_")[1:]
                            if len(part) == 4
                        ),
                        None,
                    )
        except Exception:
            script = None

    if script == "Kore":
        return "Hang"
    return script


def _normalize_theme_scheme_name(value: str | None) -> str | None:
    if not isinstance(value, str):
        return None
    return _THEME_COLOR_ALIAS.get(value.casefold(), value)


@dataclass(frozen=True)
class _DocumentThemeSemantics:
    colors: dict[str, str]
    fonts: dict[str, dict[str, object]]
    drawing_mapping: dict[str, str]
    theme_languages: dict[str, str]
    color_signature: tuple
    font_signature: tuple
    format_signature: tuple
    signature: tuple

    def font_reference_signature(self, token: str) -> tuple:
        """返回无法按语言解析时可用于判断两主题是否等价的字体签名。"""
        match = _THEME_FONT_TOKEN_RE.fullmatch(token.strip())
        if match is None:
            raise ValueError("inserted document contains an invalid theme font reference")
        family = match.group("major").casefold()
        slot = match.group("slot").casefold()
        family_data = self.fonts[family]
        if slot in {"ascii", "hansi"}:
            base_slot = "latin"
            language_key = "val"
        elif slot == "eastasia":
            base_slot = "ea"
            language_key = "eastAsia"
        else:
            base_slot = "cs"
            language_key = "bidi"
        return (
            family,
            base_slot,
            family_data[base_slot],
            tuple(sorted(family_data["supplemental"].items())),
            self.theme_languages.get(language_key),
        )

    def resolve_font(self, token: str, rfonts_element) -> str:
        match = _THEME_FONT_TOKEN_RE.fullmatch(token.strip())
        if match is None:
            raise ValueError("inserted document contains an invalid theme font reference")
        family = match.group("major").casefold()
        slot = match.group("slot").casefold()
        family_data = self.fonts[family]
        if slot in {"ascii", "hansi"}:
            typeface = family_data["latin"]
            language_key = "val"
        elif slot == "eastasia":
            typeface = family_data["ea"]
            language_key = "eastAsia"
        else:
            typeface = family_data["cs"]
            language_key = "bidi"

        if not typeface:
            language = None
            parent = rfonts_element.getparent()
            if parent is not None:
                language_element = parent.find(qn("w:lang"))
                if language_element is not None:
                    language = language_element.get(qn(f"w:{language_key}"))
            language = language or self.theme_languages.get(language_key)
            script = _theme_language_script(language)
            if script == "Latn":
                typeface = family_data["latin"]
            elif script:
                typeface = family_data["supplemental"].get(script)

        if not isinstance(typeface, str) or not typeface:
            raise ValueError(
                "inserted document theme font cannot be resolved for the requested language"
            )
        return typeface

    def resolve_wml_color(self, token: str) -> str:
        key = _normalize_theme_scheme_name(token)
        if key not in self.colors:
            raise ValueError("inserted document contains an invalid theme color reference")
        return self.colors[key]

    def resolve_drawing_color(self, token: str) -> str:
        folded = token.casefold()
        if folded == "phclr":
            raise ValueError(
                "inserted DrawingML uses phClr, which has no document-independent color"
            )
        logical = _DRAWING_THEME_LOGICAL_ALIASES.get(folded)
        if logical is None:
            direct = _normalize_theme_scheme_name(token)
            if direct not in self.colors:
                raise ValueError("inserted DrawingML contains an invalid scheme color")
            return self.colors[direct]
        mapped = self.drawing_mapping.get(logical)
        if mapped not in self.colors:
            raise ValueError("inserted document has an invalid color-scheme mapping")
        return self.colors[mapped]


@dataclass(frozen=True)
class _ThemeMaterializationContext:
    source: _DocumentThemeSemantics | None
    target: _DocumentThemeSemantics | None
    relationships_differ: bool = False

    def materialize(self, root) -> None:
        if self.source is None:
            has_font_reference = any(
                node.get(theme_attribute) is not None
                for node in root.iter()
                for theme_attribute, _direct_attribute, _slot
                in _THEME_FONT_REFERENCE_SPECS
            )
            has_wml_color_reference = any(
                node.get(attribute_name) is not None
                for node in root.iter()
                for attribute_name in (qn("w:themeColor"), qn("w:themeFill"))
            )
            has_drawing_color_reference = (
                next(root.iter(_THEME_SCHEME_COLOR_TAG), None) is not None
            )
            has_contextual_drawing_reference = any(
                node.tag in _THEME_FORMAT_REFERENCE_TAGS
                or node.tag == _THEME_FONT_REFERENCE_TAG
                for node in root.iter()
            )
            has_reference = (
                has_font_reference
                or has_wml_color_reference
                or has_drawing_color_reference
                or has_contextual_drawing_reference
            )
            if has_reference:
                raise ValueError(
                    "inserted document contains theme references but no theme Part"
                )
            return

        if (
            self.target is None
            or self.source.format_signature != self.target.format_signature
            or self.relationships_differ
        ) and any(
            node.tag in _THEME_FORMAT_REFERENCE_TAGS
            for node in root.iter()
        ):
            raise ValueError(
                "inserted DrawingML format references require a different theme format scheme"
            )
        if (
            self.target is None
            or self.source.font_signature != self.target.font_signature
        ) and next(root.iter(_THEME_FONT_REFERENCE_TAG), None) is not None:
            raise ValueError(
                "inserted DrawingML fontRef requires a different theme font scheme"
            )
        if (
            self.target is None
            or self.source.color_signature != self.target.color_signature
        ) and any(
            node.tag in _THEME_COLOR_MAPPING_OVERRIDE_TAGS
            for node in root.iter()
        ):
            raise ValueError(
                "inserted DrawingML uses a local color mapping that cannot be flattened"
            )

        for rfonts in root.iter(qn("w:rFonts")):
            for theme_attribute, direct_attribute, _slot in (
                _THEME_FONT_REFERENCE_SPECS
            ):
                token = rfonts.get(theme_attribute)
                if token is None:
                    continue
                try:
                    source_typeface = self.source.resolve_font(token, rfonts)
                except ValueError:
                    if (
                        self.target is not None
                        and self.source.font_reference_signature(token)
                        == self.target.font_reference_signature(token)
                    ):
                        continue
                    raise
                if self.target is not None:
                    try:
                        target_typeface = self.target.resolve_font(token, rfonts)
                    except ValueError:
                        target_typeface = None
                    if target_typeface == source_typeface:
                        continue
                rfonts.set(direct_attribute, source_typeface)
                del rfonts.attrib[theme_attribute]

        for node in root.iter():
            theme_color = node.get(qn("w:themeColor"))
            if theme_color is not None:
                fallback_attribute = (
                    qn("w:val") if node.tag == qn("w:color") else qn("w:color")
                )
                if theme_color.casefold() == "none":
                    continue
                else:
                    resolved = _apply_wml_theme_color_transform(
                        self.source.resolve_wml_color(theme_color),
                        node.get(qn("w:themeTint")),
                        node.get(qn("w:themeShade")),
                    )
                    if self.target is not None:
                        target_resolved = _apply_wml_theme_color_transform(
                            self.target.resolve_wml_color(theme_color),
                            node.get(qn("w:themeTint")),
                            node.get(qn("w:themeShade")),
                        )
                        if target_resolved == resolved:
                            continue
                node.set(fallback_attribute, resolved)
                del node.attrib[qn("w:themeColor")]
                node.attrib.pop(qn("w:themeTint"), None)
                node.attrib.pop(qn("w:themeShade"), None)

            theme_fill = node.get(qn("w:themeFill"))
            if theme_fill is not None:
                if theme_fill.casefold() == "none":
                    continue
                else:
                    resolved = _apply_wml_theme_color_transform(
                        self.source.resolve_wml_color(theme_fill),
                        node.get(qn("w:themeFillTint")),
                        node.get(qn("w:themeFillShade")),
                    )
                    if self.target is not None:
                        target_resolved = _apply_wml_theme_color_transform(
                            self.target.resolve_wml_color(theme_fill),
                            node.get(qn("w:themeFillTint")),
                            node.get(qn("w:themeFillShade")),
                        )
                        if target_resolved == resolved:
                            continue
                node.set(qn("w:fill"), resolved)
                del node.attrib[qn("w:themeFill")]
                node.attrib.pop(qn("w:themeFillTint"), None)
                node.attrib.pop(qn("w:themeFillShade"), None)

        for node in root.iter(_THEME_SCHEME_COLOR_TAG):
            value = node.get("val")
            if not isinstance(value, str):
                raise ValueError("inserted DrawingML scheme color has no value")
            if value.casefold() == "phclr":
                if (
                    self.target is not None
                    and self.source.color_signature
                    == self.target.color_signature
                    and self.source.format_signature
                    == self.target.format_signature
                    and not self.relationships_differ
                ):
                    continue
                raise ValueError(
                    "inserted DrawingML uses phClr, which has no document-independent color"
                )
            resolved = self.source.resolve_drawing_color(value)
            if self.target is not None:
                target_resolved = self.target.resolve_drawing_color(value)
                if target_resolved == resolved:
                    continue
            node.tag = _THEME_SRGB_COLOR_TAG
            node.set("val", resolved)


def _document_theme_part(document):
    relationships = [
        relationship
        for relationship in document.part.rels.values()
        if relationship.reltype == RT.THEME
    ]
    if not relationships:
        return None
    if len(relationships) != 1 or relationships[0].is_external:
        raise ValueError("document has an invalid theme relationship")
    theme_part = relationships[0].target_part
    if str(theme_part.content_type).casefold() != str(CT.OFC_THEME).casefold():
        raise ValueError("document theme has the wrong content type")
    return theme_part


MAX_THEME_RELATIONSHIP_GRAPH_NODES = 4_096
MAX_THEME_RELATIONSHIP_GRAPH_EDGES = 8_192
MAX_THEME_RELATIONSHIP_GRAPH_EXPANSIONS = 16_384


class _ThemeRelationshipSignatureState:
    def __init__(self):
        self.signatures: dict[int, tuple] = {}
        self.relationships: dict[int, tuple] = {}
        self.part_metadata: dict[int, tuple[str, str]] = {}
        self.edge_count = 0
        self.expansion_count = 0

    def relationships_for(self, part) -> tuple:
        part_key = id(part)
        cached = self.relationships.get(part_key)
        if cached is not None:
            return cached
        if len(self.relationships) >= MAX_THEME_RELATIONSHIP_GRAPH_NODES:
            raise ValueError(
                "document theme relationship graph has too many parts"
            )

        relationships = tuple(part.rels.values())
        edge_count = self.edge_count + len(relationships)
        if edge_count > MAX_THEME_RELATIONSHIP_GRAPH_EDGES:
            raise ValueError(
                "document theme relationship graph has too many relationships"
            )
        self.relationships[part_key] = relationships
        self.edge_count = edge_count
        return relationships

    def metadata_for(self, part) -> tuple[str, str]:
        part_key = id(part)
        cached = self.part_metadata.get(part_key)
        if cached is not None:
            return cached
        metadata = (
            str(part.content_type).casefold(),
            hashlib.sha256(part.blob).hexdigest(),
        )
        self.part_metadata[part_key] = metadata
        return metadata


def _walk_theme_relationship_signature(part, stack, state):
    part_key = id(part)
    if part_key in stack:
        return ("cycle", len(stack) - stack.index(part_key)), True
    if len(stack) >= 64:
        raise ValueError("document theme relationship graph is too deep")

    cached = state.signatures.get(part_key)
    if cached is not None:
        return cached, False
    state.expansion_count += 1
    if state.expansion_count > MAX_THEME_RELATIONSHIP_GRAPH_EXPANSIONS:
        raise ValueError(
            "document theme relationship graph requires too many expansions"
        )

    relationships = []
    depends_on_stack = False
    for relationship in state.relationships_for(part):
        if relationship.is_external:
            relationships.append(
                (
                    relationship.rId,
                    relationship.reltype,
                    "external",
                    relationship.target_ref,
                )
            )
            continue
        target = relationship.target_part
        target_content_type, target_digest = state.metadata_for(target)
        target_signature, target_depends_on_stack = (
            _walk_theme_relationship_signature(
                target,
                (*stack, part_key),
                state,
            )
        )
        depends_on_stack = depends_on_stack or target_depends_on_stack
        relationships.append(
            (
                relationship.rId,
                relationship.reltype,
                "internal",
                target_content_type,
                target_digest,
                target_signature,
            )
        )
    signature = tuple(sorted(relationships))
    # A cycle marker encodes its distance in the current recursion stack. Such
    # signatures remain path-dependent and must not be reused from another path.
    if not depends_on_stack:
        state.signatures[part_key] = signature
    return signature, depends_on_stack


def _theme_relationship_signature(part, stack=()):
    signature, _ = _walk_theme_relationship_signature(
        part,
        stack,
        _ThemeRelationshipSignatureState(),
    )
    return signature


def _theme_settings_signature(document) -> tuple:
    return tuple(
        tuple(sorted(element.attrib.items())) if element is not None else ()
        for element in (
            document.settings.element.find(qn("w:clrSchemeMapping")),
            document.settings.element.find(qn("w:themeFontLang")),
        )
    )


def _document_drawing_mapping(document) -> dict[str, str]:
    drawing_mapping = dict(_DRAWING_THEME_DEFAULT_MAPPING)
    mapping_element = document.settings.element.find(qn("w:clrSchemeMapping"))
    if mapping_element is None:
        return drawing_mapping
    for attribute_name, value in mapping_element.attrib.items():
        local_name = etree.QName(attribute_name).localname
        logical = _DRAWING_THEME_LOGICAL_ALIASES.get(local_name.casefold())
        normalized_value = _normalize_theme_scheme_name(value)
        if (
            logical is None
            or normalized_value not in _CHART_COLOR_MAPPING_VALUES
        ):
            raise ValueError("document has an invalid clrSchemeMapping")
        drawing_mapping[logical] = normalized_value
    return drawing_mapping


def _theme_scheme_elements(theme_part) -> dict[str, object]:
    root = _parse_untrusted_ooxml_part(theme_part.blob)
    if root.tag != _THEME_TAG:
        raise ValueError("document theme has the wrong root")
    theme_elements = root.find(f"{{{_THEME_NAMESPACE}}}themeElements")
    if theme_elements is None:
        raise ValueError("document theme has no themeElements")
    schemes = {}
    for scheme_tag in (
        _THEME_COLOR_SCHEME_TAG,
        _THEME_FONT_SCHEME_TAG,
        _THEME_FORMAT_SCHEME_TAG,
    ):
        matches = theme_elements.findall(scheme_tag)
        if len(matches) != 1:
            raise ValueError(
                "document theme has an invalid color/font/format scheme"
            )
        schemes[scheme_tag] = matches[0]
    return schemes


def _theme_part_semantics(document) -> _DocumentThemeSemantics | None:
    theme_part = _document_theme_part(document)
    if theme_part is None:
        return None
    schemes = _theme_scheme_elements(theme_part)

    colors = {}
    color_scheme = schemes[_THEME_COLOR_SCHEME_TAG]
    for name in _THEME_COLOR_CHILDREN:
        matches = color_scheme.findall(f"{{{_THEME_NAMESPACE}}}{name}")
        if len(matches) != 1:
            raise ValueError("document theme has missing or duplicate color entries")
        colors[name] = _theme_color_model_value(matches[0], f"theme color {name}")

    fonts = {}
    font_scheme = schemes[_THEME_FONT_SCHEME_TAG]
    for family in ("major", "minor"):
        matches = font_scheme.findall(f"{{{_THEME_NAMESPACE}}}{family}Font")
        if len(matches) != 1:
            raise ValueError("document theme has missing or duplicate font entries")
        family_element = matches[0]
        family_data = {}
        for slot, tag in _THEME_FONT_SLOT_TAGS.items():
            slot_elements = family_element.findall(tag)
            if len(slot_elements) != 1:
                raise ValueError("document theme has missing or duplicate font slots")
            family_data[slot] = slot_elements[0].get("typeface") or ""
        supplemental = {}
        for supplemental_font in family_element.findall(_THEME_FONT_TAG):
            script = supplemental_font.get("script")
            typeface = supplemental_font.get("typeface") or ""
            if not script or script in supplemental:
                raise ValueError("document theme has invalid supplemental fonts")
            supplemental[script] = typeface
        family_data["supplemental"] = supplemental
        fonts[family] = family_data

    drawing_mapping = _document_drawing_mapping(document)

    language_element = document.settings.element.find(qn("w:themeFontLang"))
    theme_languages = {}
    if language_element is not None:
        for key, attribute_name in (
            ("val", qn("w:val")),
            ("eastAsia", qn("w:eastAsia")),
            ("bidi", qn("w:bidi")),
        ):
            value = language_element.get(attribute_name)
            if value:
                theme_languages[key] = value

    font_signature = tuple(
        (
            family,
            tuple(
                sorted(
                    (
                        key,
                        value
                        if key != "supplemental"
                        else tuple(sorted(value.items())),
                    )
                    for key, value in data.items()
                )
            ),
        )
        for family, data in sorted(fonts.items())
    )
    format_signature = _xml_value_signature(
        schemes[_THEME_FORMAT_SCHEME_TAG]
    )
    color_signature = (
        tuple(sorted(colors.items())),
        tuple(sorted(drawing_mapping.items())),
    )
    signature = (
        color_signature,
        font_signature,
        format_signature,
        tuple(sorted(theme_languages.items())),
    )
    return _DocumentThemeSemantics(
        colors=colors,
        fonts=fonts,
        drawing_mapping=drawing_mapping,
        theme_languages=theme_languages,
        color_signature=color_signature,
        font_signature=font_signature,
        format_signature=format_signature,
        signature=signature,
    )


def _theme_materialization_context(source_doc, target_doc):
    source_part = _document_theme_part(source_doc)
    target_part = _document_theme_part(target_doc)
    source_relationship_signature = (
        _theme_relationship_signature(source_part)
        if source_part is not None
        else None
    )
    target_relationship_signature = (
        _theme_relationship_signature(target_part)
        if target_part is not None
        else None
    )
    relationships_differ = (
        source_relationship_signature != target_relationship_signature
    )
    if source_part is None and target_part is None:
        return None
    if (
        source_part is not None
        and target_part is not None
        and source_part.blob == target_part.blob
        and _theme_settings_signature(source_doc)
        == _theme_settings_signature(target_doc)
        and not relationships_differ
    ):
        return None
    source_theme = _theme_part_semantics(source_doc)
    target_theme = _theme_part_semantics(target_doc)
    if (
        source_theme is not None
        and target_theme is not None
        and source_theme.signature == target_theme.signature
        and not relationships_differ
    ):
        return None
    return _ThemeMaterializationContext(
        source_theme,
        target_theme,
        relationships_differ=relationships_differ,
    )


@dataclass(frozen=True)
class _ChartThemeContext:
    source_theme_part: object | None
    source_schemes: dict[str, object] | None
    source_mapping: dict[str, str]
    target_mapping: dict[str, str]
    schemes_differ: bool
    mapping_differ: bool


def _chart_theme_context(
    source_doc,
    target_doc,
    materialization_context,
) -> _ChartThemeContext:
    source_mapping = _document_drawing_mapping(source_doc)
    target_mapping = _document_drawing_mapping(target_doc)
    source_document_theme_part = _document_theme_part(source_doc)
    target_document_theme_part = _document_theme_part(target_doc)
    theme_relationships_differ = (
        (source_document_theme_part is None)
        != (target_document_theme_part is None)
        or (
            source_document_theme_part is not None
            and target_document_theme_part is not None
            and _theme_relationship_signature(source_document_theme_part)
            != _theme_relationship_signature(target_document_theme_part)
        )
    )
    schemes_differ = False
    source_theme = None
    target_theme = None
    if materialization_context is not None:
        source_theme = materialization_context.source
        target_theme = materialization_context.target
        if source_theme is None or target_theme is None:
            schemes_differ = source_theme is not target_theme
        else:
            schemes_differ = (
                source_theme.colors != target_theme.colors
                or source_theme.font_signature != target_theme.font_signature
                or source_theme.format_signature != target_theme.format_signature
            )
    schemes_differ = schemes_differ or theme_relationships_differ

    source_theme_part = (
        source_document_theme_part
        if schemes_differ and source_theme is not None
        else None
    )
    source_schemes = (
        _theme_scheme_elements(source_theme_part)
        if source_theme_part is not None
        else None
    )
    return _ChartThemeContext(
        source_theme_part=source_theme_part,
        source_schemes=source_schemes,
        source_mapping=source_mapping,
        target_mapping=target_mapping,
        schemes_differ=schemes_differ,
        mapping_differ=source_mapping != target_mapping,
    )


def _theme_scheme_children(element, required_tags, label: str):
    """Return required children after rejecting missing, duplicate, or reordered nodes."""
    children = list(element)
    if children and children[-1].tag == _THEME_EXTENSION_LIST_TAG:
        children = children[:-1]
    if [child.tag for child in children] != list(required_tags):
        raise ValueError(f"chart theme override has an invalid {label} structure")
    return children


def _validate_drawing_color_model(model, label: str) -> None:
    if model.tag not in _THEME_COLOR_MODEL_TAGS:
        raise ValueError(f"{label} has an unsupported color model")
    required_attributes = {
            f"{{{_THEME_NAMESPACE}}}scrgbClr": ("r", "g", "b"),
            f"{{{_THEME_NAMESPACE}}}srgbClr": ("val",),
            f"{{{_THEME_NAMESPACE}}}hslClr": ("hue", "sat", "lum"),
            f"{{{_THEME_NAMESPACE}}}sysClr": ("val",),
            f"{{{_THEME_NAMESPACE}}}schemeClr": ("val",),
            f"{{{_THEME_NAMESPACE}}}prstClr": ("val",),
    }[model.tag]
    if any(
        not isinstance(model.get(attribute_name), str)
        or not model.get(attribute_name)
        for attribute_name in required_attributes
    ):
        raise ValueError(f"{label} has missing color attributes")
    if (
        model.tag == f"{{{_THEME_NAMESPACE}}}srgbClr"
        and not _HEX_COLOR_RE.fullmatch(model.get("val"))
    ):
        raise ValueError(f"{label} has an invalid RGB value")
    if model.tag == f"{{{_THEME_NAMESPACE}}}sysClr":
        last_color = model.get("lastClr")
        if last_color is not None and not _HEX_COLOR_RE.fullmatch(last_color):
            raise ValueError(
                f"{label} has an invalid system-color fallback"
            )


def _validate_theme_override_color_scheme(scheme) -> None:
    if scheme.get("name") is None:
        raise ValueError("chart theme override color scheme has no name")
    colors = _theme_scheme_children(
        scheme,
        tuple(
            f"{{{_THEME_NAMESPACE}}}{name}"
            for name in _THEME_COLOR_CHILDREN
        ),
        "color scheme",
    )
    for color in colors:
        models = [
            child for child in color if child.tag in _THEME_COLOR_MODEL_TAGS
        ]
        if len(color) != 1 or len(models) != 1:
            raise ValueError(
                "chart theme override color entry has an invalid color model"
            )
        _validate_drawing_color_model(
            models[0],
            "chart theme override color model",
        )


def _validate_theme_override_font_scheme(scheme) -> None:
    if scheme.get("name") is None:
        raise ValueError("chart theme override font scheme has no name")
    families = _theme_scheme_children(
        scheme,
        (
            f"{{{_THEME_NAMESPACE}}}majorFont",
            f"{{{_THEME_NAMESPACE}}}minorFont",
        ),
        "font scheme",
    )
    required_slots = tuple(_THEME_FONT_SLOT_TAGS.values())
    font_tag = _THEME_FONT_TAG
    for family in families:
        children = list(family)
        if children and children[-1].tag == _THEME_EXTENSION_LIST_TAG:
            children = children[:-1]
        if (
            len(children) < len(required_slots)
            or tuple(child.tag for child in children[:3]) != required_slots
            or any(child.tag != font_tag for child in children[3:])
        ):
            raise ValueError(
                "chart theme override has an invalid font collection"
            )
        if any(slot.get("typeface") is None for slot in children[:3]):
            raise ValueError(
                "chart theme override font collection has a missing typeface"
            )
        scripts = []
        for supplemental_font in children[3:]:
            script = supplemental_font.get("script")
            typeface = supplemental_font.get("typeface")
            if not script or typeface is None:
                raise ValueError(
                    "chart theme override has an invalid supplemental font"
                )
            scripts.append(script)
        if len(scripts) != len(set(scripts)):
            raise ValueError(
                "chart theme override has duplicate supplemental fonts"
            )


def _validate_theme_override_format_scheme(scheme) -> None:
    style_lists = _theme_scheme_children(
        scheme,
        tuple(tag for tag, _allowed_children in _THEME_FORMAT_SCHEME_LIST_SPECS),
        "format scheme",
    )
    for style_list, (_tag, allowed_children) in zip(
        style_lists,
        _THEME_FORMAT_SCHEME_LIST_SPECS,
    ):
        if len(style_list) < 3 or any(
            child.tag not in allowed_children for child in style_list
        ):
            raise ValueError(
                "chart theme override has an invalid format style list"
            )


def _validate_drawingml_relationship_attribute_shape(node, label: str) -> None:
    required_relationship_attribute = (
        _DRAWINGML_REQUIRED_RELATIONSHIP_ATTRIBUTES.get(node.tag)
    )
    if required_relationship_attribute is not None:
        relationship_attributes = {
            attribute_name
            for attribute_name in node.attrib
            if etree.QName(attribute_name).namespace
            == _RELATIONSHIP_NAMESPACE
        }
        if required_relationship_attribute not in relationship_attributes:
            raise ValueError(
                f"{label} has media or attribution without its required "
                "relationship attribute"
            )
        if relationship_attributes != {required_relationship_attribute}:
            raise ValueError(
                f"{label} has media or attribution with an invalid "
                "relationship attribute"
            )
    if (
        node.tag in _DRAWINGML_HYPERLINK_TAGS
        and _RELATIONSHIP_ID_ATTRIBUTE not in node.attrib
    ):
        raise ValueError(
            f"{label} has a hyperlink with a missing relationship ID"
        )


def _relationship_references(root, part, label: str) -> set[str]:
    """Validate relationship-namespace attributes and return their IDs."""
    referenced_relationships = set()
    for node in root.iter():
        _validate_drawingml_relationship_attribute_shape(node, label)
        for attribute_name, relationship_id in node.attrib.items():
            if etree.QName(attribute_name).namespace != _RELATIONSHIP_NAMESPACE:
                continue
            local_name = etree.QName(attribute_name).localname
            # DrawingML hyperlinks may intentionally use an explicitly empty
            # r:id, for example when an action supplies the target.
            if (
                not relationship_id
                and node.tag in _DRAWINGML_HYPERLINK_TAGS
                and local_name == "id"
            ):
                continue
            if not relationship_id or relationship_id not in part.rels:
                raise ValueError(f"{label} has a missing relationship")
            relationship = part.rels[relationship_id]
            if local_name == "embed" and relationship.is_external:
                raise ValueError(f"{label} has an external embedded relationship")
            image_relationship = (
                node.tag in _DRAWINGML_IMAGE_TAGS
                and local_name in {"embed", "link"}
            ) or (
                node.tag in _DRAWINGML_EMBEDDED_IMAGE_ONLY_TAGS
                and attribute_name == _RELATIONSHIP_EMBED_ATTRIBUTE
            )
            if image_relationship:
                if (
                    relationship.reltype != RT.IMAGE
                    or (
                        not relationship.is_external
                        and not str(
                            relationship.target_part.content_type
                        ).casefold().startswith("image/")
                    )
                ):
                    raise ValueError(
                        f"{label} has an image with the wrong target"
                    )

            if (
                node.tag
                == f"{{{_DRAWING_2017_MODEL3D_NAMESPACE}}}model3d"
                and local_name in {"embed", "link"}
                and (
                    relationship.reltype != _MODEL3D_RELATIONSHIP_TYPE
                    or (local_name == "embed" and relationship.is_external)
                    or (
                        not relationship.is_external
                        and str(
                            relationship.target_part.content_type
                        ).casefold()
                        != _MODEL3D_CONTENT_TYPE.casefold()
                    )
                )
            ):
                raise ValueError(
                    f"{label} has a 3D model with the wrong target"
                )

            media_spec = next(
                (
                    (relationship_type, must_be_internal)
                    for (
                        tag,
                        expected_attribute,
                        relationship_type,
                        must_be_internal,
                    ) in _DRAWINGML_MEDIA_RELATIONSHIP_SPECS
                    if node.tag == tag and attribute_name == expected_attribute
                ),
                None,
            )
            if media_spec is not None:
                expected_type, must_be_internal = media_spec
                if (
                    relationship.reltype != expected_type
                    or (must_be_internal and relationship.is_external)
                ):
                    raise ValueError(
                        f"{label} has media with the wrong target"
                    )

            if (
                node.tag in _DRAWINGML_HYPERLINK_TAGS
                and local_name == "id"
                and relationship.reltype != RT.HYPERLINK
            ):
                raise ValueError(
                    f"{label} has a hyperlink with the wrong target"
                )
            referenced_relationships.add(relationship_id)
    return referenced_relationships


def _validated_theme_override(part):
    if str(part.content_type).casefold() != str(
        CT.OFC_THEME_OVERRIDE
    ).casefold():
        raise ValueError("chart theme override has the wrong content type")
    root = _parse_untrusted_ooxml_part(part.blob)
    if root.tag != _THEME_OVERRIDE_TAG:
        raise ValueError("chart theme override has the wrong root")
    if root.attrib:
        raise ValueError("chart theme override has invalid root attributes")
    scheme_order = (
        _THEME_COLOR_SCHEME_TAG,
        _THEME_FONT_SCHEME_TAG,
        _THEME_FORMAT_SCHEME_TAG,
    )
    schemes = {}
    observed_indexes = []
    for child in root:
        if child.tag not in scheme_order or child.tag in schemes:
            raise ValueError("chart theme override has invalid scheme children")
        schemes[child.tag] = child
        observed_indexes.append(scheme_order.index(child.tag))
    if observed_indexes != sorted(observed_indexes):
        raise ValueError("chart theme override schemes are out of order")

    validators = {
        _THEME_COLOR_SCHEME_TAG: _validate_theme_override_color_scheme,
        _THEME_FONT_SCHEME_TAG: _validate_theme_override_font_scheme,
        _THEME_FORMAT_SCHEME_TAG: _validate_theme_override_format_scheme,
    }
    for scheme_tag, scheme in schemes.items():
        validators[scheme_tag](scheme)

    for relationship in part.rels.values():
        if relationship.is_external:
            raise ValueError("chart theme override has an external relationship")
    referenced_relationships = _relationship_references(
        root,
        part,
        "chart theme override",
    )
    if referenced_relationships != set(part.rels):
        raise ValueError("chart theme override has an unreferenced relationship")
    return root, schemes


def _validate_chart_color_mapping(mapping_element) -> None:
    if set(mapping_element.attrib) != set(
        _CHART_COLOR_MAPPING_ATTRIBUTES
    ):
        raise ValueError("chart color mapping has missing or unknown attributes")
    if any(
        value not in _CHART_COLOR_MAPPING_VALUES
        for value in mapping_element.attrib.values()
    ):
        raise ValueError("chart color mapping has an invalid scheme value")
    extension_lists = mapping_element.findall(
        f"{{{_THEME_NAMESPACE}}}extLst"
    )
    if len(extension_lists) != len(mapping_element) or len(extension_lists) > 1:
        raise ValueError("chart color mapping has invalid child elements")


def _validate_chart_structure(chart_root, relationship_type: str) -> None:
    if relationship_type == RT.CHART:
        chart_tag = f"{{{_CHART_NAMESPACE}}}chart"
        plot_area_tag = f"{{{_CHART_NAMESPACE}}}plotArea"
        chart_type_tags = frozenset(
            f"{{{_CHART_NAMESPACE}}}{name}"
            for name in (
                "areaChart",
                "area3DChart",
                "lineChart",
                "line3DChart",
                "stockChart",
                "radarChart",
                "scatterChart",
                "pieChart",
                "pie3DChart",
                "doughnutChart",
                "barChart",
                "bar3DChart",
                "ofPieChart",
                "surfaceChart",
                "surface3DChart",
                "bubbleChart",
            )
        )
        charts = chart_root.findall(chart_tag)
        if len(charts) != 1:
            raise ValueError("chart Part has no unique chart")
        plot_areas = charts[0].findall(plot_area_tag)
        if len(plot_areas) != 1 or not any(
            child.tag in chart_type_tags for child in plot_areas[0]
        ):
            raise ValueError("chart Part has an invalid plot area")
        return

    if relationship_type == _CHARTEX_RELATIONSHIP_TYPE:
        chart_data_tag = f"{{{_CHARTEX_NAMESPACE}}}chartData"
        chart_tag = f"{{{_CHARTEX_NAMESPACE}}}chart"
        plot_area_tag = f"{{{_CHARTEX_NAMESPACE}}}plotArea"
        plot_region_tag = f"{{{_CHARTEX_NAMESPACE}}}plotAreaRegion"
        chart_data = chart_root.findall(chart_data_tag)
        charts = chart_root.findall(chart_tag)
        if (
            len(chart_data) != 1
            or not len(chart_data[0])
            or len(charts) != 1
        ):
            raise ValueError("ChartEx Part has invalid chart data")
        plot_areas = charts[0].findall(plot_area_tag)
        plot_regions = (
            plot_areas[0].findall(plot_region_tag)
            if len(plot_areas) == 1
            else []
        )
        if (
            len(plot_areas) != 1
            or len(plot_regions) != 1
            or not len(plot_regions[0])
        ):
            raise ValueError("ChartEx Part has an invalid plot area")


def _validate_chart_relationship_roles(
    chart_root,
    source_part,
    relationship_type: str,
) -> None:
    external_data_tag = (
        f"{{{_CHART_NAMESPACE}}}externalData"
        if relationship_type == RT.CHART
        else f"{{{_CHARTEX_NAMESPACE}}}externalData"
    )
    external_data_nodes = list(chart_root.iter(external_data_tag))
    if len(external_data_nodes) > 1:
        raise ValueError("chart Part has duplicate external data references")
    for external_data in external_data_nodes:
        relationship_id = external_data.get(qn("r:id"))
        relationship = source_part.rels.get(relationship_id)
        if (
            not relationship_id
            or relationship is None
            or relationship.is_external
            or relationship.reltype != RT.PACKAGE
        ):
            raise ValueError(
                "chart external data has the wrong relationship type"
            )

    for sidecar_relationship in source_part.rels.values():
        if (
            sidecar_relationship.reltype
            in {
                _CHART_STYLE_RELATIONSHIP_TYPE,
                _CHART_COLOR_STYLE_RELATIONSHIP_TYPE,
            }
            and sidecar_relationship.is_external
        ):
            raise ValueError("chart style relationship cannot be external")

    if relationship_type != RT.CHART:
        return
    user_shape_nodes = list(
        chart_root.iter(f"{{{_CHART_NAMESPACE}}}userShapes")
    )
    if len(user_shape_nodes) > 1:
        raise ValueError("chart Part has duplicate user-shape references")
    for user_shapes in user_shape_nodes:
        relationship_id = user_shapes.get(qn("r:id"))
        relationship = source_part.rels.get(relationship_id)
        if (
            not relationship_id
            or relationship is None
            or relationship.is_external
            or relationship.reltype != RT.CHART_USER_SHAPES
            or str(relationship.target_part.content_type).casefold()
            != str(CT.DML_CHARTSHAPES).casefold()
        ):
            raise ValueError(
                "chart user shapes have an invalid relationship"
            )


def _chart_style_part_spec(relationship_type: str):
    return {
        _CHART_STYLE_RELATIONSHIP_TYPE: (
            _CHART_STYLE_CONTENT_TYPE,
            _CHART_STYLE_ROOT_TAG,
            "style",
        ),
        _CHART_COLOR_STYLE_RELATIONSHIP_TYPE: (
            _CHART_COLOR_STYLE_CONTENT_TYPE,
            _CHART_COLOR_STYLE_ROOT_TAG,
            "color",
        ),
    }.get(relationship_type)


def _chart_mc_namespace_for_prefix(element, prefix: str) -> str | None:
    if prefix == "xml":
        return _XML_NAMESPACE
    namespace = element.nsmap.get(prefix)
    return namespace if isinstance(namespace, str) and namespace else None


def _chart_mc_tokens(value: str | None) -> tuple[str, ...]:
    return tuple(
        token
        for token in re.split(r"[ \t\r\n]+", value or "")
        if token
    )


def _chart_mc_prefix_namespaces(
    element,
    value: str | None,
    label: str,
    *,
    forbid_mc: bool = False,
) -> frozenset[str]:
    namespaces = set()
    for prefix in _chart_mc_tokens(value):
        namespace = _chart_mc_namespace_for_prefix(element, prefix)
        if namespace is None or (
            forbid_mc and namespace == _MARKUP_COMPATIBILITY_NAMESPACE
        ):
            raise ValueError(f"chart style Part has an invalid {label} prefix")
        namespaces.add(namespace)
    return frozenset(namespaces)


def _chart_mc_qnames(
    element,
    value: str | None,
    local_ignorable: frozenset[str],
    label: str,
) -> frozenset[tuple[str, str]]:
    qnames = set()
    for token in _chart_mc_tokens(value):
        if token.count(":") != 1:
            raise ValueError(f"chart style Part has an invalid {label}")
        prefix, local_name = token.split(":", 1)
        namespace = _chart_mc_namespace_for_prefix(element, prefix)
        valid_local_name = local_name == "*"
        if not valid_local_name:
            try:
                etree.QName(namespace or "urn:invalid", local_name)
                valid_local_name = True
            except ValueError:
                pass
        if (
            namespace is None
            or namespace not in local_ignorable
            or not valid_local_name
        ):
            raise ValueError(f"chart style Part has an invalid {label}")
        qnames.add((namespace, local_name))
    return frozenset(qnames)


def _chart_mc_element_context(
    element,
    inherited_ignorable: frozenset[str],
    inherited_process_content: frozenset[tuple[str, str]],
    *,
    enforce_must_understand: bool = True,
) -> tuple[frozenset[str], frozenset[tuple[str, str]]]:
    mc_attributes = {}
    for attribute_name, value in element.attrib.items():
        qname = etree.QName(attribute_name)
        if qname.namespace != _MARKUP_COMPATIBILITY_NAMESPACE:
            continue
        if qname.localname not in _CHART_MC_ATTRIBUTE_LOCAL_NAMES:
            raise ValueError("chart style Part has an invalid MC attribute")
        mc_attributes[qname.localname] = value

    local_ignorable = _chart_mc_prefix_namespaces(
        element,
        mc_attributes.get("Ignorable"),
        "Ignorable",
        forbid_mc=True,
    )
    active_ignorable = inherited_ignorable | local_ignorable
    local_process_content = _chart_mc_qnames(
        element,
        mc_attributes.get("ProcessContent"),
        active_ignorable,
        "ProcessContent",
    )
    _chart_mc_qnames(
        element,
        mc_attributes.get("PreserveAttributes"),
        local_ignorable,
        "PreserveAttributes",
    )
    _chart_mc_qnames(
        element,
        mc_attributes.get("PreserveElements"),
        local_ignorable,
        "PreserveElements",
    )
    must_understand = _chart_mc_prefix_namespaces(
        element,
        mc_attributes.get("MustUnderstand"),
        "MustUnderstand",
    )
    if enforce_must_understand and not must_understand.issubset(
        _CHART_STYLE_UNDERSTOOD_ELEMENT_NAMESPACES
    ):
        raise ValueError(
            "chart style Part requires an unsupported MC namespace"
        )
    return (
        active_ignorable,
        inherited_process_content | local_process_content,
    )


def _chart_mc_processes(
    element,
    process_content: frozenset[tuple[str, str]],
) -> bool:
    qname = etree.QName(element)
    return (
        (qname.namespace, qname.localname) in process_content
        or (qname.namespace, "*") in process_content
    )


def _chart_mc_validate_alternate_attributes(
    element,
    allowed_unqualified: frozenset[str],
    ignorable: frozenset[str],
) -> None:
    for attribute_name in element.attrib:
        qname = etree.QName(attribute_name)
        if qname.namespace == _XML_NAMESPACE:
            raise ValueError(
                "chart style Part has invalid XML attributes on AlternateContent"
            )
        if qname.namespace is None and qname.localname not in allowed_unqualified:
            raise ValueError(
                "chart style Part has invalid AlternateContent attributes"
            )
        if (
            qname.namespace not in {None, _MARKUP_COMPATIBILITY_NAMESPACE}
            and qname.namespace not in ignorable
        ):
            raise ValueError(
                "chart style Part has non-ignorable AlternateContent attributes"
            )


def _chart_mc_choice_namespaces(choice) -> frozenset[str]:
    if "Requires" not in choice.attrib:
        raise ValueError("chart style Part has an AlternateContent choice without Requires")
    namespaces = set()
    prefixes = _chart_mc_tokens(choice.get("Requires"))
    if not prefixes:
        raise ValueError("chart style Part has an empty AlternateContent Requires")
    for prefix in prefixes:
        namespace = _chart_mc_namespace_for_prefix(choice, prefix)
        if namespace is None or namespace == _MARKUP_COMPATIBILITY_NAMESPACE:
            raise ValueError(
                "chart style Part has an invalid AlternateContent Requires prefix"
            )
        namespaces.add(namespace)
    return frozenset(namespaces)


def _chart_mc_select_alternate_content(
    alternate_content,
    inherited_ignorable: frozenset[str],
    inherited_process_content: frozenset[tuple[str, str]],
) -> list:
    choice_tag = f"{{{_MARKUP_COMPATIBILITY_NAMESPACE}}}Choice"
    fallback_tag = f"{{{_MARKUP_COMPATIBILITY_NAMESPACE}}}Fallback"
    children = _chart_schema_children(alternate_content)
    choices = []
    fallback = None
    state = "choice"
    for child in children:
        if child.tag == choice_tag and state == "choice":
            choices.append(child)
            continue
        if child.tag == fallback_tag and state == "choice":
            fallback = child
            state = "fallback"
            continue
        child_qname = etree.QName(child)
        if (
            child_qname.namespace in inherited_ignorable
            and child_qname.namespace
            not in _CHART_STYLE_UNDERSTOOD_ELEMENT_NAMESPACES
        ):
            continue
        raise ValueError("chart style Part has invalid AlternateContent children")
    if not choices:
        raise ValueError("chart style Part has no AlternateContent choice")

    selected = None
    selected_context = None
    for choice in choices:
        choice_context = _chart_mc_element_context(
            choice,
            inherited_ignorable,
            inherited_process_content,
            enforce_must_understand=False,
        )
        _chart_mc_validate_alternate_attributes(
            choice,
            frozenset({"Requires"}),
            choice_context[0],
        )
        requirements = _chart_mc_choice_namespaces(choice)
        if (
            selected is None
            and requirements
            and requirements.issubset(
                _CHART_STYLE_UNDERSTOOD_ELEMENT_NAMESPACES
            )
        ):
            selected = choice
            selected_context = _chart_mc_element_context(
                choice,
                inherited_ignorable,
                inherited_process_content,
            )
    if fallback is not None:
        fallback_context = _chart_mc_element_context(
            fallback,
            inherited_ignorable,
            inherited_process_content,
            enforce_must_understand=False,
        )
        _chart_mc_validate_alternate_attributes(
            fallback,
            frozenset(),
            fallback_context[0],
        )
        if selected is None:
            selected = fallback
            selected_context = _chart_mc_element_context(
                fallback,
                inherited_ignorable,
                inherited_process_content,
            )
    if selected is None or selected_context is None:
        return []

    logical_children = []
    for child in list(selected):
        logical_children.extend(
            _chart_mc_transform_element(
                child,
                selected_context[0],
                selected_context[1],
            )
        )
    return logical_children


def _chart_mc_transform_element(
    element,
    inherited_ignorable: frozenset[str],
    inherited_process_content: frozenset[tuple[str, str]],
) -> list:
    if not isinstance(element.tag, str):
        return [element]

    alternate_content_tag = (
        f"{{{_MARKUP_COMPATIBILITY_NAMESPACE}}}AlternateContent"
    )
    ignorable, process_content = _chart_mc_element_context(
        element,
        inherited_ignorable,
        inherited_process_content,
    )
    if element.tag == alternate_content_tag:
        _chart_mc_validate_alternate_attributes(
            element,
            frozenset(),
            ignorable,
        )
    element_qname = etree.QName(element)
    understood = (
        element_qname.namespace
        in _CHART_STYLE_UNDERSTOOD_ELEMENT_NAMESPACES
    )
    process_wrapper = (
        not understood
        and element_qname.namespace in ignorable
        and _chart_mc_processes(element, process_content)
    )
    if process_wrapper and any(
        etree.QName(attribute_name).namespace == _XML_NAMESPACE
        and etree.QName(attribute_name).localname in {"base", "lang", "space"}
        for attribute_name in element.attrib
    ):
        raise ValueError(
            "chart style Part has XML attributes on ProcessContent"
        )
    for attribute_name in tuple(element.attrib):
        qname = etree.QName(attribute_name)
        if qname.namespace == _MARKUP_COMPATIBILITY_NAMESPACE:
            del element.attrib[attribute_name]
        elif qname.namespace == _XML_NAMESPACE:
            del element.attrib[attribute_name]
        elif (
            qname.namespace in ignorable
            and qname.namespace
            not in _CHART_STYLE_UNDERSTOOD_ATTRIBUTE_NAMESPACES
            and attribute_name
            != _LEGACY_SPREADSHEET_COLOR_INDEX_ATTRIBUTE
        ):
            del element.attrib[attribute_name]

    if element.tag == alternate_content_tag:
        return _chart_mc_select_alternate_content(
            element,
            ignorable,
            process_content,
        )

    if not understood and element_qname.namespace in ignorable:
        if not process_wrapper:
            return []
        logical_children = []
        for child in list(element):
            logical_children.extend(
                _chart_mc_transform_element(
                    child,
                    ignorable,
                    process_content,
                )
            )
        return logical_children

    if element.tag == _DRAWING_EXTENSION_TAG:
        return [element]

    logical_children = []
    for child in list(element):
        logical_children.extend(
            _chart_mc_transform_element(
                child,
                ignorable,
                process_content,
            )
        )
    for child in list(element):
        element.remove(child)
    element.extend(logical_children)
    return [element]


def _chart_mc_schema_view(root):
    """生成应用 MC Ignorable/ProcessContent 后的只读校验视图。"""
    view = deepcopy(root)
    transformed = _chart_mc_transform_element(
        view,
        frozenset(),
        frozenset(),
    )
    if len(transformed) != 1 or transformed[0] is not view:
        raise ValueError("chart style Part has an invalid MC root")
    return view


def _chart_schema_children(parent) -> tuple:
    """返回 schema 元素子节点；SDK 粒子匹配会跳过注释和 PI。"""
    return tuple(child for child in parent if isinstance(child.tag, str))


def _chart_leaf_text(element) -> str | None:
    """读取允许穿插注释/PI 的叶节点文本，出现元素子节点时返回 None。"""
    pieces = [element.text or ""]
    for child in element:
        if isinstance(child.tag, str):
            return None
        pieces.append(child.tail or "")
    return "".join(pieces)


def _validate_chart_composite_text(element, label: str) -> None:
    """复合 DrawingML 元素只允许 XML 四种空白作为混合文本。"""
    text_nodes = [element.text or ""]
    text_nodes.extend(child.tail or "" for child in element)
    if any(text.strip(" \t\r\n") for text in text_nodes):
        raise ValueError(f"{label} has invalid text content")


def _validate_chart_extension_list(
    extension_list,
    label: str,
    *,
    uri_required: bool = False,
) -> None:
    if extension_list.attrib:
        raise ValueError(f"{label} has invalid attributes")
    extension_tag = f"{{{_THEME_NAMESPACE}}}ext"
    for extension in _chart_schema_children(extension_list):
        if (
            extension.tag != extension_tag
            or not set(extension.attrib).issubset({"uri"})
            or (uri_required and "uri" not in extension.attrib)
        ):
            raise ValueError(f"{label} has invalid children")
        uri = extension.get("uri")
        if uri is not None and (
            "\t" in uri
            or "\r" in uri
            or "\n" in uri
            or uri.startswith(" ")
            or uri.endswith(" ")
            or "  " in uri
        ):
            raise ValueError(f"{label} has an invalid extension URI")
        if len(_chart_schema_children(extension)) > 1:
            raise ValueError(f"{label} has invalid extension content")


def _validate_chart_ordered_choice_children(
    element,
    groups: tuple[frozenset[str], ...],
    label: str,
) -> tuple:
    group_by_tag = {
        tag: group_index
        for group_index, group in enumerate(groups)
        for tag in group
    }
    children = _chart_schema_children(element)
    observed_groups = []
    for child in children:
        group_index = group_by_tag.get(child.tag)
        if group_index is None:
            raise ValueError(f"{label} has invalid children")
        observed_groups.append(group_index)
    if (
        len(observed_groups) != len(set(observed_groups))
        or observed_groups != sorted(observed_groups)
    ):
        raise ValueError(f"{label} has invalid child order")
    return children


def _drawing_tags(*local_names: str) -> frozenset[str]:
    return frozenset(
        f"{{{_THEME_NAMESPACE}}}{local_name}"
        for local_name in local_names
    )


def _validate_chart_boolean_attributes(
    element,
    attribute_names: tuple[str, ...],
    label: str,
) -> None:
    if any(
        name in element.attrib and not _is_xsd_boolean(element.get(name))
        for name in attribute_names
    ):
        raise ValueError(f"{label} has an invalid boolean value")


def _validate_chart_int32_attribute(
    element,
    attribute_name: str,
    label: str,
    *,
    minimum: int = -0x80000000,
    maximum: int = 0x7FFFFFFF,
) -> None:
    value = element.get(attribute_name)
    if value is None:
        return
    parsed = _parse_xsd_int32(value)
    if parsed is None or not minimum <= parsed <= maximum:
        raise ValueError(f"{label} has an invalid {attribute_name} value")


def _validate_chart_int64_attribute(
    element,
    attribute_name: str,
    label: str,
    *,
    minimum: int = -0x8000000000000000,
    maximum: int = 0x7FFFFFFFFFFFFFFF,
) -> None:
    value = element.get(attribute_name)
    if value is None:
        return
    parsed = _parse_xsd_int64(value)
    if parsed is None or not minimum <= parsed <= maximum:
        raise ValueError(f"{label} has an invalid {attribute_name} value")


def _validate_chart_enum_attribute(
    element,
    attribute_name: str,
    values: frozenset[str],
    label: str,
) -> None:
    value = element.get(attribute_name)
    if value is not None and value not in values:
        raise ValueError(f"{label} has an invalid {attribute_name} value")


def _validate_chart_token_attribute(
    element,
    attribute_name: str,
    label: str,
    *,
    required: bool = False,
) -> None:
    value = element.get(attribute_name)
    if value is None:
        if required:
            raise ValueError(f"{label} is missing {attribute_name}")
        return
    if (
        value != value.strip(" \t\r\n")
        or any(character in value for character in "\t\r\n")
        or "  " in value
    ):
        raise ValueError(f"{label} has an invalid {attribute_name} value")


def _validate_chart_empty_drawing_leaf(element, label: str) -> None:
    if element.attrib:
        raise ValueError(f"{label} has invalid content")
    _validate_chart_drawing_leaf_text(element, label)


def _validate_chart_drawing_leaf_text(element, label: str) -> None:
    text = _chart_leaf_text(element)
    if text is None or text.strip(" \t\r\n"):
        raise ValueError(f"{label} has invalid content")


def _validate_chart_solid_fill(solid_fill) -> None:
    if solid_fill.attrib:
        raise ValueError("chart style solid fill has invalid attributes")
    _validate_chart_composite_text(
        solid_fill,
        "chart style solid fill",
    )
    children = _chart_schema_children(solid_fill)
    if len(children) > 1 or (
        children and children[0].tag not in _THEME_COLOR_MODEL_TAGS
    ):
        raise ValueError("chart style solid fill has invalid children")
    if children:
        _validate_chart_style_color_model(
            children[0],
            "chart style solid fill",
        )


def _validate_chart_relative_rectangle(rectangle, label: str) -> None:
    if not set(rectangle.attrib).issubset({"l", "t", "r", "b"}):
        raise ValueError(f"{label} has invalid attributes")
    for attribute_name in ("l", "t", "r", "b"):
        _validate_chart_int32_attribute(
            rectangle,
            attribute_name,
            label,
        )
    _validate_chart_drawing_leaf_text(rectangle, label)


def _validate_chart_gradient_stop(stop) -> None:
    label = "chart style gradient stop"
    if set(stop.attrib) != {"pos"}:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_int32_attribute(
        stop,
        "pos",
        label,
        minimum=0,
        maximum=100000,
    )
    children = _chart_schema_children(stop)
    if len(children) != 1 or children[0].tag not in _THEME_COLOR_MODEL_TAGS:
        raise ValueError(f"{label} has invalid color content")
    _validate_chart_composite_text(stop, label)
    _validate_chart_style_color_model(children[0], label)


def _validate_chart_gradient_fill(gradient_fill) -> None:
    label = "chart style gradient fill"
    if not set(gradient_fill.attrib).issubset({"flip", "rotWithShape"}):
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_enum_attribute(
        gradient_fill,
        "flip",
        _CHART_TILE_FLIP_VALUES,
        label,
    )
    _validate_chart_boolean_attributes(
        gradient_fill,
        ("rotWithShape",),
        label,
    )
    _validate_chart_composite_text(gradient_fill, label)
    children = _validate_chart_ordered_choice_children(
        gradient_fill,
        (
            _drawing_tags("gsLst"),
            _drawing_tags("lin", "path"),
            _drawing_tags("tileRect"),
        ),
        label,
    )
    for child in children:
        if child.tag == f"{{{_THEME_NAMESPACE}}}gsLst":
            if child.attrib:
                raise ValueError(f"{label} gradient-stop list has attributes")
            _validate_chart_composite_text(child, f"{label} gradient-stop list")
            stops = _chart_schema_children(child)
            if len(stops) < 2 or any(
                stop.tag != f"{{{_THEME_NAMESPACE}}}gs" for stop in stops
            ):
                raise ValueError(f"{label} has invalid gradient stops")
            for stop in stops:
                _validate_chart_gradient_stop(stop)
        elif child.tag == f"{{{_THEME_NAMESPACE}}}lin":
            if not set(child.attrib).issubset({"ang", "scaled"}):
                raise ValueError(f"{label} linear shade has invalid attributes")
            _validate_chart_int32_attribute(
                child,
                "ang",
                f"{label} linear shade",
                minimum=0,
                maximum=21599999,
            )
            _validate_chart_boolean_attributes(
                child,
                ("scaled",),
                f"{label} linear shade",
            )
            _validate_chart_drawing_leaf_text(
                child,
                f"{label} linear shade",
            )
        elif child.tag == f"{{{_THEME_NAMESPACE}}}path":
            if not set(child.attrib).issubset({"path"}):
                raise ValueError(f"{label} path shade has invalid attributes")
            _validate_chart_enum_attribute(
                child,
                "path",
                _CHART_PATH_SHADE_VALUES,
                f"{label} path shade",
            )
            path_children = _chart_schema_children(child)
            if len(path_children) > 1 or (
                path_children
                and path_children[0].tag
                != f"{{{_THEME_NAMESPACE}}}fillToRect"
            ):
                raise ValueError(f"{label} path shade has invalid children")
            _validate_chart_composite_text(child, f"{label} path shade")
            if path_children:
                _validate_chart_relative_rectangle(
                    path_children[0],
                    f"{label} fill-to-rectangle",
                )
        else:
            _validate_chart_relative_rectangle(
                child,
                f"{label} tile rectangle",
            )


def _validate_chart_color_wrapper(color_wrapper, label: str) -> None:
    if color_wrapper.attrib:
        raise ValueError(f"{label} has invalid attributes")
    children = _chart_schema_children(color_wrapper)
    if len(children) != 1 or children[0].tag not in _THEME_COLOR_MODEL_TAGS:
        raise ValueError(f"{label} has invalid color content")
    _validate_chart_composite_text(color_wrapper, label)
    _validate_chart_style_color_model(children[0], label)


def _validate_chart_pattern_fill(pattern_fill) -> None:
    label = "chart style pattern fill"
    if not set(pattern_fill.attrib).issubset({"prst"}):
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_enum_attribute(
        pattern_fill,
        "prst",
        _CHART_PRESET_PATTERN_VALUES,
        label,
    )
    _validate_chart_composite_text(pattern_fill, label)
    children = _validate_chart_ordered_choice_children(
        pattern_fill,
        (
            _drawing_tags("fgClr"),
            _drawing_tags("bgClr"),
        ),
        label,
    )
    for child in children:
        _validate_chart_color_wrapper(
            child,
            f"{label} {'foreground' if child.tag.endswith('fgClr') else 'background'} color",
        )


def _validate_chart_blip(blip) -> None:
    label = "chart style blip"
    embed_attribute = f"{{{_RELATIONSHIP_NAMESPACE}}}embed"
    link_attribute = f"{{{_RELATIONSHIP_NAMESPACE}}}link"
    if not set(blip.attrib).issubset(
        {embed_attribute, link_attribute, "cstate"}
    ):
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_enum_attribute(
        blip,
        "cstate",
        _CHART_BLIP_COMPRESSION_VALUES,
        label,
    )
    _validate_chart_composite_text(blip, label)
    effect_tags = _drawing_tags(*_CHART_BLIP_EFFECT_NAMES)
    extension_list_tag = f"{{{_THEME_NAMESPACE}}}extLst"
    saw_extension_list = False
    for child in _chart_schema_children(blip):
        if child.tag == extension_list_tag:
            if saw_extension_list:
                raise ValueError(f"{label} has duplicate extension lists")
            saw_extension_list = True
            _validate_chart_extension_list(
                child,
                f"{label} extension list",
                uri_required=True,
            )
            continue
        if child.tag not in effect_tags or saw_extension_list:
            raise ValueError(f"{label} has invalid children")
        _validate_chart_effect_node(
            child,
            f"{label} {etree.QName(child).localname}",
        )


def _validate_chart_blip_fill(blip_fill) -> None:
    label = "chart style blip fill"
    if not set(blip_fill.attrib).issubset({"dpi", "rotWithShape"}):
        raise ValueError(f"{label} has invalid attributes")
    dpi = blip_fill.get("dpi")
    if dpi is not None and _parse_xsd_uint32(dpi) is None:
        raise ValueError(f"{label} has an invalid dpi value")
    _validate_chart_boolean_attributes(
        blip_fill,
        ("rotWithShape",),
        label,
    )
    _validate_chart_composite_text(blip_fill, label)
    children = _validate_chart_ordered_choice_children(
        blip_fill,
        (
            _drawing_tags("blip"),
            _drawing_tags("srcRect"),
            _drawing_tags("tile", "stretch"),
        ),
        label,
    )
    for child in children:
        local_name = etree.QName(child).localname
        child_label = f"{label} {local_name}"
        if local_name == "blip":
            _validate_chart_blip(child)
        elif local_name == "srcRect":
            _validate_chart_relative_rectangle(child, child_label)
        elif local_name == "tile":
            if not set(child.attrib).issubset(
                {"tx", "ty", "sx", "sy", "flip", "algn"}
            ):
                raise ValueError(f"{child_label} has invalid attributes")
            for attribute_name in ("tx", "ty"):
                _validate_chart_int64_attribute(
                    child,
                    attribute_name,
                    child_label,
                    minimum=-27273042329600,
                    maximum=27273042316900,
                )
            for attribute_name in ("sx", "sy"):
                _validate_chart_int32_attribute(
                    child,
                    attribute_name,
                    child_label,
                )
            _validate_chart_enum_attribute(
                child,
                "flip",
                _CHART_TILE_FLIP_VALUES,
                child_label,
            )
            _validate_chart_enum_attribute(
                child,
                "algn",
                _CHART_RECTANGLE_ALIGNMENT_VALUES,
                child_label,
            )
            _validate_chart_drawing_leaf_text(child, child_label)
        else:
            if child.attrib:
                raise ValueError(f"{child_label} has invalid attributes")
            _validate_chart_composite_text(child, child_label)
            stretch_children = _chart_schema_children(child)
            if len(stretch_children) > 1 or (
                stretch_children
                and stretch_children[0].tag
                != f"{{{_THEME_NAMESPACE}}}fillRect"
            ):
                raise ValueError(f"{child_label} has invalid children")
            if stretch_children:
                _validate_chart_relative_rectangle(
                    stretch_children[0],
                    f"{child_label} fill rectangle",
                )


def _validate_chart_preset_dash(preset_dash) -> None:
    text = _chart_leaf_text(preset_dash)
    if (
        not set(preset_dash.attrib).issubset({"val"})
        or text is None
        or text.strip(" \t\r\n")
    ):
        raise ValueError("chart style preset dash has invalid content")
    _validate_chart_enum_attribute(
        preset_dash,
        "val",
        _CHART_PRESET_LINE_DASH_VALUES,
        "chart style preset dash",
    )


def _validate_chart_custom_dash(custom_dash) -> None:
    label = "chart style custom dash"
    if custom_dash.attrib:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_composite_text(custom_dash, label)
    for dash_stop in _chart_schema_children(custom_dash):
        if dash_stop.tag != f"{{{_THEME_NAMESPACE}}}ds":
            raise ValueError(f"{label} has invalid children")
        if set(dash_stop.attrib) != {"d", "sp"}:
            raise ValueError(f"{label} dash stop has invalid attributes")
        _validate_chart_int32_attribute(
            dash_stop,
            "d",
            f"{label} dash stop",
            minimum=1,
        )
        _validate_chart_int32_attribute(
            dash_stop,
            "sp",
            f"{label} dash stop",
            minimum=1,
        )
        _validate_chart_drawing_leaf_text(
            dash_stop,
            f"{label} dash stop",
        )


def _validate_chart_line_join(join) -> None:
    local_name = etree.QName(join).localname
    label = f"chart style {local_name} line join"
    if local_name in {"round", "bevel"}:
        _validate_chart_empty_drawing_leaf(join, label)
        return
    if local_name != "miter":
        raise ValueError(f"{label} has an invalid element")
    if not set(join.attrib).issubset({"lim"}):
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_int32_attribute(join, "lim", label, minimum=0)
    _validate_chart_drawing_leaf_text(join, label)


def _validate_chart_line_end(line_end) -> None:
    label = f"chart style {etree.QName(line_end).localname} line end"
    if not set(line_end.attrib).issubset({"type", "w", "len"}):
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_enum_attribute(
        line_end,
        "type",
        _CHART_LINE_END_TYPE_VALUES,
        label,
    )
    for attribute_name in ("w", "len"):
        _validate_chart_enum_attribute(
            line_end,
            attribute_name,
            _CHART_LINE_END_SIZE_VALUES,
            label,
        )
    _validate_chart_drawing_leaf_text(line_end, label)


def _validate_chart_line_properties(line) -> None:
    label = "chart style line properties"
    if not set(line.attrib).issubset({"w", "cap", "cmpd", "algn"}):
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_composite_text(line, label)
    _validate_chart_int32_attribute(
        line,
        "w",
        label,
        minimum=0,
        maximum=20116800,
    )
    for attribute_name, values in (
        ("cap", _CHART_LINE_CAP_VALUES),
        ("cmpd", _CHART_COMPOUND_LINE_VALUES),
        ("algn", _CHART_PEN_ALIGNMENT_VALUES),
    ):
        _validate_chart_enum_attribute(line, attribute_name, values, label)
    groups = (
        _drawing_tags("noFill", "solidFill", "gradFill", "pattFill"),
        _drawing_tags("prstDash", "custDash"),
        _drawing_tags("round", "bevel", "miter"),
        _drawing_tags("headEnd"),
        _drawing_tags("tailEnd"),
        _drawing_tags("extLst"),
    )
    children = _validate_chart_ordered_choice_children(line, groups, label)
    if children and children[-1].tag == f"{{{_THEME_NAMESPACE}}}extLst":
        _validate_chart_extension_list(
            children[-1],
            "chart style line-properties extension list",
            uri_required=True,
        )
    for child in children:
        if child.tag == f"{{{_THEME_NAMESPACE}}}custDash":
            _validate_chart_custom_dash(child)
        elif child.tag in {
            f"{{{_THEME_NAMESPACE}}}round",
            f"{{{_THEME_NAMESPACE}}}bevel",
            f"{{{_THEME_NAMESPACE}}}miter",
        }:
            _validate_chart_line_join(child)
        elif child.tag in {
            f"{{{_THEME_NAMESPACE}}}headEnd",
            f"{{{_THEME_NAMESPACE}}}tailEnd",
        }:
            _validate_chart_line_end(child)


def _validate_chart_geometry_guide_list(guide_list, label: str) -> None:
    if guide_list.attrib:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_composite_text(guide_list, label)
    guide_tag = f"{{{_THEME_NAMESPACE}}}gd"
    for guide in _chart_schema_children(guide_list):
        if guide.tag != guide_tag:
            raise ValueError(f"{label} has invalid children")
        if set(guide.attrib) != {"name", "fmla"}:
            raise ValueError(f"{label} guide has invalid attributes")
        _validate_chart_token_attribute(guide, "name", f"{label} guide", required=True)
        if guide.get("fmla") is None:
            raise ValueError(f"{label} guide is missing fmla")
        _validate_chart_drawing_leaf_text(guide, f"{label} guide")


def _validate_chart_preset_text_warp(warp) -> None:
    label = "chart style preset text warp"
    if set(warp.attrib) != {"prst"}:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_enum_attribute(
        warp,
        "prst",
        _CHART_TEXT_SHAPE_VALUES,
        label,
    )
    _validate_chart_composite_text(warp, label)
    children = _validate_chart_ordered_choice_children(
        warp,
        (_drawing_tags("avLst"),),
        label,
    )
    if children:
        _validate_chart_geometry_guide_list(children[0], f"{label} adjustment list")


def _validate_chart_preset_geometry(geometry) -> None:
    label = "chart style preset geometry"
    if set(geometry.attrib) != {"prst"}:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_enum_attribute(
        geometry,
        "prst",
        _CHART_SHAPE_TYPE_VALUES,
        label,
    )
    _validate_chart_composite_text(geometry, label)
    children = _validate_chart_ordered_choice_children(
        geometry,
        (_drawing_tags("avLst"),),
        label,
    )
    if children:
        _validate_chart_geometry_guide_list(children[0], f"{label} adjustment list")


def _validate_chart_union_attribute(
    element,
    attribute_name: str,
    label: str,
    *,
    required: bool = False,
) -> None:
    if element.get(attribute_name) is None and required:
        raise ValueError(f"{label} is missing {attribute_name}")


def _validate_chart_geometry_position(position, label: str) -> None:
    if set(position.attrib) != {"x", "y"}:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_union_attribute(position, "x", label, required=True)
    _validate_chart_union_attribute(position, "y", label, required=True)
    _validate_chart_drawing_leaf_text(position, label)


def _validate_chart_geometry_rectangle(rectangle, label: str) -> None:
    if set(rectangle.attrib) != {"l", "t", "r", "b"}:
        raise ValueError(f"{label} has invalid attributes")
    for attribute_name in ("l", "t", "r", "b"):
        _validate_chart_union_attribute(
            rectangle,
            attribute_name,
            label,
            required=True,
        )
    _validate_chart_drawing_leaf_text(rectangle, label)


def _validate_chart_adjust_handle(handle, label: str) -> None:
    local_name = etree.QName(handle).localname
    if local_name == "ahXY":
        allowed = {"gdRefX", "gdRefY", "minX", "maxX", "minY", "maxY"}
        token_names = ("gdRefX", "gdRefY")
        coordinate_names = ("minX", "maxX", "minY", "maxY")
    elif local_name == "ahPolar":
        allowed = {"gdRefR", "gdRefAng", "minR", "maxR", "minAng", "maxAng"}
        token_names = ("gdRefR", "gdRefAng")
        coordinate_names = ("minR", "maxR", "minAng", "maxAng")
    else:
        raise ValueError(f"{label} has an invalid handle element")
    if not set(handle.attrib).issubset(allowed):
        raise ValueError(f"{label} has invalid attributes")
    for attribute_name in token_names:
        _validate_chart_token_attribute(handle, attribute_name, label)
    for attribute_name in coordinate_names:
        _validate_chart_union_attribute(handle, attribute_name, label)
    _validate_chart_composite_text(handle, label)
    children = _validate_chart_ordered_choice_children(
        handle,
        (_drawing_tags("pos"),),
        label,
    )
    if len(children) != 1:
        raise ValueError(f"{label} is missing position")
    _validate_chart_geometry_position(children[0], f"{label} position")


def _validate_chart_adjust_handle_list(handle_list, label: str) -> None:
    if handle_list.attrib:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_composite_text(handle_list, label)
    allowed = _drawing_tags("ahXY", "ahPolar")
    for child in _chart_schema_children(handle_list):
        if child.tag not in allowed:
            raise ValueError(f"{label} has invalid children")
        _validate_chart_adjust_handle(child, f"{label} {etree.QName(child).localname}")


def _validate_chart_connection_site(site, label: str) -> None:
    if set(site.attrib) != {"ang"}:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_union_attribute(site, "ang", label, required=True)
    _validate_chart_composite_text(site, label)
    children = _validate_chart_ordered_choice_children(
        site,
        (_drawing_tags("pos"),),
        label,
    )
    if len(children) != 1:
        raise ValueError(f"{label} is missing position")
    _validate_chart_geometry_position(children[0], f"{label} position")


def _validate_chart_connection_site_list(site_list, label: str) -> None:
    if site_list.attrib:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_composite_text(site_list, label)
    site_tag = f"{{{_THEME_NAMESPACE}}}cxn"
    for child in _chart_schema_children(site_list):
        if child.tag != site_tag:
            raise ValueError(f"{label} has invalid children")
        _validate_chart_connection_site(child, f"{label} connection site")


def _validate_chart_geometry_path_command(command, label: str) -> None:
    local_name = etree.QName(command).localname
    if local_name == "close":
        _validate_chart_empty_drawing_leaf(command, label)
        return
    if local_name in {"moveTo", "lnTo"}:
        expected_points = 1
    elif local_name == "quadBezTo":
        expected_points = 2
    elif local_name == "cubicBezTo":
        expected_points = 3
    elif local_name == "arcTo":
        if set(command.attrib) != {"wR", "hR", "stAng", "swAng"}:
            raise ValueError(f"{label} has invalid attributes")
        for attribute_name in ("wR", "hR", "stAng", "swAng"):
            _validate_chart_union_attribute(
                command,
                attribute_name,
                label,
                required=True,
            )
        _validate_chart_drawing_leaf_text(command, label)
        return
    else:
        raise ValueError(f"{label} has an invalid path command")
    if command.attrib:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_composite_text(command, label)
    points = _chart_schema_children(command)
    if len(points) != expected_points:
        raise ValueError(f"{label} has an invalid point count")
    for point in points:
        if point.tag != f"{{{_THEME_NAMESPACE}}}pt":
            raise ValueError(f"{label} has invalid point children")
        _validate_chart_geometry_position(point, f"{label} point")


def _validate_chart_geometry_path(path, label: str) -> None:
    allowed = {"w", "h", "fill", "stroke", "extrusionOk"}
    if not set(path.attrib).issubset(allowed):
        raise ValueError(f"{label} has invalid attributes")
    for attribute_name in ("w", "h"):
        _validate_chart_int64_attribute(
            path,
            attribute_name,
            label,
            minimum=0,
            maximum=0x7FFFFFFF,
        )
    _validate_chart_enum_attribute(path, "fill", _CHART_PATH_FILL_VALUES, label)
    _validate_chart_boolean_attributes(path, ("stroke", "extrusionOk"), label)
    _validate_chart_composite_text(path, label)
    allowed_commands = _drawing_tags(
        "close", "moveTo", "lnTo", "arcTo", "quadBezTo", "cubicBezTo"
    )
    for child in _chart_schema_children(path):
        if child.tag not in allowed_commands:
            raise ValueError(f"{label} has invalid commands")
        _validate_chart_geometry_path_command(
            child,
            f"{label} {etree.QName(child).localname}",
        )


def _validate_chart_geometry_path_list(path_list, label: str) -> None:
    if path_list.attrib:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_composite_text(path_list, label)
    path_tag = f"{{{_THEME_NAMESPACE}}}path"
    for child in _chart_schema_children(path_list):
        if child.tag != path_tag:
            raise ValueError(f"{label} has invalid children")
        _validate_chart_geometry_path(child, f"{label} path")


def _validate_chart_custom_geometry(geometry) -> None:
    label = "chart style custom geometry"
    if geometry.attrib:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_composite_text(geometry, label)
    children = _validate_chart_ordered_choice_children(
        geometry,
        (
            _drawing_tags("avLst"),
            _drawing_tags("gdLst"),
            _drawing_tags("ahLst"),
            _drawing_tags("cxnLst"),
            _drawing_tags("rect"),
            _drawing_tags("pathLst"),
        ),
        label,
    )
    path_list_tag = f"{{{_THEME_NAMESPACE}}}pathLst"
    if not any(child.tag == path_list_tag for child in children):
        raise ValueError(f"{label} is missing path list")
    for child in children:
        local_name = etree.QName(child).localname
        if local_name in {"avLst", "gdLst"}:
            _validate_chart_geometry_guide_list(child, f"{label} {local_name}")
        elif local_name == "ahLst":
            _validate_chart_adjust_handle_list(child, f"{label} adjust handles")
        elif local_name == "cxnLst":
            _validate_chart_connection_site_list(child, f"{label} connection sites")
        elif local_name == "rect":
            _validate_chart_geometry_rectangle(child, f"{label} rectangle")
        else:
            _validate_chart_geometry_path_list(child, f"{label} paths")


def _validate_chart_transform_2d(transform) -> None:
    label = "chart style 2D transform"
    if not set(transform.attrib).issubset({"rot", "flipH", "flipV"}):
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_int32_attribute(transform, "rot", label)
    _validate_chart_boolean_attributes(transform, ("flipH", "flipV"), label)
    _validate_chart_composite_text(transform, label)
    children = _validate_chart_ordered_choice_children(
        transform,
        (_drawing_tags("off"), _drawing_tags("ext")),
        label,
    )
    for child in children:
        local_name = etree.QName(child).localname
        child_label = f"{label} {local_name}"
        if local_name == "off":
            if set(child.attrib) != {"x", "y"}:
                raise ValueError(f"{child_label} has invalid attributes")
            for attribute_name in ("x", "y"):
                _validate_chart_int64_attribute(
                    child,
                    attribute_name,
                    child_label,
                    minimum=-27273042329600,
                    maximum=27273042316900,
                )
        else:
            if set(child.attrib) != {"cx", "cy"}:
                raise ValueError(f"{child_label} has invalid attributes")
            for attribute_name in ("cx", "cy"):
                _validate_chart_int64_attribute(
                    child,
                    attribute_name,
                    child_label,
                    minimum=0,
                    maximum=0x7FFFFFFF,
                )
        _validate_chart_drawing_leaf_text(child, child_label)


def _validate_chart_text_autofit(element) -> None:
    local_name = etree.QName(element).localname
    label = f"chart style {local_name}"
    if local_name in {"noAutofit", "spAutoFit"}:
        _validate_chart_empty_drawing_leaf(element, label)
        return
    if local_name != "normAutofit":
        raise ValueError(f"{label} has an invalid element")
    if not set(element.attrib).issubset({"fontScale", "lnSpcReduction"}):
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_int32_attribute(
        element,
        "fontScale",
        label,
        minimum=1000,
        maximum=100000,
    )
    _validate_chart_int32_attribute(
        element,
        "lnSpcReduction",
        label,
        minimum=0,
        maximum=13200000,
    )
    _validate_chart_drawing_leaf_text(element, label)


def _validate_chart_text_font(font) -> None:
    label = f"chart style {etree.QName(font).localname} font"
    if not set(font.attrib).issubset(
        {"typeface", "panose", "pitchFamily", "charset"}
    ):
        raise ValueError(f"{label} has invalid attributes")
    panose = font.get("panose")
    if panose is not None:
        normalized_panose = re.sub(r"[ \t\r\n]", "", panose)
        if (
            len(normalized_panose) != 20
            or not normalized_panose.isascii()
            or not all(
                character in "0123456789abcdefABCDEF"
                for character in normalized_panose
            )
        ):
            raise ValueError(f"{label} has an invalid panose value")
    for attribute_name in ("pitchFamily", "charset"):
        _validate_chart_int32_attribute(
            font,
            attribute_name,
            label,
            minimum=-128,
            maximum=127,
        )
    _validate_chart_drawing_leaf_text(font, label)


def _validate_chart_hyperlink_extension_list(extension_list, label: str) -> None:
    if extension_list.attrib:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_composite_text(extension_list, label)
    extension_tag = f"{{{_THEME_NAMESPACE}}}ext"
    for extension in _chart_schema_children(extension_list):
        if extension.tag != extension_tag or "uri" not in extension.attrib:
            raise ValueError(f"{label} has invalid children")
        if set(extension.attrib) != {"uri"}:
            raise ValueError(f"{label} extension has invalid attributes")
        _validate_chart_token_attribute(extension, "uri", f"{label} extension", required=True)
        _validate_chart_composite_text(extension, f"{label} extension")
        if len(_chart_schema_children(extension)) > 1:
            raise ValueError(f"{label} extension has invalid content")


def _validate_chart_hyperlink(hyperlink) -> None:
    label = f"chart style {etree.QName(hyperlink).localname} hyperlink"
    allowed = {
        _RELATIONSHIP_ID_ATTRIBUTE,
        "invalidUrl", "action", "tgtFrame", "tooltip",
        "history", "highlightClick", "endSnd",
    }
    if not set(hyperlink.attrib).issubset(allowed):
        raise ValueError(f"{label} has invalid attributes")
    if _RELATIONSHIP_ID_ATTRIBUTE not in hyperlink.attrib:
        raise ValueError(f"{label} is missing relationship ID")
    _validate_chart_boolean_attributes(
        hyperlink,
        ("history", "highlightClick", "endSnd"),
        label,
    )
    _validate_chart_composite_text(hyperlink, label)
    children = _validate_chart_ordered_choice_children(
        hyperlink,
        (_drawing_tags("snd"), _drawing_tags("extLst")),
        label,
    )
    for child in children:
        if child.tag == f"{{{_THEME_NAMESPACE}}}snd":
            embed = f"{{{_RELATIONSHIP_NAMESPACE}}}embed"
            if embed not in child.attrib or not set(child.attrib).issubset(
                {embed, "name", "builtIn"}
            ):
                raise ValueError(f"{label} sound has invalid attributes")
            _validate_chart_boolean_attributes(child, ("builtIn",), label)
            _validate_chart_drawing_leaf_text(child, f"{label} sound")
        else:
            _validate_chart_hyperlink_extension_list(child, f"{label} extensions")


def _validate_chart_flat_text(element) -> None:
    label = "chart style flat text"
    if not set(element.attrib).issubset({"z"}):
        raise ValueError(f"{label} has invalid attributes")
    value = element.get("z")
    if value is not None:
        parsed = _parse_xsd_int64(value)
        if parsed is None or not -27273042329600 <= parsed <= 27273042316900:
            raise ValueError(f"{label} has an invalid z value")
    _validate_chart_drawing_leaf_text(element, label)


def _validate_chart_rotation(rotation, label: str) -> None:
    if set(rotation.attrib) != {"lat", "lon", "rev"}:
        raise ValueError(f"{label} has invalid attributes")
    for attribute_name in ("lat", "lon", "rev"):
        _validate_chart_int32_attribute(
            rotation,
            attribute_name,
            label,
            minimum=0,
            maximum=21599999,
        )
    _validate_chart_drawing_leaf_text(rotation, label)


def _validate_chart_camera(camera) -> None:
    label = "chart style 3D camera"
    if not set(camera.attrib).issubset({"prst", "fov", "zoom"}):
        raise ValueError(f"{label} has invalid attributes")
    if camera.get("prst") is None:
        raise ValueError(f"{label} is missing prst")
    _validate_chart_enum_attribute(
        camera,
        "prst",
        _CHART_PRESET_CAMERA_VALUES,
        label,
    )
    _validate_chart_int32_attribute(
        camera,
        "fov",
        label,
        minimum=0,
        maximum=10800000,
    )
    _validate_chart_int32_attribute(camera, "zoom", label, minimum=0)
    _validate_chart_composite_text(camera, label)
    children = _validate_chart_ordered_choice_children(
        camera,
        (_drawing_tags("rot"),),
        label,
    )
    if children:
        _validate_chart_rotation(children[0], f"{label} rotation")


def _validate_chart_light_rig(light_rig) -> None:
    label = "chart style 3D light rig"
    if set(light_rig.attrib) != {"rig", "dir"}:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_enum_attribute(
        light_rig,
        "rig",
        _CHART_LIGHT_RIG_VALUES,
        label,
    )
    _validate_chart_enum_attribute(
        light_rig,
        "dir",
        _CHART_LIGHT_RIG_DIRECTION_VALUES,
        label,
    )
    _validate_chart_composite_text(light_rig, label)
    children = _validate_chart_ordered_choice_children(
        light_rig,
        (_drawing_tags("rot"),),
        label,
    )
    if children:
        _validate_chart_rotation(children[0], f"{label} rotation")


def _validate_chart_backdrop(backdrop) -> None:
    label = "chart style 3D backdrop"
    if backdrop.attrib:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_composite_text(backdrop, label)
    children = _validate_chart_ordered_choice_children(
        backdrop,
        (
            _drawing_tags("anchor"),
            _drawing_tags("norm"),
            _drawing_tags("up"),
            _drawing_tags("extLst"),
        ),
        label,
    )
    required_tags = _drawing_tags("anchor", "norm", "up")
    if not required_tags.issubset({child.tag for child in children}):
        raise ValueError(f"{label} is missing required vectors")
    for child in children:
        local_name = etree.QName(child).localname
        if local_name == "extLst":
            _validate_chart_extension_list(
                child,
                f"{label} extension list",
            )
            continue
        attribute_names = (
            ("x", "y", "z")
            if local_name == "anchor"
            else ("dx", "dy", "dz")
        )
        if set(child.attrib) != set(attribute_names):
            raise ValueError(f"{label} {local_name} has invalid attributes")
        for attribute_name in attribute_names:
            _validate_chart_int64_attribute(
                child,
                attribute_name,
                f"{label} {local_name}",
                minimum=-27273042329600,
                maximum=27273042316900,
            )
        _validate_chart_drawing_leaf_text(child, f"{label} {local_name}")


def _validate_chart_scene3d(scene) -> None:
    label = "chart style 3D scene"
    if scene.attrib:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_composite_text(scene, label)
    children = _validate_chart_ordered_choice_children(
        scene,
        (
            _drawing_tags("camera"),
            _drawing_tags("lightRig"),
            _drawing_tags("backdrop"),
            _drawing_tags("extLst"),
        ),
        label,
    )
    required_tags = _drawing_tags("camera", "lightRig")
    if not required_tags.issubset({child.tag for child in children}):
        raise ValueError(f"{label} is missing required children")
    for child in children:
        local_name = etree.QName(child).localname
        if local_name == "camera":
            _validate_chart_camera(child)
        elif local_name == "lightRig":
            _validate_chart_light_rig(child)
        elif local_name == "backdrop":
            _validate_chart_backdrop(child)
        else:
            _validate_chart_extension_list(
                child,
                f"{label} extension list",
            )


def _validate_chart_shape3d(shape3d) -> None:
    label = "chart style 3D shape"
    allowed_attributes = {"z", "extrusionH", "contourW", "prstMaterial"}
    if not set(shape3d.attrib).issubset(allowed_attributes):
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_int64_attribute(
        shape3d,
        "z",
        label,
        minimum=-27273042329600,
        maximum=27273042316900,
    )
    for attribute_name in ("extrusionH", "contourW"):
        _validate_chart_int64_attribute(
            shape3d,
            attribute_name,
            label,
            minimum=0,
            maximum=0x7FFFFFFF,
        )
    _validate_chart_enum_attribute(
        shape3d,
        "prstMaterial",
        _CHART_MATERIAL_VALUES,
        label,
    )
    _validate_chart_composite_text(shape3d, label)
    children = _validate_chart_ordered_choice_children(
        shape3d,
        (
            _drawing_tags("bevelT"),
            _drawing_tags("bevelB"),
            _drawing_tags("extrusionClr"),
            _drawing_tags("contourClr"),
            _drawing_tags("extLst"),
        ),
        label,
    )
    for child in children:
        local_name = etree.QName(child).localname
        child_label = f"{label} {local_name}"
        if local_name in {"bevelT", "bevelB"}:
            if not set(child.attrib).issubset({"w", "h", "prst"}):
                raise ValueError(f"{child_label} has invalid attributes")
            for attribute_name in ("w", "h"):
                _validate_chart_int64_attribute(
                    child,
                    attribute_name,
                    child_label,
                    minimum=0,
                    maximum=0x7FFFFFFF,
                )
            _validate_chart_enum_attribute(
                child,
                "prst",
                _CHART_BEVEL_PRESET_VALUES,
                child_label,
            )
            _validate_chart_drawing_leaf_text(child, child_label)
        elif local_name in {"extrusionClr", "contourClr"}:
            _validate_chart_color_wrapper(child, child_label)
        else:
            _validate_chart_extension_list(
                child,
                f"{label} extension list",
            )


def _validate_chart_effect_color_child(element, label: str) -> None:
    children = _chart_schema_children(element)
    if len(children) != 1 or children[0].tag not in _THEME_COLOR_MODEL_TAGS:
        raise ValueError(f"{label} has invalid color content")
    _validate_chart_composite_text(element, label)
    _validate_chart_style_color_model(children[0], label)


def _validate_chart_effect_shadow(
    effect,
    label: str,
    *,
    attributes: frozenset[str],
    required_attributes: frozenset[str] = frozenset(),
) -> None:
    if not set(effect.attrib).issubset(attributes):
        raise ValueError(f"{label} has invalid attributes")
    for attribute_name in required_attributes:
        if effect.get(attribute_name) is None:
            raise ValueError(f"{label} is missing {attribute_name}")
    for attribute_name in ("blurRad", "dist"):
        if attribute_name in attributes:
            _validate_chart_int64_attribute(
                effect,
                attribute_name,
                label,
                minimum=0,
                maximum=0x7FFFFFFF,
            )
    for attribute_name in ("dir", "fadeDir"):
        if attribute_name in attributes:
            _validate_chart_int32_attribute(
                effect,
                attribute_name,
                label,
                minimum=0,
                maximum=21599999,
            )
    for attribute_name in ("sx", "sy"):
        if attribute_name in attributes:
            _validate_chart_int32_attribute(effect, attribute_name, label)
    for attribute_name in ("kx", "ky"):
        if attribute_name in attributes:
            _validate_chart_int32_attribute(
                effect,
                attribute_name,
                label,
                minimum=-5399999,
                maximum=5399999,
            )
    if "algn" in attributes:
        _validate_chart_enum_attribute(
            effect,
            "algn",
            _CHART_RECTANGLE_ALIGNMENT_VALUES,
            label,
        )
    if "rotWithShape" in attributes:
        _validate_chart_boolean_attributes(effect, ("rotWithShape",), label)
    _validate_chart_effect_color_child(effect, label)


def _validate_chart_effect_list(effect_list) -> None:
    label = "chart style effect list"
    if effect_list.attrib:
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_composite_text(effect_list, label)
    children = _validate_chart_ordered_choice_children(
        effect_list,
        (
            _drawing_tags("blur"),
            _drawing_tags("fillOverlay"),
            _drawing_tags("glow"),
            _drawing_tags("innerShdw"),
            _drawing_tags("outerShdw"),
            _drawing_tags("prstShdw"),
            _drawing_tags("reflection"),
            _drawing_tags("softEdge"),
        ),
        label,
    )
    for child in children:
        local_name = etree.QName(child).localname
        child_label = f"{label} {local_name}"
        if local_name == "blur":
            if not set(child.attrib).issubset({"rad", "grow"}):
                raise ValueError(f"{child_label} has invalid attributes")
            _validate_chart_int64_attribute(
                child,
                "rad",
                child_label,
                minimum=0,
                maximum=0x7FFFFFFF,
            )
            _validate_chart_boolean_attributes(child, ("grow",), child_label)
            _validate_chart_drawing_leaf_text(child, child_label)
        elif local_name == "fillOverlay":
            if set(child.attrib) != {"blend"}:
                raise ValueError(f"{child_label} has invalid attributes")
            _validate_chart_enum_attribute(
                child,
                "blend",
                _CHART_BLEND_MODE_VALUES,
                child_label,
            )
            fill_children = _chart_schema_children(child)
            if len(fill_children) != 1:
                raise ValueError(f"{child_label} has invalid fill content")
            _validate_chart_composite_text(child, child_label)
            _validate_chart_effect_fill(fill_children[0], child_label)
        elif local_name == "glow":
            if not set(child.attrib).issubset({"rad"}):
                raise ValueError(f"{child_label} has invalid attributes")
            _validate_chart_int64_attribute(
                child,
                "rad",
                child_label,
                minimum=0,
                maximum=0x7FFFFFFF,
            )
            _validate_chart_effect_color_child(child, child_label)
        elif local_name == "innerShdw":
            _validate_chart_effect_shadow(
                child,
                child_label,
                attributes=frozenset({"blurRad", "dist", "dir"}),
            )
        elif local_name == "outerShdw":
            _validate_chart_effect_shadow(
                child,
                child_label,
                attributes=frozenset(
                    {
                        "blurRad", "dist", "dir", "sx", "sy", "kx", "ky",
                        "algn", "rotWithShape",
                    }
                ),
            )
        elif local_name == "prstShdw":
            if not set(child.attrib).issubset({"prst", "dist", "dir"}):
                raise ValueError(f"{child_label} has invalid attributes")
            if child.get("prst") is None:
                raise ValueError(f"{child_label} is missing prst")
            _validate_chart_enum_attribute(
                child,
                "prst",
                _CHART_PRESET_SHADOW_VALUES,
                child_label,
            )
            _validate_chart_int64_attribute(
                child,
                "dist",
                child_label,
                minimum=0,
                maximum=0x7FFFFFFF,
            )
            _validate_chart_int32_attribute(
                child,
                "dir",
                child_label,
                minimum=0,
                maximum=21599999,
            )
            _validate_chart_effect_color_child(child, child_label)
        elif local_name == "reflection":
            allowed = {
                "blurRad", "stA", "stPos", "endA", "endPos", "dist", "dir",
                "fadeDir", "sx", "sy", "kx", "ky", "algn", "rotWithShape",
            }
            if not set(child.attrib).issubset(allowed):
                raise ValueError(f"{child_label} has invalid attributes")
            for attribute_name in ("stA", "stPos", "endA", "endPos"):
                _validate_chart_int32_attribute(
                    child,
                    attribute_name,
                    child_label,
                    minimum=0,
                    maximum=100000,
                )
            for attribute_name in ("blurRad", "dist"):
                _validate_chart_int64_attribute(
                    child,
                    attribute_name,
                    child_label,
                    minimum=0,
                    maximum=0x7FFFFFFF,
                )
            for attribute_name in ("dir", "fadeDir"):
                _validate_chart_int32_attribute(
                    child,
                    attribute_name,
                    child_label,
                    minimum=0,
                    maximum=21599999,
                )
            for attribute_name in ("sx", "sy"):
                _validate_chart_int32_attribute(child, attribute_name, child_label)
            for attribute_name in ("kx", "ky"):
                _validate_chart_int32_attribute(
                    child,
                    attribute_name,
                    child_label,
                    minimum=-5399999,
                    maximum=5399999,
                )
            _validate_chart_enum_attribute(
                child,
                "algn",
                _CHART_RECTANGLE_ALIGNMENT_VALUES,
                child_label,
            )
            _validate_chart_boolean_attributes(
                child,
                ("rotWithShape",),
                child_label,
            )
            _validate_chart_drawing_leaf_text(child, child_label)
        elif local_name == "softEdge":
            if set(child.attrib) != {"rad"}:
                raise ValueError(f"{child_label} has invalid attributes")
            _validate_chart_int64_attribute(
                child,
                "rad",
                child_label,
                minimum=0,
                maximum=0x7FFFFFFF,
            )
            _validate_chart_drawing_leaf_text(child, child_label)


def _validate_chart_effect_fill(fill, label: str) -> None:
    if fill.tag == f"{{{_THEME_NAMESPACE}}}noFill":
        _validate_chart_empty_drawing_leaf(fill, f"{label} no-fill")
    elif fill.tag == f"{{{_THEME_NAMESPACE}}}solidFill":
        _validate_chart_solid_fill(fill)
    elif fill.tag == f"{{{_THEME_NAMESPACE}}}gradFill":
        _validate_chart_gradient_fill(fill)
    elif fill.tag == f"{{{_THEME_NAMESPACE}}}pattFill":
        _validate_chart_pattern_fill(fill)
    elif fill.tag == f"{{{_THEME_NAMESPACE}}}blipFill":
        _validate_chart_blip_fill(fill)
    elif fill.tag == f"{{{_THEME_NAMESPACE}}}grpFill":
        _validate_chart_empty_drawing_leaf(fill, f"{label} group-fill")
    else:
        raise ValueError(f"{label} has an invalid fill element")
    _validate_chart_composite_text(fill, label)


def _validate_chart_color_container(container, expected_name: str, label: str) -> None:
    if etree.QName(container).localname != expected_name or container.attrib:
        raise ValueError(f"{label} has invalid color container")
    children = _chart_schema_children(container)
    if len(children) != 1 or children[0].tag not in _THEME_COLOR_MODEL_TAGS:
        raise ValueError(f"{label} has invalid color content")
    _validate_chart_composite_text(container, label)
    _validate_chart_style_color_model(children[0], label)


def _validate_chart_blip_effect(effect, label: str) -> None:
    """校验 blip 可直接承载的 17 种 DrawingML effect。"""
    local_name = etree.QName(effect).localname
    if local_name == "alphaBiLevel":
        if set(effect.attrib) != {"thresh"}:
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_int32_attribute(effect, "thresh", label, minimum=0, maximum=100000)
        _validate_chart_drawing_leaf_text(effect, label)
    elif local_name in {"alphaCeiling", "alphaFloor", "grayscl"}:
        _validate_chart_empty_drawing_leaf(effect, label)
    elif local_name == "alphaInv":
        if effect.attrib:
            raise ValueError(f"{label} has invalid attributes")
        children = _chart_schema_children(effect)
        if len(children) > 1 or (children and children[0].tag not in _THEME_COLOR_MODEL_TAGS):
            raise ValueError(f"{label} has invalid color content")
        _validate_chart_composite_text(effect, label)
        if children:
            _validate_chart_style_color_model(children[0], label)
    elif local_name == "alphaMod":
        if effect.attrib:
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_composite_text(effect, label)
        children = _validate_chart_ordered_choice_children(effect, (_drawing_tags("cont"),), label)
        if len(children) != 1:
            raise ValueError(f"{label} is missing cont")
        _validate_chart_effect_container(children[0], f"{label} container")
    elif local_name == "alphaModFix":
        if not set(effect.attrib).issubset({"amt"}):
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_int32_attribute(effect, "amt", label, minimum=0)
        _validate_chart_drawing_leaf_text(effect, label)
    elif local_name == "alphaRepl":
        if set(effect.attrib) != {"a"}:
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_int32_attribute(effect, "a", label, minimum=0, maximum=100000)
        _validate_chart_drawing_leaf_text(effect, label)
    elif local_name == "biLevel":
        if set(effect.attrib) != {"thresh"}:
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_int32_attribute(effect, "thresh", label, minimum=0, maximum=100000)
        _validate_chart_drawing_leaf_text(effect, label)
    elif local_name == "blur":
        if not set(effect.attrib).issubset({"rad", "grow"}):
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_int64_attribute(effect, "rad", label, minimum=0, maximum=0x7FFFFFFF)
        _validate_chart_boolean_attributes(effect, ("grow",), label)
        _validate_chart_drawing_leaf_text(effect, label)
    elif local_name == "clrChange":
        if not set(effect.attrib).issubset({"useA"}):
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_boolean_attributes(effect, ("useA",), label)
        _validate_chart_composite_text(effect, label)
        children = _validate_chart_ordered_choice_children(
            effect,
            (_drawing_tags("clrFrom"), _drawing_tags("clrTo")),
            label,
        )
        if len(children) != 2:
            raise ValueError(f"{label} is missing color endpoints")
        _validate_chart_color_container(children[0], "clrFrom", label)
        _validate_chart_color_container(children[1], "clrTo", label)
    elif local_name == "clrRepl":
        if effect.attrib:
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_effect_color_child(effect, label)
    elif local_name == "duotone":
        if effect.attrib:
            raise ValueError(f"{label} has invalid attributes")
        children = _chart_schema_children(effect)
        if len(children) != 2 or any(child.tag not in _THEME_COLOR_MODEL_TAGS for child in children):
            raise ValueError(f"{label} has invalid color content")
        _validate_chart_composite_text(effect, label)
        for child in children:
            _validate_chart_style_color_model(child, label)
    elif local_name == "fillOverlay":
        if set(effect.attrib) != {"blend"}:
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_enum_attribute(effect, "blend", _CHART_BLEND_MODE_VALUES, label)
        children = _chart_schema_children(effect)
        if len(children) != 1:
            raise ValueError(f"{label} has invalid fill content")
        _validate_chart_composite_text(effect, label)
        _validate_chart_effect_fill(children[0], label)
    elif local_name == "hsl":
        if not set(effect.attrib).issubset({"hue", "sat", "lum"}):
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_int32_attribute(effect, "hue", label, minimum=0, maximum=21599999)
        for attribute_name in ("sat", "lum"):
            _validate_chart_int32_attribute(effect, attribute_name, label, minimum=-100000, maximum=100000)
        _validate_chart_drawing_leaf_text(effect, label)
    elif local_name == "lum":
        if not set(effect.attrib).issubset({"bright", "contrast"}):
            raise ValueError(f"{label} has invalid attributes")
        for attribute_name in ("bright", "contrast"):
            _validate_chart_int32_attribute(effect, attribute_name, label, minimum=-100000, maximum=100000)
        _validate_chart_drawing_leaf_text(effect, label)
    elif local_name == "tint":
        if not set(effect.attrib).issubset({"hue", "amt"}):
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_int32_attribute(effect, "hue", label, minimum=0, maximum=21599999)
        _validate_chart_int32_attribute(effect, "amt", label, minimum=-100000, maximum=100000)
        _validate_chart_drawing_leaf_text(effect, label)
    else:
        raise ValueError(f"{label} has an invalid blip effect")


def _validate_chart_effect_container(container, label: str) -> None:
    if not set(container.attrib).issubset({"type", "name"}):
        raise ValueError(f"{label} has invalid attributes")
    _validate_chart_enum_attribute(container, "type", frozenset({"sib", "tree"}), label)
    _validate_chart_token_attribute(container, "name", label)
    _validate_chart_composite_text(container, label)
    allowed_names = _CHART_BLIP_EFFECT_NAMES | {
        "cont", "effect", "alphaOutset", "blend", "fill", "glow", "innerShdw",
        "outerShdw", "prstShdw", "reflection", "relOff", "softEdge", "xfrm",
    }
    allowed = _drawing_tags(*allowed_names)
    for child in _chart_schema_children(container):
        if child.tag not in allowed:
            raise ValueError(f"{label} has invalid children")
        _validate_chart_effect_node(child, f"{label} {etree.QName(child).localname}")


def _validate_chart_effect_node(effect, label: str) -> None:
    local_name = etree.QName(effect).localname
    if local_name in _CHART_BLIP_EFFECT_NAMES:
        _validate_chart_blip_effect(effect, label)
        return
    if local_name == "cont":
        _validate_chart_effect_container(effect, label)
        return
    if local_name == "effect":
        if not set(effect.attrib).issubset({"ref"}):
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_token_attribute(effect, "ref", label)
        _validate_chart_drawing_leaf_text(effect, label)
        return
    if local_name == "alphaOutset":
        if not set(effect.attrib).issubset({"rad"}):
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_int64_attribute(effect, "rad", label, minimum=-27273042329600, maximum=27273042316900)
        _validate_chart_drawing_leaf_text(effect, label)
        return
    if local_name == "blend":
        if set(effect.attrib) != {"blend"}:
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_enum_attribute(effect, "blend", _CHART_BLEND_MODE_VALUES, label)
        _validate_chart_composite_text(effect, label)
        children = _validate_chart_ordered_choice_children(effect, (_drawing_tags("cont"),), label)
        if len(children) != 1:
            raise ValueError(f"{label} is missing cont")
        _validate_chart_effect_container(children[0], f"{label} container")
        return
    if local_name == "fill":
        if effect.attrib:
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_composite_text(effect, label)
        children = _chart_schema_children(effect)
        if len(children) != 1:
            raise ValueError(f"{label} has invalid fill content")
        _validate_chart_effect_fill(children[0], label)
        return
    if local_name == "glow":
        if not set(effect.attrib).issubset({"rad"}):
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_int64_attribute(effect, "rad", label, minimum=0, maximum=0x7FFFFFFF)
        _validate_chart_effect_color_child(effect, label)
        return
    if local_name in {"innerShdw", "outerShdw", "prstShdw", "reflection"}:
        # Reuse the strict effect-list implementations by validating through a
        # one-child effect list where possible; reflection is handled inline.
        if local_name == "reflection":
            # The direct effect-list validator has no reflection color child;
            # invoke its attribute logic explicitly below.
            allowed = {
                "blurRad", "stA", "stPos", "endA", "endPos", "dist", "dir",
                "fadeDir", "sx", "sy", "kx", "ky", "algn", "rotWithShape",
            }
            if not set(effect.attrib).issubset(allowed):
                raise ValueError(f"{label} has invalid attributes")
            for attribute_name in ("blurRad", "dist"):
                _validate_chart_int64_attribute(effect, attribute_name, label, minimum=0, maximum=0x7FFFFFFF)
            for attribute_name in ("stA", "stPos", "endA", "endPos"):
                _validate_chart_int32_attribute(effect, attribute_name, label, minimum=0, maximum=100000)
            for attribute_name in ("dir", "fadeDir"):
                _validate_chart_int32_attribute(effect, attribute_name, label, minimum=0, maximum=21599999)
            for attribute_name in ("sx", "sy"):
                _validate_chart_int32_attribute(effect, attribute_name, label)
            for attribute_name in ("kx", "ky"):
                _validate_chart_int32_attribute(effect, attribute_name, label, minimum=-5399999, maximum=5399999)
            _validate_chart_enum_attribute(effect, "algn", _CHART_RECTANGLE_ALIGNMENT_VALUES, label)
            _validate_chart_boolean_attributes(effect, ("rotWithShape",), label)
            _validate_chart_drawing_leaf_text(effect, label)
            return
        wrapper = etree.Element(f"{{{_THEME_NAMESPACE}}}effectLst")
        wrapper.append(deepcopy(effect))
        _validate_chart_effect_list(wrapper)
        return
    if local_name == "relOff":
        if not set(effect.attrib).issubset({"tx", "ty"}):
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_int32_attribute(effect, "tx", label)
        _validate_chart_int32_attribute(effect, "ty", label)
        _validate_chart_drawing_leaf_text(effect, label)
        return
    if local_name == "softEdge":
        if set(effect.attrib) != {"rad"}:
            raise ValueError(f"{label} has invalid attributes")
        _validate_chart_int64_attribute(effect, "rad", label, minimum=0, maximum=0x7FFFFFFF)
        _validate_chart_drawing_leaf_text(effect, label)
        return
    if local_name == "xfrm":
        allowed = {"sx", "sy", "kx", "ky", "tx", "ty"}
        if not set(effect.attrib).issubset(allowed):
            raise ValueError(f"{label} has invalid attributes")
        for attribute_name in ("sx", "sy"):
            _validate_chart_int32_attribute(effect, attribute_name, label)
        for attribute_name in ("kx", "ky"):
            _validate_chart_int32_attribute(effect, attribute_name, label, minimum=-5399999, maximum=5399999)
        for attribute_name in ("tx", "ty"):
            _validate_chart_int64_attribute(effect, attribute_name, label, minimum=-27273042329600, maximum=27273042316900)
        _validate_chart_drawing_leaf_text(effect, label)
        return
    raise ValueError(f"{label} has an invalid effect")


def _validate_chart_drawing_subtree(element) -> None:
    no_fill_tag = f"{{{_THEME_NAMESPACE}}}noFill"
    solid_fill_tag = f"{{{_THEME_NAMESPACE}}}solidFill"
    gradient_fill_tag = f"{{{_THEME_NAMESPACE}}}gradFill"
    blip_fill_tag = f"{{{_THEME_NAMESPACE}}}blipFill"
    pattern_fill_tag = f"{{{_THEME_NAMESPACE}}}pattFill"
    group_fill_tag = f"{{{_THEME_NAMESPACE}}}grpFill"
    preset_geometry_tag = f"{{{_THEME_NAMESPACE}}}prstGeom"
    custom_geometry_tag = f"{{{_THEME_NAMESPACE}}}custGeom"
    transform_2d_tag = f"{{{_THEME_NAMESPACE}}}xfrm"
    effect_list_tag = f"{{{_THEME_NAMESPACE}}}effectLst"
    effect_dag_tag = f"{{{_THEME_NAMESPACE}}}effectDag"
    scene3d_tag = f"{{{_THEME_NAMESPACE}}}scene3d"
    shape3d_tag = f"{{{_THEME_NAMESPACE}}}sp3d"
    line_tag = f"{{{_THEME_NAMESPACE}}}ln"
    preset_dash_tag = f"{{{_THEME_NAMESPACE}}}prstDash"
    for child in _chart_schema_children(element):
        if child.tag == _DRAWING_EXTENSION_TAG:
            continue
        if child.tag == no_fill_tag:
            _validate_chart_empty_drawing_leaf(
                child,
                "chart style no-fill properties",
            )
            continue
        if child.tag == solid_fill_tag:
            _validate_chart_solid_fill(child)
            continue
        if child.tag == gradient_fill_tag:
            _validate_chart_gradient_fill(child)
            continue
        if child.tag == blip_fill_tag:
            _validate_chart_blip_fill(child)
            continue
        if child.tag == pattern_fill_tag:
            _validate_chart_pattern_fill(child)
            continue
        if child.tag == group_fill_tag:
            _validate_chart_empty_drawing_leaf(
                child,
                "chart style group-fill properties",
            )
            continue
        if child.tag == preset_geometry_tag:
            _validate_chart_preset_geometry(child)
            continue
        if child.tag == custom_geometry_tag:
            _validate_chart_custom_geometry(child)
            continue
        if child.tag == transform_2d_tag:
            _validate_chart_transform_2d(child)
            continue
        if child.tag == effect_list_tag:
            _validate_chart_effect_list(child)
            continue
        if child.tag == effect_dag_tag:
            _validate_chart_effect_container(child, "chart style effect DAG")
            continue
        if child.tag == scene3d_tag:
            _validate_chart_scene3d(child)
            continue
        if child.tag == shape3d_tag:
            _validate_chart_shape3d(child)
            continue
        if child.tag == line_tag:
            _validate_chart_line_properties(child)
        elif child.tag == preset_dash_tag:
            _validate_chart_preset_dash(child)
        _validate_chart_drawing_subtree(child)


def _validate_chart_shape_properties(shape_properties) -> None:
    if not set(shape_properties.attrib).issubset({"bwMode"}):
        raise ValueError("chart style shape properties have invalid attributes")
    _validate_chart_enum_attribute(
        shape_properties,
        "bwMode",
        _CHART_BLACK_WHITE_MODE_VALUES,
        "chart style shape properties",
    )
    groups = (
        _drawing_tags("xfrm"),
        _drawing_tags("custGeom", "prstGeom"),
        _drawing_tags(
            "noFill",
            "solidFill",
            "gradFill",
            "blipFill",
            "pattFill",
            "grpFill",
        ),
        _drawing_tags("ln"),
        _drawing_tags("effectLst", "effectDag"),
        _drawing_tags("scene3d"),
        _drawing_tags("sp3d"),
        _drawing_tags("extLst"),
    )
    children = _validate_chart_ordered_choice_children(
        shape_properties,
        groups,
        "chart style shape properties",
    )
    if children and children[-1].tag == f"{{{_THEME_NAMESPACE}}}extLst":
        _validate_chart_extension_list(
            children[-1],
            "chart style shape-properties extension list",
            uri_required=True,
        )
    _validate_chart_drawing_subtree(shape_properties)


def _validate_chart_text_character_properties(properties) -> None:
    allowed_attributes = {
        "kumimoji", "lang", "altLang", "sz", "b", "i", "u",
        "strike", "kern", "cap", "spc", "normalizeH", "baseline",
        "noProof", "dirty", "err", "smtClean", "smtId", "bmk",
    }
    if not set(properties.attrib).issubset(allowed_attributes):
        raise ValueError(
            "chart style text character properties have invalid attributes"
        )
    label = "chart style text character properties"
    _validate_chart_boolean_attributes(
        properties,
        (
            "kumimoji", "b", "i", "normalizeH", "noProof", "dirty",
            "err", "smtClean",
        ),
        label,
    )
    for attribute_name, minimum, maximum in (
        ("sz", 100, 400000),
        ("kern", 0, 400000),
        ("spc", -400000, 400000),
        ("baseline", -0x80000000, 0x7FFFFFFF),
    ):
        _validate_chart_int32_attribute(
            properties,
            attribute_name,
            label,
            minimum=minimum,
            maximum=maximum,
        )
    smart_tag_id = properties.get("smtId")
    if smart_tag_id is not None and _parse_xsd_uint32(smart_tag_id) is None:
        raise ValueError(f"{label} has an invalid smtId value")
    for attribute_name, values in (
        ("u", _CHART_TEXT_UNDERLINE_VALUES),
        ("strike", _CHART_TEXT_STRIKE_VALUES),
        ("cap", _CHART_TEXT_CAPS_VALUES),
    ):
        _validate_chart_enum_attribute(
            properties,
            attribute_name,
            values,
            label,
        )
    groups = (
        _drawing_tags("ln"),
        _drawing_tags(
            "noFill",
            "solidFill",
            "gradFill",
            "blipFill",
            "pattFill",
            "grpFill",
        ),
        _drawing_tags("effectLst", "effectDag"),
        _drawing_tags("highlight"),
        _drawing_tags("uLnTx", "uLn"),
        _drawing_tags("uFillTx", "uFill"),
        _drawing_tags("latin"),
        _drawing_tags("ea"),
        _drawing_tags("cs"),
        _drawing_tags("sym"),
        _drawing_tags("hlinkClick"),
        _drawing_tags("hlinkMouseOver"),
        _drawing_tags("rtl"),
        _drawing_tags("extLst"),
    )
    children = _validate_chart_ordered_choice_children(
        properties,
        groups,
        "chart style text character properties",
    )
    if children and children[-1].tag == f"{{{_THEME_NAMESPACE}}}extLst":
        _validate_chart_extension_list(
            children[-1],
            "chart style text-character extension list",
        )
    for child in children:
        local_name = etree.QName(child).localname
        if local_name == "highlight":
            _validate_chart_color_wrapper(child, "chart style highlight")
        elif local_name in {"uLnTx", "uFillTx"}:
            _validate_chart_empty_drawing_leaf(
                child,
                f"chart style {local_name}",
            )
        elif local_name == "uLn":
            _validate_chart_line_properties(child)
        elif local_name == "uFill":
            if child.attrib:
                raise ValueError("chart style underline fill has invalid attributes")
            _validate_chart_composite_text(child, "chart style underline fill")
            fill_children = _chart_schema_children(child)
            if len(fill_children) != 1:
                raise ValueError("chart style underline fill has invalid content")
            _validate_chart_effect_fill(fill_children[0], "chart style underline fill")
        elif local_name in {"latin", "ea", "cs", "sym"}:
            _validate_chart_text_font(child)
        elif local_name == "rtl":
            if not set(child.attrib).issubset({"val"}):
                raise ValueError("chart style rtl has invalid attributes")
            _validate_chart_boolean_attributes(child, ("val",), "chart style rtl")
            _validate_chart_drawing_leaf_text(child, "chart style rtl")
        elif local_name in {"hlinkClick", "hlinkMouseOver"}:
            _validate_chart_hyperlink(child)
    _validate_chart_drawing_subtree(properties)


def _validate_chart_text_body_properties(properties) -> None:
    allowed_attributes = {
        "rot", "spcFirstLastPara", "vertOverflow", "horzOverflow",
        "vert", "wrap", "lIns", "tIns", "rIns", "bIns", "numCol",
        "spcCol", "rtlCol", "fromWordArt", "anchor", "anchorCtr",
        "forceAA", "upright", "compatLnSpc",
    }
    if not set(properties.attrib).issubset(allowed_attributes):
        raise ValueError("chart style text body properties have invalid attributes")
    label = "chart style text body properties"
    _validate_chart_composite_text(properties, label)
    _validate_chart_boolean_attributes(
        properties,
        (
            "spcFirstLastPara", "rtlCol", "fromWordArt", "anchorCtr",
            "forceAA", "upright", "compatLnSpc",
        ),
        label,
    )
    for attribute_name, minimum, maximum in (
        ("rot", -0x80000000, 0x7FFFFFFF),
        ("lIns", -0x80000000, 0x7FFFFFFF),
        ("tIns", -0x80000000, 0x7FFFFFFF),
        ("rIns", -0x80000000, 0x7FFFFFFF),
        ("bIns", -0x80000000, 0x7FFFFFFF),
        ("numCol", 1, 16),
        ("spcCol", 0, 0x7FFFFFFF),
    ):
        _validate_chart_int32_attribute(
            properties,
            attribute_name,
            label,
            minimum=minimum,
            maximum=maximum,
        )
    for attribute_name, values in (
        ("vertOverflow", _CHART_TEXT_VERTICAL_OVERFLOW_VALUES),
        ("horzOverflow", _CHART_TEXT_HORIZONTAL_OVERFLOW_VALUES),
        ("vert", _CHART_TEXT_VERTICAL_VALUES),
        ("wrap", _CHART_TEXT_WRAPPING_VALUES),
        ("anchor", _CHART_TEXT_ANCHORING_VALUES),
    ):
        _validate_chart_enum_attribute(
            properties,
            attribute_name,
            values,
            label,
        )
    groups = (
        _drawing_tags("prstTxWarp"),
        _drawing_tags("noAutofit", "normAutofit", "spAutoFit"),
        _drawing_tags("scene3d"),
        _drawing_tags("sp3d", "flatTx"),
        _drawing_tags("extLst"),
    )
    children = _validate_chart_ordered_choice_children(
        properties,
        groups,
        "chart style text body properties",
    )
    if children and children[-1].tag == f"{{{_THEME_NAMESPACE}}}extLst":
        _validate_chart_extension_list(
            children[-1],
            "chart style text-body extension list",
        )
    for child in children:
        if child.tag == f"{{{_THEME_NAMESPACE}}}prstTxWarp":
            _validate_chart_preset_text_warp(child)
        elif child.tag in {
            f"{{{_THEME_NAMESPACE}}}noAutofit",
            f"{{{_THEME_NAMESPACE}}}normAutofit",
            f"{{{_THEME_NAMESPACE}}}spAutoFit",
        }:
            _validate_chart_text_autofit(child)
        elif child.tag == f"{{{_THEME_NAMESPACE}}}flatTx":
            _validate_chart_flat_text(child)
    _validate_chart_drawing_subtree(properties)


def _validate_chart_color_transforms(parent, label: str) -> None:
    int32_min = -0x80000000
    int32_max = 0x7FFFFFFF
    bounded_values = {
        "tint": (0, 100000),
        "shade": (0, 100000),
        "alpha": (0, 100000),
        "alphaOff": (-100000, 100000),
        "alphaMod": (0, int32_max),
        "hue": (0, 21599999),
        "hueOff": (int32_min, int32_max),
        "hueMod": (0, int32_max),
    }
    for name in (
        "sat", "satOff", "satMod", "lum", "lumOff", "lumMod",
        "red", "redOff", "redMod", "green", "greenOff", "greenMod",
        "blue", "blueOff", "blueMod",
    ):
        bounded_values[name] = (int32_min, int32_max)
    bounded_values["lum"] = (int32_min, 100000)

    for transform in _chart_schema_children(parent):
        transform_name = etree.QName(transform).localname
        transform_text = _chart_leaf_text(transform)
        if (
            transform.tag
            != f"{{{_THEME_NAMESPACE}}}{transform_name}"
            or transform_name not in _CHART_COLOR_TRANSFORM_NAMES
            or transform_text is None
            or transform_text.strip()
        ):
            raise ValueError(f"{label} has an invalid color transform")
        if transform_name in _CHART_EMPTY_COLOR_TRANSFORMS:
            if transform.attrib:
                raise ValueError(
                    f"{label} has an invalid empty color transform"
                )
            continue
        if set(transform.attrib) != {"val"}:
            raise ValueError(f"{label} has a color transform without a value")
        normalized = _normalize_ooxml_integer(
            transform.get("val"),
            max_digits=10,
        )
        if normalized is None:
            raise ValueError(f"{label} has an invalid color-transform value")
        value = int(normalized)
        minimum, maximum = bounded_values[transform_name]
        if not minimum <= value <= maximum:
            raise ValueError(f"{label} has an out-of-range color transform")


def _validate_chart_style_color_model(model, label: str) -> None:
    _validate_drawing_color_model(model, label)
    allowed_attributes = {
        f"{{{_THEME_NAMESPACE}}}scrgbClr": {"r", "g", "b"},
        f"{{{_THEME_NAMESPACE}}}srgbClr": {
            "val",
            _LEGACY_SPREADSHEET_COLOR_INDEX_ATTRIBUTE,
        },
        f"{{{_THEME_NAMESPACE}}}hslClr": {"hue", "sat", "lum"},
        f"{{{_THEME_NAMESPACE}}}sysClr": {"val", "lastClr"},
        f"{{{_THEME_NAMESPACE}}}schemeClr": {"val"},
        f"{{{_THEME_NAMESPACE}}}prstClr": {"val"},
    }[model.tag]
    if not set(model.attrib).issubset(allowed_attributes):
        raise ValueError(f"{label} has unknown color attributes")

    scrgb_tag = f"{{{_THEME_NAMESPACE}}}scrgbClr"
    hsl_tag = f"{{{_THEME_NAMESPACE}}}hslClr"
    srgb_tag = f"{{{_THEME_NAMESPACE}}}srgbClr"
    system_tag = f"{{{_THEME_NAMESPACE}}}sysClr"
    scheme_tag = f"{{{_THEME_NAMESPACE}}}schemeClr"
    preset_tag = f"{{{_THEME_NAMESPACE}}}prstClr"
    if model.tag == scrgb_tag and any(
        _parse_xsd_int32(model.get(channel)) is None
        for channel in ("r", "g", "b")
    ):
        raise ValueError(f"{label} has an invalid scRGB value")
    if model.tag == hsl_tag:
        hue = _parse_xsd_int32(model.get("hue"))
        if (
            hue is None
            or not 0 <= hue < 21600000
            or _parse_xsd_int32(model.get("sat")) is None
            or _parse_xsd_int32(model.get("lum")) is None
        ):
            raise ValueError(f"{label} has an invalid HSL value")
    if model.tag == srgb_tag:
        legacy_color_index = model.get(
            _LEGACY_SPREADSHEET_COLOR_INDEX_ATTRIBUTE
        )
        parsed_legacy_color_index = (
            _parse_xsd_int32(legacy_color_index)
            if legacy_color_index is not None
            else None
        )
        if legacy_color_index is not None and (
            parsed_legacy_color_index is None
            or not 0 <= parsed_legacy_color_index <= 80
        ):
            raise ValueError(
                f"{label} has an invalid legacy spreadsheet color index"
            )
    enum_values = {
        system_tag: _CHART_SYSTEM_COLOR_VALUES,
        scheme_tag: _CHART_SCHEME_COLOR_VALUES,
        preset_tag: _CHART_PRESET_COLOR_VALUES,
    }
    if model.tag in enum_values and model.get("val") not in enum_values[
        model.tag
    ]:
        raise ValueError(f"{label} has an invalid color value")
    _validate_chart_color_transforms(model, label)


def _validate_chart_style_reference(reference, *, font: bool) -> None:
    allowed_attributes = {"idx", "mods"}
    if not set(reference.attrib).issubset(allowed_attributes):
        raise ValueError("chart style reference has unknown attributes")
    if "mods" in reference.attrib and not _has_xsd_list_items(
        reference.get("mods")
    ):
        raise ValueError("chart style reference has invalid modifiers")
    index = reference.get("idx")
    if font:
        if index not in {"major", "minor", "none"}:
            raise ValueError("chart style font reference has an invalid index")
    elif _parse_bounded_decimal(index, 0xFFFFFFFF) is None:
        raise ValueError("chart style reference has an invalid index")

    style_color_tag = f"{{{_CHART_STYLE_NAMESPACE}}}styleClr"
    state = "color"
    for child in _chart_schema_children(reference):
        if child.tag in _THEME_COLOR_MODEL_TAGS and state == "color":
            _validate_chart_style_color_model(child, "chart style reference")
            state = "style"
            continue
        if child.tag == style_color_tag and state in {"color", "style"}:
            if not set(child.attrib).issubset({"val"}):
                raise ValueError("chart style color has unknown attributes")
            _validate_chart_color_transforms(child, "chart style color")
            state = "done"
            continue
        raise ValueError("chart style reference has invalid children")


def _validate_chart_style_entry(entry) -> None:
    if not set(entry.attrib).issubset({"mods"}):
        raise ValueError("chart style entry has unknown attributes")
    if "mods" in entry.attrib and not _has_xsd_list_items(
        entry.get("mods")
    ):
        raise ValueError("chart style entry has invalid modifiers")
    order = {
        f"{{{_CHART_STYLE_NAMESPACE}}}{name}": index
        for index, name in enumerate(_CHART_STYLE_ENTRY_CHILDREN)
    }
    children = _chart_schema_children(entry)
    observed = [child.tag for child in children]
    if (
        any(tag not in order for tag in observed)
        or len(observed) != len(set(observed))
        or [order[tag] for tag in observed]
        != sorted(order[tag] for tag in observed)
    ):
        raise ValueError("chart style entry has invalid children")
    required = {
        f"{{{_CHART_STYLE_NAMESPACE}}}{name}"
        for name in _CHART_STYLE_ENTRY_CHILDREN
        if name not in _CHART_STYLE_ENTRY_OPTIONAL_CHILDREN
    }
    if not required.issubset(observed):
        raise ValueError("chart style entry is missing required references")

    for reference_name in ("lnRef", "fillRef", "effectRef"):
        _validate_chart_style_reference(
            entry.find(f"{{{_CHART_STYLE_NAMESPACE}}}{reference_name}"),
            font=False,
        )
    _validate_chart_style_reference(
        entry.find(f"{{{_CHART_STYLE_NAMESPACE}}}fontRef"),
        font=True,
    )
    line_width_scale = entry.find(
        f"{{{_CHART_STYLE_NAMESPACE}}}lineWidthScale"
    )
    if line_width_scale is not None:
        line_width_text = _chart_leaf_text(line_width_scale)
        if (
            line_width_scale.attrib
            or line_width_text is None
            or not _is_xsd_double_lexical(line_width_text)
        ):
            raise ValueError("chart style entry has an invalid line-width scale")
    shape_properties = entry.find(
        f"{{{_CHART_STYLE_NAMESPACE}}}spPr"
    )
    if shape_properties is not None:
        _validate_chart_shape_properties(shape_properties)
    character_properties = entry.find(
        f"{{{_CHART_STYLE_NAMESPACE}}}defRPr"
    )
    if character_properties is not None:
        _validate_chart_text_character_properties(character_properties)
    body_properties = entry.find(
        f"{{{_CHART_STYLE_NAMESPACE}}}bodyPr"
    )
    if body_properties is not None:
        _validate_chart_text_body_properties(body_properties)
    extension_list = entry.find(
        f"{{{_CHART_STYLE_NAMESPACE}}}extLst"
    )
    if extension_list is not None:
        _validate_chart_extension_list(
            extension_list,
            "chart style entry extension list",
        )


def _validate_chart_marker_layout(marker_layout) -> None:
    if _chart_schema_children(marker_layout) or not set(
        marker_layout.attrib
    ).issubset({"symbol", "size"}):
        raise ValueError("chart marker layout has invalid content")
    symbol = marker_layout.get("symbol")
    if symbol is not None and symbol not in {
        "circle", "dash", "diamond", "dot", "plus", "square", "star",
        "triangle", "x",
    }:
        raise ValueError("chart marker layout has an invalid symbol")
    size = marker_layout.get("size")
    if size is not None and _parse_bounded_decimal(
        size,
        72,
        minimum=2,
    ) is None:
        raise ValueError("chart marker layout has an invalid size")


def _validate_chart_style_root(root, kind: str) -> None:
    root = _chart_mc_schema_view(root)
    allowed_attributes = {"id"}
    if kind == "color":
        allowed_attributes.add("meth")
    if not set(root.attrib).issubset(allowed_attributes):
        raise ValueError("chart style Part has unknown root attributes")
    identifier = root.get("id")
    if (
        identifier is not None
        and _parse_bounded_decimal(identifier, 0xFFFFFFFF) is None
    ):
        raise ValueError("chart style Part has an invalid ID")

    if kind == "style":
        order = {
            f"{{{_CHART_STYLE_NAMESPACE}}}{name}": index
            for index, name in enumerate(_CHART_STYLE_CHILDREN)
        }
        children = _chart_schema_children(root)
        observed = [child.tag for child in children]
        if (
            any(tag not in order for tag in observed)
            or len(observed) != len(set(observed))
            or [order[tag] for tag in observed]
            != sorted(order[tag] for tag in observed)
        ):
            raise ValueError("chart style Part has invalid children")
        required = {
            f"{{{_CHART_STYLE_NAMESPACE}}}{name}"
            for name in _CHART_STYLE_CHILDREN
            if name not in _OPTIONAL_CHART_STYLE_CHILDREN
        }
        if not required.issubset(observed):
            raise ValueError("chart style Part is missing required entries")
        marker_layout_tag = (
            f"{{{_CHART_STYLE_NAMESPACE}}}dataPointMarkerLayout"
        )
        extension_tag = f"{{{_CHART_STYLE_NAMESPACE}}}extLst"
        for child in children:
            if child.tag == marker_layout_tag:
                _validate_chart_marker_layout(child)
            elif child.tag == extension_tag:
                _validate_chart_extension_list(
                    child,
                    "chart style extension list",
                )
            else:
                _validate_chart_style_entry(child)
        return

    if "meth" not in root.attrib:
        raise ValueError("chart color style has an invalid method")
    variation_tag = f"{{{_CHART_STYLE_NAMESPACE}}}variation"
    extension_tag = f"{{{_CHART_STYLE_NAMESPACE}}}extLst"
    state = "colors"
    color_count = 0
    extension_count = 0
    for child in _chart_schema_children(root):
        if child.tag in _THEME_COLOR_MODEL_TAGS and state == "colors":
            _validate_chart_style_color_model(child, "chart color style")
            color_count += 1
            continue
        if child.tag == variation_tag and state in {"colors", "variations"}:
            if child.attrib:
                raise ValueError(
                    "chart color style variation has invalid attributes"
                )
            _validate_chart_color_transforms(
                child,
                "chart color style variation",
            )
            state = "variations"
            continue
        if child.tag == extension_tag and state in {"colors", "variations"}:
            _validate_chart_extension_list(
                child,
                "chart color style extension list",
            )
            state = "extension"
            extension_count += 1
            continue
        raise ValueError("chart color style has invalid children")
    if color_count < 1 or extension_count > 1:
        raise ValueError("chart color style has no valid color sequence")


def _chart_part_spec(relationship_type: str):
    if relationship_type == RT.CHART:
        return (
            CT.DML_CHART,
            _CHART_ROOT_TAG,
            _CHART_COLOR_MAPPING_TAG,
            (
                f"{{{_CHART_NAMESPACE}}}pivotSource",
                f"{{{_CHART_NAMESPACE}}}protection",
                f"{{{_CHART_NAMESPACE}}}chart",
                f"{{{_CHART_NAMESPACE}}}spPr",
                f"{{{_CHART_NAMESPACE}}}txPr",
                f"{{{_CHART_NAMESPACE}}}externalData",
                f"{{{_CHART_NAMESPACE}}}printSettings",
                f"{{{_CHART_NAMESPACE}}}userShapes",
                f"{{{_CHART_NAMESPACE}}}extLst",
            ),
        )
    if relationship_type == _CHARTEX_RELATIONSHIP_TYPE:
        return (
            _CHARTEX_CONTENT_TYPE,
            _CHARTEX_ROOT_TAG,
            _CHARTEX_COLOR_MAPPING_TAG,
            (
                f"{{{_CHARTEX_NAMESPACE}}}fmtOvrs",
                f"{{{_CHARTEX_NAMESPACE}}}printSettings",
                f"{{{_CHARTEX_NAMESPACE}}}extLst",
            ),
        )
    return None


def _diagram_part_spec(relationship_type: str):
    return {
        RT.DIAGRAM_DATA: (
            CT.DML_DIAGRAM_DATA,
            qn("dgm:dataModel"),
        ),
        RT.DIAGRAM_LAYOUT: (
            CT.DML_DIAGRAM_LAYOUT,
            qn("dgm:layoutDef"),
        ),
        RT.DIAGRAM_QUICK_STYLE: (
            CT.DML_DIAGRAM_STYLE,
            qn("dgm:styleDef"),
        ),
        RT.DIAGRAM_COLORS: (
            CT.DML_DIAGRAM_COLORS,
            qn("dgm:colorsDef"),
        ),
    }.get(relationship_type)


def _has_modern_comment_sidecars(document) -> bool:
    """返回主文档 Part 是否关联任一现代批注 sidecar。"""
    return any(
        relationship.reltype in _MODERN_COMMENT_RELATIONSHIP_TYPES
        for relationship in document.part.rels.values()
    )


def _reject_unsupported_inserted_modern_comments(document) -> None:
    """在完整线程合并实现前，禁止静默丢弃现代批注 sidecar。"""
    if _has_modern_comment_sidecars(document):
        raise DocumentConcatError(
            "插入文档包含现代线程批注（回复/已解决状态），"
            "当前无法无损拼接；请先在 Word 中删除这些批注或转为普通批注。"
        )


def _custom_xml_item_metadata(custom_xml_part):
    """返回 Custom XML 数据存储的 itemID、属性 Part 与安全解析根节点。"""
    try:
        properties_part = custom_xml_part.rels.part_with_reltype(
            RT.CUSTOM_XML_PROPS
        )
    except KeyError:
        return None
    except ValueError as exc:
        raise ValueError("custom XML item has multiple properties Parts") from exc

    if properties_part.content_type != CT.OFC_CUSTOM_XML_PROPERTIES:
        raise ValueError("custom XML properties Part has the wrong content type")
    properties_root = _composer_xml_part_root(properties_part)
    if properties_root.tag != _CUSTOM_XML_DATASTORE_TAG:
        raise ValueError("custom XML properties Part has the wrong root")

    item_id = properties_root.get(_CUSTOM_XML_ITEM_ID_ATTRIBUTE)
    if not isinstance(item_id, str) or not _CUSTOM_XML_ITEM_ID_RE.fullmatch(item_id):
        raise ValueError("custom XML data store has an invalid item ID")
    return item_id, properties_part, properties_root


def _allocate_custom_xml_item_id(
    source_id: str,
    source_part,
    used_ids: set[str],
) -> str:
    """为冲突的数据存储生成稳定且未占用的 GUID。"""
    source_digest = hashlib.sha256(source_part.blob).digest()
    for attempt in range(1, len(used_ids) + 2):
        digest = hashlib.sha256(
            source_id.casefold().encode("ascii")
            + b"\x00"
            + source_digest
            + b"\x00"
            + str(attempt).encode("ascii")
        ).hexdigest().upper()
        candidate = (
            f"{{{digest[:8]}-{digest[8:12]}-{digest[12:16]}-"
            f"{digest[16:20]}-{digest[20:32]}}}"
        )
        if candidate.casefold() not in used_ids:
            return candidate
    raise ValueError("custom XML data stores have no available item ID")


_BOOKMARK_NAME_LIMIT = 40
_BOOKMARK_REFERENCE_FIELD_RE = re.compile(
    r'(?P<prefix>\b(?:REF|PAGEREF|NOTEREF)\s+)'
    r'(?:(?:"(?P<quoted>(?:""|[^"])*)")|(?P<plain>[^\s\\]+))',
    flags=re.IGNORECASE,
)
_STYLEREF_FIELD_COMMAND_RE = re.compile(r"^\s*STYLEREF\b", re.IGNORECASE)
_STYLEREF_FIELD_RE = re.compile(
    r'(?P<prefix>^\s*STYLEREF\s+)'
    r'(?:(?:"(?P<quoted>(?:""|[^"])*)")|(?P<plain>[^\s\\]+))',
    re.IGNORECASE,
)
_TOC_FIELD_COMMAND_RE = re.compile(r"^\s*TOC\b", re.IGNORECASE)


def _allocate_inserted_bookmark_name(
    source_name: str,
    used_names: set[str],
) -> str:
    """为插入文档中的同名书签生成稳定名称。"""
    suffix_index = 1
    while True:
        suffix = f"__inserted_{suffix_index}"
        base = source_name[: max(1, _BOOKMARK_NAME_LIMIT - len(suffix))]
        candidate = f"{base}{suffix}"
        if candidate.casefold() not in used_names:
            return candidate
        suffix_index += 1


def _rewrite_bookmark_field_instruction(
    instruction: str | None,
    renamed_bookmarks: dict[str, str],
) -> str | None:
    if not instruction or not renamed_bookmarks:
        return instruction

    def replace(match):
        source_name = match.group("quoted")
        if source_name is not None:
            source_name = source_name.replace('""', '"')
        else:
            source_name = match.group("plain")
        target_name = renamed_bookmarks.get(source_name.casefold())
        if target_name is None or target_name == source_name:
            return match.group(0)
        if match.group("quoted") is None:
            return f'{match.group("prefix")}{target_name}'
        escaped_name = target_name.replace('"', '""')
        return f'{match.group("prefix")}"{escaped_name}"'

    return _BOOKMARK_REFERENCE_FIELD_RE.sub(replace, instruction)


def _rewrite_split_bookmark_instruction(
    instruction_nodes,
    renamed_bookmarks: dict[str, str],
) -> None:
    if not instruction_nodes:
        return
    original_fragments = [node.text or "" for node in instruction_nodes]
    original = "".join(original_fragments)
    rewritten = _rewrite_bookmark_field_instruction(
        original,
        renamed_bookmarks,
    )
    if rewritten == original:
        return
    remaining = rewritten
    for index, (node, fragment) in enumerate(
        zip(instruction_nodes, original_fragments)
    ):
        if index == len(instruction_nodes) - 1:
            node.text = remaining
            break
        fragment_length = min(len(fragment), len(remaining))
        node.text = remaining[:fragment_length]
        remaining = remaining[fragment_length:]


def _field_match_value(match) -> str:
    quoted = match.group("quoted")
    if quoted is not None:
        return quoted.replace('""', '"')
    return match.group("plain")


def _field_token(value: str, *, quoted: bool) -> str:
    if quoted or re.search(r'[\s\\"]', value):
        return f'"{value.replace(chr(34), chr(34) * 2)}"'
    return value


def _toc_style_list_parts(payload: str) -> tuple[list[str], list[str], str]:
    present_delimiters = [
        delimiter for delimiter in (",", ";") if delimiter in payload
    ]
    if len(present_delimiters) != 1:
        raise ValueError(
            "inserted TOC style list has no unique list delimiter"
        )
    delimiter = present_delimiters[0]
    parts = payload.split(delimiter)
    if not parts or len(parts) % 2:
        raise ValueError("inserted TOC style list has invalid name/level pairs")
    style_names = []
    for index in range(0, len(parts), 2):
        style_name = parts[index].strip()
        level = parts[index + 1].strip()
        if not style_name or not re.fullmatch(r"[1-9]", level):
            raise ValueError("inserted TOC style list has an invalid entry")
        style_names.append(style_name)
    return parts, style_names, delimiter


def _toc_style_switches(instruction: str) -> list[tuple[int, int, str]]:
    """Return quoted ``TOC \\t`` token spans found outside other quoted args."""
    switches = []
    index = 0
    in_quote = False
    while index < len(instruction):
        character = instruction[index]
        if character == '"':
            if in_quote and index + 1 < len(instruction) and instruction[index + 1] == '"':
                index += 2
                continue
            in_quote = not in_quote
            index += 1
            continue
        if (
            not in_quote
            and character == "\\"
            and index + 1 < len(instruction)
            and instruction[index + 1].casefold() == "t"
            and (
                index + 2 == len(instruction)
                or instruction[index + 2].isspace()
            )
        ):
            token_start = index + 2
            while token_start < len(instruction) and instruction[token_start].isspace():
                token_start += 1
            if token_start >= len(instruction) or instruction[token_start] != '"':
                raise ValueError(
                    "inserted TOC field has an unquoted style-list switch"
                )
            cursor = token_start + 1
            payload_fragments = []
            while cursor < len(instruction):
                if instruction[cursor] != '"':
                    payload_fragments.append(instruction[cursor])
                    cursor += 1
                    continue
                if cursor + 1 < len(instruction) and instruction[cursor + 1] == '"':
                    payload_fragments.append('"')
                    cursor += 2
                    continue
                token_end = cursor + 1
                switches.append(
                    (token_start, token_end, "".join(payload_fragments))
                )
                index = token_end
                break
            else:
                raise ValueError(
                    "inserted TOC field has an unterminated style list"
                )
            continue
        index += 1
    if in_quote:
        # An unrelated unterminated quoted switch makes locating ``\\t``
        # ambiguous, so fail closed rather than rewriting text inside it.
        raise ValueError("inserted TOC field has an unterminated quoted argument")
    if len(switches) > 1:
        raise ValueError("inserted TOC field has duplicate style-list switches")
    return switches


def _style_names_in_field_instruction(instruction: str | None) -> list[str]:
    if not instruction:
        return []
    if _STYLEREF_FIELD_COMMAND_RE.match(instruction):
        match = _STYLEREF_FIELD_RE.match(instruction)
        if match is None:
            raise ValueError("inserted STYLEREF field has no valid style name")
        remainder = instruction[match.end() :].lstrip()
        if remainder and not remainder.startswith("\\"):
            raise ValueError("inserted STYLEREF field has an ambiguous style name")
        return [_field_match_value(match)]
    if not _TOC_FIELD_COMMAND_RE.match(instruction):
        return []
    style_names = []
    for _token_start, _token_end, payload in _toc_style_switches(instruction):
        _parts, names, _delimiter = _toc_style_list_parts(payload)
        style_names.extend(names)
    return style_names


def _rewrite_style_name_field_instruction(
    instruction: str | None,
    renamed_styles: dict[str, str],
) -> str | None:
    if not instruction or not renamed_styles:
        return instruction
    # Validate the command before rewriting so malformed fields cannot bind to
    # an unrelated target style after composition.
    _style_names_in_field_instruction(instruction)

    if _STYLEREF_FIELD_COMMAND_RE.match(instruction):
        def replace_styleref(match):
            source_name = _field_match_value(match)
            target_name = renamed_styles.get(source_name.casefold())
            if target_name is None or target_name == source_name:
                return match.group(0)
            token = _field_token(
                target_name,
                quoted=match.group("quoted") is not None,
            )
            return f'{match.group("prefix")}{token}'

        return _STYLEREF_FIELD_RE.sub(
            replace_styleref,
            instruction,
            count=1,
        )

    if not _TOC_FIELD_COMMAND_RE.match(instruction):
        return instruction

    rewritten = instruction
    for token_start, token_end, payload in reversed(
        _toc_style_switches(instruction)
    ):
        parts, _style_names, delimiter = _toc_style_list_parts(payload)
        for index in range(0, len(parts), 2):
            original = parts[index]
            stripped = original.strip()
            target_name = renamed_styles.get(stripped.casefold())
            if target_name is None or target_name == stripped:
                continue
            if delimiter in target_name:
                raise ValueError(
                    "renamed style cannot be represented in the TOC style list"
                )
            leading = original[: len(original) - len(original.lstrip())]
            trailing = original[len(original.rstrip()) :]
            parts[index] = f"{leading}{target_name}{trailing}"
        token = _field_token(delimiter.join(parts), quoted=True)
        rewritten = f"{rewritten[:token_start]}{token}{rewritten[token_end:]}"
    return rewritten


def _complex_field_instruction_groups(root) -> tuple[list[list], set]:
    """Return instruction-node groups without mistaking field results for instructions."""
    groups = []
    owned_instruction_nodes = set()
    field_stack = []
    for node in root.iter():
        if node.tag == qn("w:fldChar"):
            field_type = node.get(qn("w:fldCharType"))
            if field_type == "begin":
                field_stack.append({"collecting": True, "nodes": []})
            elif field_type == "separate" and field_stack:
                current = field_stack[-1]
                if current["collecting"] and current["nodes"]:
                    groups.append(current["nodes"])
                current["collecting"] = False
            elif field_type == "end" and field_stack:
                current = field_stack.pop()
                if current["collecting"] and current["nodes"]:
                    groups.append(current["nodes"])
            continue
        if node.tag == qn("w:instrText") and field_stack:
            owned_instruction_nodes.add(node)
            if field_stack[-1]["collecting"]:
                field_stack[-1]["nodes"].append(node)

    for current in field_stack:
        if current["collecting"] and current["nodes"]:
            groups.append(current["nodes"])
    return groups, owned_instruction_nodes


def _style_field_instructions(root) -> list[str]:
    simple_fields = list(root.iter(qn("w:fldSimple")))
    instructions = [
        instruction
        for field in simple_fields
        if (instruction := field.get(qn("w:instr"))) is not None
    ]
    simple_result_nodes = {
        node
        for field in simple_fields
        for node in field.iter(qn("w:instrText"))
    }
    groups, owned_nodes = _complex_field_instruction_groups(root)
    instructions.extend("".join(node.text or "" for node in group) for group in groups)
    instructions.extend(
        node.text or ""
        for node in root.iter(qn("w:instrText"))
        if node not in owned_nodes and node not in simple_result_nodes
    )
    return instructions


def _replace_split_field_instruction(nodes, rewritten: str) -> None:
    fragments = [node.text or "" for node in nodes]
    remaining = rewritten
    for index, (node, fragment) in enumerate(zip(nodes, fragments)):
        if index == len(nodes) - 1:
            node.text = remaining
        else:
            fragment_length = min(len(fragment), len(remaining))
            node.text = remaining[:fragment_length]
            remaining = remaining[fragment_length:]
        if node.text and (node.text[0].isspace() or node.text[-1].isspace()):
            node.set(
                "{http://www.w3.org/XML/1998/namespace}space",
                "preserve",
            )


def _rewrite_style_fields_in_root(root, renamed_styles: dict[str, str]) -> None:
    for simple_field in root.iter(qn("w:fldSimple")):
        instruction_attribute = qn("w:instr")
        instruction = simple_field.get(instruction_attribute)
        rewritten = _rewrite_style_name_field_instruction(
            instruction,
            renamed_styles,
        )
        if rewritten != instruction:
            simple_field.set(instruction_attribute, rewritten)

    simple_result_nodes = {
        node
        for field in root.iter(qn("w:fldSimple"))
        for node in field.iter(qn("w:instrText"))
    }
    groups, owned_nodes = _complex_field_instruction_groups(root)
    for group in groups:
        instruction = "".join(node.text or "" for node in group)
        rewritten = _rewrite_style_name_field_instruction(
            instruction,
            renamed_styles,
        )
        if rewritten != instruction:
            _replace_split_field_instruction(group, rewritten)

    for instruction_node in root.iter(qn("w:instrText")):
        if instruction_node in owned_nodes or instruction_node in simple_result_nodes:
            continue
        instruction_node.text = _rewrite_style_name_field_instruction(
            instruction_node.text,
            renamed_styles,
        )


class _StructurePreservingComposerMixin:
    """Keep semantic story Parts that docxcompose 2.2 does not migrate safely."""

    def reset_reference_mapping(self):
        super().reset_reference_mapping()
        self._academic_story_id_mapping = {}
        self._academic_story_item_indexes = {}
        self._academic_story_used_ids = {}
        self._academic_custom_xml_item_id_mapping = {}
        self._academic_custom_xml_parts_copied = False
        self._academic_part_mapping = {}
        self._academic_picture_bullet_id_mapping = {}
        self._academic_style_id_mapping = {}
        self._academic_style_name_mapping = {}
        self._academic_doc_defaults_cache = {}
        self._academic_reserved_style_names = set()
        self._academic_reserved_style_ids = set()
        self._academic_style_indexes = {}
        self._academic_target_style_names = None
        self._academic_bookmark_name_mapping = dict(
            getattr(
                self,
                "_academic_pending_bookmark_name_mapping",
                {},
            )
        )
        pending_theme = getattr(
            self,
            "_academic_pending_theme_materialization",
            None,
        )
        if pending_theme is None:
            self._academic_theme_source_part_id = None
            self._academic_theme_materialization = None
        else:
            (
                self._academic_theme_source_part_id,
                self._academic_theme_materialization,
            ) = pending_theme
        self._academic_chart_theme_context = getattr(
            self,
            "_academic_pending_chart_theme_context",
            None,
        )
        package_partnames = [str(part.partname) for part in self.pkg.iter_parts()]
        folded_partnames = {name.casefold() for name in package_partnames}
        if len(folded_partnames) != len(package_partnames):
            raise ValueError("target package has case-insensitive Part name collisions")
        self._academic_reserved_partnames = folded_partnames

    def insert(self, index, doc, remove_property_fields=True):
        _reject_unsupported_inserted_modern_comments(doc)
        theme_materialization = _theme_materialization_context(
            doc,
            self.doc,
        )
        self._academic_pending_theme_materialization = (
            id(doc.part),
            theme_materialization,
        )
        self._academic_pending_chart_theme_context = _chart_theme_context(
            doc,
            self.doc,
            theme_materialization,
        )
        self._academic_pending_bookmark_name_mapping = (
            self._prepare_bookmark_name_mapping(doc)
        )
        try:
            result = super().insert(
                index,
                doc,
                remove_property_fields=remove_property_fields,
            )
            if not self._academic_custom_xml_parts_copied:
                self._copy_custom_xml_parts(doc)
            self._merge_inserted_font_table(doc)
            return result
        finally:
            self.__dict__.pop(
                "_academic_pending_bookmark_name_mapping",
                None,
            )
            self.__dict__.pop(
                "_academic_pending_theme_materialization",
                None,
            )
            self.__dict__.pop(
                "_academic_pending_chart_theme_context",
                None,
            )

    @staticmethod
    def _inserted_style_field_roots(source_doc):
        roots_and_parts = [(source_doc.element.body, None)]
        seen_parts = set()
        for relationship in source_doc.part.rels.values():
            if (
                relationship.is_external
                or relationship.reltype
                not in {
                    RT.HEADER,
                    RT.FOOTER,
                    RT.FOOTNOTES,
                    RT.ENDNOTES,
                    RT.COMMENTS,
                }
            ):
                continue
            part = relationship.target_part
            if id(part) in seen_parts:
                continue
            seen_parts.add(id(part))
            roots_and_parts.append((_composer_xml_part_root(part), part))
        return roots_and_parts

    def _prepare_inserted_style_name_fields(self, source_doc) -> None:
        roots_and_parts = self._inserted_style_field_roots(source_doc)
        referenced_names = []
        for root, _part in roots_and_parts:
            for instruction in _style_field_instructions(root):
                referenced_names.extend(
                    _style_names_in_field_instruction(instruction)
                )
        if not referenced_names:
            return

        source_styles_by_name = {}
        for style in source_doc.styles.element.findall(qn("w:style")):
            if style.get(qn("w:type")) != "paragraph":
                continue
            name_nodes = style.findall(qn("w:name"))
            if len(name_nodes) > 1:
                raise ValueError("inserted style has duplicate display names")
            if not name_nodes:
                continue
            name = name_nodes[0].get(qn("w:val"))
            if not isinstance(name, str) or not name:
                raise ValueError("inserted style has an invalid display name")
            for lookup_name in (
                name,
                *self._style_aliases(style, "inserted style"),
            ):
                matches = source_styles_by_name.setdefault(
                    lookup_name.casefold(),
                    [],
                )
                if style not in matches:
                    matches.append(style)

        target_paragraph_style_names = {
            name.casefold()
            for style in self.doc.styles.element.findall(qn("w:style"))
            if style.get(qn("w:type")) == "paragraph"
            for name in (
                *(
                    name_node.get(qn("w:val"))
                    for name_node in style.findall(qn("w:name"))
                    if name_node.get(qn("w:val"))
                ),
                *self._style_aliases(style, "target style"),
            )
        }
        renamed_styles = {}
        for source_name in dict.fromkeys(referenced_names):
            matching_styles = source_styles_by_name.get(
                source_name.casefold(),
                [],
            )
            if len(matching_styles) > 1:
                raise ValueError(
                    "inserted style field name does not resolve uniquely"
                )
            if not matching_styles:
                if source_name.casefold() in target_paragraph_style_names:
                    raise ValueError(
                        "missing inserted field style would bind to a target style"
                    )
                continue
            source_style_id = matching_styles[0].get(qn("w:styleId"))
            if not source_style_id:
                raise ValueError("inserted style field targets a style with no ID")
            target_style_id = self._ensure_inserted_style_mapping(
                source_doc,
                source_style_id,
            )
            target_style = self._style_element_by_id(
                self.doc,
                target_style_id,
            )
            if target_style is None:
                raise ValueError("inserted style field target was not copied")
            target_name_nodes = target_style.findall(qn("w:name"))
            if len(target_name_nodes) != 1:
                raise ValueError(
                    "inserted style field target has no unique display name"
                )
            target_name = target_name_nodes[0].get(qn("w:val"))
            if not isinstance(target_name, str) or not target_name:
                raise ValueError(
                    "inserted style field target has an invalid display name"
                )
            if target_name != source_name:
                renamed_styles[source_name.casefold()] = target_name

        self._academic_style_name_mapping = renamed_styles
        if not renamed_styles:
            return
        for root, part in roots_and_parts:
            _rewrite_style_fields_in_root(root, renamed_styles)
            if part is not None:
                _commit_composer_xml_part_root(part, root)

    def _create_style_id_mapping(self, doc):
        super()._create_style_id_mapping(doc)
        self._prepare_inserted_style_name_fields(doc)

    def _materialize_inserted_theme_references(self, source_doc, root) -> None:
        if id(source_doc.part) != self._academic_theme_source_part_id:
            return
        context = self._academic_theme_materialization
        if context is not None:
            context.materialize(root)

    def _prepare_bookmark_name_mapping(self, source_doc) -> dict[str, str]:
        name_attribute = qn("w:name")
        bookmark_tag = qn("w:bookmarkStart")
        used_names = set()
        for bookmark in self.doc.element.body.iter(bookmark_tag):
            name = bookmark.get(name_attribute)
            if not isinstance(name, str) or not name:
                raise ValueError("target bookmark has an invalid name")
            folded_name = name.casefold()
            if folded_name in used_names:
                raise ValueError("target bookmark names are not unique")
            used_names.add(folded_name)

        mapping = {}
        for bookmark in source_doc.element.body.iter(bookmark_tag):
            name = bookmark.get(name_attribute)
            if not isinstance(name, str) or not name:
                raise ValueError("inserted bookmark has an invalid name")
            folded_name = name.casefold()
            if folded_name in mapping:
                raise ValueError("inserted bookmark names are not unique")
            target_name = name
            if folded_name in used_names:
                target_name = _allocate_inserted_bookmark_name(
                    name,
                    used_names,
                )
            mapping[folded_name] = target_name
            used_names.add(target_name.casefold())
        return mapping

    def _rewrite_bookmark_names_and_references(self, root) -> None:
        mapping = self._academic_bookmark_name_mapping
        if not mapping:
            return

        name_attribute = qn("w:name")
        for bookmark in root.iter(qn("w:bookmarkStart")):
            source_name = bookmark.get(name_attribute)
            if not isinstance(source_name, str):
                raise ValueError("inserted bookmark has no name")
            target_name = mapping.get(source_name.casefold())
            if target_name is not None:
                bookmark.set(name_attribute, target_name)

        anchor_attribute = qn("w:anchor")
        for hyperlink in root.iter(qn("w:hyperlink")):
            source_anchor = hyperlink.get(anchor_attribute)
            if source_anchor is None:
                continue
            target_anchor = mapping.get(source_anchor.casefold())
            if target_anchor is not None:
                hyperlink.set(anchor_attribute, target_anchor)

        instruction_attribute = qn("w:instr")
        for simple_field in root.iter(qn("w:fldSimple")):
            instruction = simple_field.get(instruction_attribute)
            rewritten = _rewrite_bookmark_field_instruction(
                instruction,
                mapping,
            )
            if rewritten != instruction:
                simple_field.set(instruction_attribute, rewritten)

        processed_instruction_nodes = set()
        field_stack = []
        for node in root.iter():
            if node.tag == qn("w:fldChar"):
                field_type = node.get(qn("w:fldCharType"))
                if field_type == "begin":
                    field_stack.append([])
                elif field_type in {"separate", "end"} and field_stack:
                    instruction_nodes = field_stack[-1]
                    _rewrite_split_bookmark_instruction(
                        instruction_nodes,
                        mapping,
                    )
                    processed_instruction_nodes.update(instruction_nodes)
                    field_stack[-1] = []
                    if field_type == "end":
                        field_stack.pop()
                continue
            if node.tag == qn("w:instrText") and field_stack:
                field_stack[-1].append(node)

        for instruction_nodes in field_stack:
            _rewrite_split_bookmark_instruction(instruction_nodes, mapping)
            processed_instruction_nodes.update(instruction_nodes)

        for instruction_text in root.iter(qn("w:instrText")):
            if instruction_text in processed_instruction_nodes:
                continue
            rewritten = _rewrite_bookmark_field_instruction(
                instruction_text.text,
                mapping,
            )
            if rewritten != instruction_text.text:
                instruction_text.text = rewritten

    def _allocate_copied_partname(self, source_partname: PackURI) -> PackURI:
        reserved = self._academic_reserved_partnames
        reserved.update(
            str(part.partname).casefold() for part in self.pkg.iter_parts()
        )
        source_key = str(source_partname).casefold()
        if source_key not in reserved:
            reserved.add(source_key)
            return source_partname

        source_name = str(source_partname)
        match = re.match(r"^(.*?)(?:\d+)?(\.[^./]+)$", source_name)
        if match is None:
            base, extension = source_name, ""
        else:
            base, extension = match.groups()

        candidate_index = 1
        while True:
            candidate = PackURI(f"{base}{candidate_index}{extension}")
            candidate_key = str(candidate).casefold()
            if candidate_key not in reserved:
                reserved.add(candidate_key)
                return candidate
            candidate_index += 1

    def _copy_chart_part(self, source_part, relationship_type, spec):
        expected_content_type, root_tag, mapping_tag, mapping_successors = spec
        if str(source_part.content_type).casefold() != str(
            expected_content_type
        ).casefold():
            raise ValueError("chart relationship targets the wrong content type")
        chart_root = _parse_untrusted_ooxml_part(source_part.blob)
        if chart_root.tag != root_tag:
            raise ValueError("chart Part has the wrong root")
        _validate_chart_structure(chart_root, relationship_type)
        _relationship_references(chart_root, source_part, "chart Part")
        _validate_chart_relationship_roles(
            chart_root,
            source_part,
            relationship_type,
        )

        fallback_relationship_id = chart_root.get("fallbackImg")
        if fallback_relationship_id is not None:
            fallback_relationship = source_part.rels.get(
                fallback_relationship_id
            )
            if (
                not fallback_relationship_id
                or fallback_relationship is None
                or fallback_relationship.is_external
                or fallback_relationship.reltype != RT.IMAGE
                or not str(
                    fallback_relationship.target_part.content_type
                ).casefold().startswith("image/")
            ):
                raise ValueError(
                    "chart fallback image does not resolve to an internal image relationship"
                )

        mapping_elements = chart_root.findall(mapping_tag)
        if len(mapping_elements) > 1:
            raise ValueError("chart contains duplicate color mappings")
        if mapping_elements:
            _validate_chart_color_mapping(mapping_elements[0])

        override_relationships = [
            relationship
            for relationship in source_part.rels.values()
            if relationship.reltype == RT.THEME_OVERRIDE
        ]
        if len(override_relationships) > 1:
            raise ValueError("chart has duplicate theme override relationships")
        override_relationship = (
            override_relationships[0] if override_relationships else None
        )
        override_part = None
        override_schemes = {}
        if override_relationship is not None:
            if override_relationship.is_external:
                raise ValueError("chart has an external theme override")
            override_part = override_relationship.target_part
            _override_root, override_schemes = _validated_theme_override(
                override_part
            )

        chart_context = self._academic_chart_theme_context
        generated_override_blob = None
        if chart_context is not None and chart_context.schemes_differ:
            required_scheme_tags = (
                _THEME_COLOR_SCHEME_TAG,
                _THEME_FONT_SCHEME_TAG,
                _THEME_FORMAT_SCHEME_TAG,
            )
            missing_scheme_tags = [
                tag for tag in required_scheme_tags if tag not in override_schemes
            ]
            if missing_scheme_tags:
                if chart_context.source_schemes is None:
                    raise ValueError(
                        "chart theme cannot be preserved without a source theme"
                    )
                if (
                    chart_context.source_theme_part is not None
                    and chart_context.source_theme_part.rels
                ):
                    raise ValueError(
                        "chart theme synthesis does not support theme relationships"
                    )
                if override_part is not None and override_part.rels:
                    raise ValueError(
                        "partial chart theme override has dependent relationships"
                    )
                generated_root = etree.Element(
                    _THEME_OVERRIDE_TAG,
                    nsmap={"a": _THEME_NAMESPACE},
                )
                for scheme_tag in required_scheme_tags:
                    scheme = override_schemes.get(
                        scheme_tag,
                        chart_context.source_schemes[scheme_tag],
                    )
                    generated_root.append(deepcopy(scheme))
                generated_override_blob = etree.tostring(
                    generated_root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

        chart_changed = False
        if (
            chart_context is not None
            and chart_context.mapping_differ
            and not mapping_elements
        ):
            mapping_element = etree.Element(mapping_tag)
            for attribute_name in _CHART_COLOR_MAPPING_ATTRIBUTES:
                logical_name = (
                    "folhlink"
                    if attribute_name == "folHlink"
                    else attribute_name
                )
                mapping_element.set(
                    attribute_name,
                    chart_context.source_mapping[logical_name],
                )
            successor = next(
                (
                    child
                    for child in chart_root
                    if child.tag in mapping_successors
                ),
                None,
            )
            if successor is None:
                chart_root.append(mapping_element)
            else:
                successor.addprevious(mapping_element)
            chart_changed = True

        copied_part = Part(
            self._allocate_copied_partname(source_part.partname),
            source_part.content_type,
            (
                etree.tostring(
                    chart_root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )
                if chart_changed
                else source_part.blob
            ),
            self.pkg,
        )
        self._academic_part_mapping[id(source_part)] = copied_part

        for relationship in source_part.rels.values():
            if relationship.reltype == RT.THEME_OVERRIDE:
                continue
            if relationship.is_external:
                target = relationship.target_ref
            else:
                target = self._copy_relationship_target_part(
                    relationship.target_part,
                    relationship.reltype,
                    allow_chart_style=(
                        _chart_style_part_spec(relationship.reltype)
                        is not None
                    ),
                )
            copied_part.rels.add_relationship(
                relationship.reltype,
                target,
                relationship.rId,
                is_external=relationship.is_external,
            )

        target_override_part = None
        if generated_override_blob is not None:
            if override_part is not None:
                target_override_part = self._academic_part_mapping.get(
                    id(override_part)
                )
            if target_override_part is None:
                source_partname = (
                    override_part.partname
                    if override_part is not None
                    else PackURI("/word/theme/themeOverride1.xml")
                )
                target_override_part = Part(
                    self._allocate_copied_partname(source_partname),
                    CT.OFC_THEME_OVERRIDE,
                    generated_override_blob,
                    self.pkg,
                )
                if override_part is not None:
                    self._academic_part_mapping[
                        id(override_part)
                    ] = target_override_part
        elif override_part is not None:
            target_override_part = self._copy_relationship_target_part(
                override_part,
                RT.THEME_OVERRIDE,
            )

        if target_override_part is not None:
            relationship_id = (
                override_relationship.rId
                if override_relationship is not None
                else copied_part.rels._next_rId
            )
            copied_part.rels.add_relationship(
                RT.THEME_OVERRIDE,
                target_override_part,
                relationship_id,
            )
        return copied_part

    def _copy_chart_style_part(self, source_part, spec):
        expected_content_type, root_tag, kind = spec
        if str(source_part.content_type).casefold() != str(
            expected_content_type
        ).casefold():
            raise ValueError(
                "chart style relationship targets the wrong content type"
            )
        root = _parse_untrusted_ooxml_part(source_part.blob)
        if root.tag != root_tag:
            raise ValueError("chart style Part has the wrong root")
        _validate_chart_style_root(root, kind)
        referenced_relationships = _relationship_references(
            root,
            source_part,
            "chart style Part",
        )
        if referenced_relationships != set(source_part.rels):
            raise ValueError(
                "chart style Part has an unreferenced relationship"
            )

        copied_part = Part(
            self._allocate_copied_partname(source_part.partname),
            source_part.content_type,
            source_part.blob,
            self.pkg,
        )
        self._academic_part_mapping[id(source_part)] = copied_part
        for relationship in source_part.rels.values():
            if _chart_style_part_spec(relationship.reltype) is not None:
                raise ValueError(
                    "chart style relationship has an invalid parent"
                )
            if relationship.is_external:
                target = relationship.target_ref
            else:
                target = self._copy_relationship_target_part(
                    relationship.target_part,
                    relationship.reltype,
                )
            copied_part.rels.add_relationship(
                relationship.reltype,
                target,
                relationship.rId,
                is_external=relationship.is_external,
            )
        return copied_part

    def _copy_diagram_part(self, source_part, spec):
        expected_content_type, root_tag = spec
        if str(source_part.content_type).casefold() != str(
            expected_content_type
        ).casefold():
            raise ValueError(
                "diagram relationship targets the wrong content type"
            )
        diagram_root = _parse_untrusted_ooxml_part(source_part.blob)
        if diagram_root.tag != root_tag:
            raise ValueError("diagram Part has the wrong root")
        _relationship_references(diagram_root, source_part, "diagram Part")

        original_signature = _xml_value_signature(diagram_root)
        context = self._academic_theme_materialization
        if context is not None:
            context.materialize(diagram_root)
        diagram_changed = (
            _xml_value_signature(diagram_root) != original_signature
        )
        copied_part = Part(
            self._allocate_copied_partname(source_part.partname),
            source_part.content_type,
            (
                etree.tostring(
                    diagram_root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )
                if diagram_changed
                else source_part.blob
            ),
            self.pkg,
        )
        self._academic_part_mapping[id(source_part)] = copied_part
        for relationship in source_part.rels.values():
            if _chart_style_part_spec(relationship.reltype) is not None:
                raise ValueError(
                    "chart style relationship has an invalid parent"
                )
            if relationship.is_external:
                target = relationship.target_ref
            else:
                target = self._copy_relationship_target_part(
                    relationship.target_part,
                    relationship.reltype,
                )
            copied_part.rels.add_relationship(
                relationship.reltype,
                target,
                relationship.rId,
                is_external=relationship.is_external,
            )
        return copied_part

    def _copy_relationship_target_part(
        self,
        source_part,
        relationship_type=None,
        *,
        allow_chart_style=False,
    ):
        chart_spec = _chart_part_spec(relationship_type)
        relationship_chart_style_spec = _chart_style_part_spec(
            relationship_type
        )
        chart_style_spec = (
            relationship_chart_style_spec
            if allow_chart_style
            else None
        )
        diagram_spec = _diagram_part_spec(relationship_type)
        chart_content_types = {
            str(CT.DML_CHART).casefold(),
            _CHARTEX_CONTENT_TYPE.casefold(),
        }
        chart_style_content_types = {
            _CHART_STYLE_CONTENT_TYPE.casefold(),
            _CHART_COLOR_STYLE_CONTENT_TYPE.casefold(),
        }
        diagram_content_types = {
            str(content_type).casefold()
            for content_type in (
                CT.DML_DIAGRAM_DATA,
                CT.DML_DIAGRAM_LAYOUT,
                CT.DML_DIAGRAM_STYLE,
                CT.DML_DIAGRAM_COLORS,
            )
        }
        source_content_type = str(source_part.content_type).casefold()
        if chart_spec is None and source_content_type in chart_content_types:
            raise ValueError("chart Part uses the wrong relationship type")
        if (
            chart_spec is not None
            and source_content_type != str(chart_spec[0]).casefold()
        ):
            raise ValueError("chart relationship targets the wrong content type")
        if (
            not allow_chart_style
            and (
                relationship_chart_style_spec is not None
                or source_content_type in chart_style_content_types
            )
        ):
            raise ValueError("chart style Part has an invalid parent relationship")
        if (
            allow_chart_style
            and relationship_chart_style_spec is None
        ):
            raise ValueError("chart style Part uses the wrong relationship type")
        if (
            chart_style_spec is not None
            and source_content_type != str(chart_style_spec[0]).casefold()
        ):
            raise ValueError(
                "chart style relationship targets the wrong content type"
            )
        if (
            diagram_spec is None
            and source_content_type in diagram_content_types
        ):
            raise ValueError("diagram Part uses the wrong relationship type")
        if (
            diagram_spec is not None
            and source_content_type != str(diagram_spec[0]).casefold()
        ):
            raise ValueError(
                "diagram relationship targets the wrong content type"
            )
        if source_part.package is self.pkg:
            return source_part

        mapping_key = id(source_part)
        copied_part = self._academic_part_mapping.get(mapping_key)
        if copied_part is not None:
            return copied_part
        if chart_spec is not None:
            return self._copy_chart_part(
                source_part,
                relationship_type,
                chart_spec,
            )
        if chart_style_spec is not None:
            return self._copy_chart_style_part(
                source_part,
                chart_style_spec,
            )
        if diagram_spec is not None:
            return self._copy_diagram_part(
                source_part,
                diagram_spec,
            )

        copied_part = Part(
            self._allocate_copied_partname(source_part.partname),
            source_part.content_type,
            source_part.blob,
            self.pkg,
        )
        # 在递归前登记，既复用共享目标，也能安全处理关系图中的环。
        self._academic_part_mapping[mapping_key] = copied_part

        for relationship in source_part.rels.values():
            if _chart_style_part_spec(relationship.reltype) is not None:
                raise ValueError(
                    "chart style relationship has an invalid parent"
                )
            if relationship.is_external:
                target = relationship.target_ref
            else:
                target = self._copy_relationship_target_part(
                    relationship.target_part,
                    relationship.reltype,
                )
            # 新 Part 的关系表为空，因此可以原样保留任意合法 rId（包括
            # 稀疏或非数字 ID），无需改写其 XML blob 中的引用。
            copied_part.rels.add_relationship(
                relationship.reltype,
                target,
                relationship.rId,
                is_external=relationship.is_external,
            )
        return copied_part

    def add_relationship(self, src_part, dst_part, relationship):
        """复制关系图，同时保留嵌套 Part 的原始 relationship IDs。"""
        if relationship.is_external:
            if (
                _chart_part_spec(relationship.reltype) is not None
                or _chart_style_part_spec(relationship.reltype) is not None
            ):
                raise ValueError("chart relationship cannot be external")
            target_rid = dst_part.rels.get_or_add_ext_rel(
                relationship.reltype,
                relationship.target_ref,
            )
            return dst_part.rels[target_rid]

        target_part = self._copy_relationship_target_part(
            relationship.target_part,
            relationship.reltype,
        )
        return dst_part.rels.get_or_add(relationship.reltype, target_part)

    def numbering_part(self):
        """缺失编号 Part 时以全包唯一名创建，不绕过 OPC 分配器。"""
        try:
            numbering_part = self.doc.part.rels.part_with_reltype(
                RT.NUMBERING
            )
        except KeyError:
            partname = self._allocate_copied_partname(
                PackURI("/word/numbering.xml")
            )
            numbering_part = NumberingPart(
                partname,
                CT.WML_NUMBERING,
                OxmlElement("w:numbering"),
                self.pkg,
            )
            self.doc.part.relate_to(numbering_part, RT.NUMBERING)
        if not isinstance(numbering_part, NumberingPart):
            raise ValueError("numbering relationship targets the wrong Part type")
        return numbering_part

    @staticmethod
    def _picture_bullet_definitions_by_id(
        numbering_root,
        label: str,
    ) -> dict[int, object]:
        definitions = {}
        for definition in numbering_root.findall(qn("w:numPicBullet")):
            picture_bullet_id = _parse_bounded_decimal(
                definition.get(qn("w:numPicBulletId")),
                MAX_OOXML_DECIMAL_NUMBER,
            )
            if picture_bullet_id is None or picture_bullet_id in definitions:
                raise ValueError(f"{label} picture bullet IDs are invalid")
            definitions[picture_bullet_id] = definition
        return definitions

    def _copy_numbering_picture_bullets(
        self,
        source_doc,
        target_numbering_element,
    ) -> None:
        picture_references = list(
            target_numbering_element.iter(qn("w:lvlPicBulletId"))
        )
        if not picture_references:
            return

        source_part = source_doc.part.numbering_part
        target_part = self.numbering_part()
        source_definitions = self._picture_bullet_definitions_by_id(
            source_part.element,
            "inserted numbering",
        )
        target_definitions = self._picture_bullet_definitions_by_id(
            target_part.element,
            "target numbering",
        )
        used_target_ids = set(target_definitions)
        mapping = self._academic_picture_bullet_id_mapping

        for reference in picture_references:
            source_id = _parse_bounded_decimal(
                reference.get(qn("w:val")),
                MAX_OOXML_DECIMAL_NUMBER,
            )
            if source_id is None:
                raise ValueError("inserted picture bullet reference is invalid")

            target_id = mapping.get(source_id)
            if target_id is None:
                source_definition = source_definitions.get(source_id)
                if source_definition is None:
                    raise ValueError(
                        "inserted picture bullet reference does not resolve"
                    )

                target_id = source_id
                if target_id in used_target_ids:
                    target_id = 0
                    while target_id in used_target_ids:
                        target_id += 1
                    if target_id > MAX_OOXML_DECIMAL_NUMBER:  # pragma: no cover
                        raise ValueError("picture bullet IDs are exhausted")

                copied_definition = deepcopy(source_definition)
                copied_definition.set(
                    qn("w:numPicBulletId"),
                    str(target_id),
                )
                self._materialize_inserted_theme_references(
                    source_doc,
                    copied_definition,
                )
                _copy_composer_element_relationships(
                    self,
                    source_part,
                    target_part,
                    copied_definition,
                )

                insert_index = len(target_part.element)
                for index, child in enumerate(target_part.element):
                    if child.tag in {
                        qn("w:abstractNum"),
                        qn("w:num"),
                        qn("w:numIdMacAtCleanup"),
                    }:
                        insert_index = index
                        break
                target_part.element.insert(insert_index, copied_definition)
                mapping[source_id] = target_id
                used_target_ids.add(target_id)

            reference.set(qn("w:val"), str(target_id))

    def add_numberings(self, doc, element):
        """逐一定义迁移编号，避免多 numId 复用同一目标 ID。"""
        numbering_references = list(element.iter(qn("w:numId")))
        referenced_num_ids = set()
        for num_id in numbering_references:
            normalized_id = _parse_bounded_decimal(
                num_id.get(qn("w:val")),
                MAX_OOXML_DECIMAL_NUMBER,
            )
            if normalized_id is None:
                raise ValueError("inserted numbering reference is invalid")
            referenced_num_ids.add(normalized_id)

        nonzero_num_ids = referenced_num_ids - {0}
        source_definitions = {}
        if nonzero_num_ids:
            try:
                source_part = doc.part.rels.part_with_reltype(RT.NUMBERING)
            except (KeyError, ValueError) as exc:
                raise ValueError(
                    "inserted numbering reference has no numbering Part"
                ) from exc
            if not isinstance(source_part, NumberingPart):
                raise ValueError(
                    "inserted numbering relationship targets the wrong Part type"
                )
            source_root = source_part.element
            for referenced_id in nonzero_num_ids:
                matching_nums = [
                    num
                    for num in source_root.findall(qn("w:num"))
                    if _parse_bounded_decimal(
                        num.get(qn("w:numId")),
                        MAX_OOXML_DECIMAL_NUMBER,
                    ) == referenced_id
                ]
                if len(matching_nums) != 1:
                    raise ValueError(
                        "inserted numbering reference does not resolve uniquely"
                    )
                abstract_references = matching_nums[0].findall(
                    qn("w:abstractNumId")
                )
                if len(abstract_references) != 1:
                    raise ValueError(
                        "inserted numbering definition has an invalid abstract reference"
                    )
                abstract_id = _parse_bounded_decimal(
                    abstract_references[0].get(qn("w:val")),
                    MAX_OOXML_DECIMAL_NUMBER,
                )
                matching_abstract_nums = [
                    abstract_num
                    for abstract_num in source_root.findall(qn("w:abstractNum"))
                    if _parse_bounded_decimal(
                        abstract_num.get(qn("w:abstractNumId")),
                        MAX_OOXML_DECIMAL_NUMBER,
                    ) == abstract_id
                ]
                if abstract_id is None or len(matching_abstract_nums) != 1:
                    raise ValueError(
                        "inserted numbering abstract definition does not resolve uniquely"
                    )
                source_definitions[referenced_id] = (
                    matching_nums[0],
                    abstract_id,
                    matching_abstract_nums[0],
                )

        for source_num_id in sorted(nonzero_num_ids):
            if source_num_id in self.num_id_mapping:
                continue
            source_num, source_abstract_id, source_abstract_num = (
                source_definitions[source_num_id]
            )
            target_part = self.numbering_part()
            next_num_id, next_abstract_id = self._next_numbering_ids()
            if (
                next_num_id > MAX_OOXML_DECIMAL_NUMBER
                or next_abstract_id > MAX_OOXML_DECIMAL_NUMBER
            ):
                raise ValueError("target numbering IDs are exhausted")

            copied_num = deepcopy(source_num)
            copied_num.set(qn("w:numId"), str(next_num_id))
            self.num_id_mapping[source_num_id] = next_num_id

            copied_abstract_num = None
            target_abstract_id = self.anum_id_mapping.get(source_abstract_id)
            if target_abstract_id is None:
                target_abstract_id = next_abstract_id
                self.anum_id_mapping[source_abstract_id] = target_abstract_id
                copied_abstract_num = deepcopy(source_abstract_num)
                copied_abstract_num.set(
                    qn("w:abstractNumId"),
                    str(target_abstract_id),
                )
                nsid = copied_abstract_num.find(qn("w:nsid"))
                if nsid is not None:
                    used_nsids = {
                        value.casefold()
                        for abstract_num in target_part.element.findall(
                            qn("w:abstractNum")
                        )
                        for nsid_element in abstract_num.findall(qn("w:nsid"))
                        if (value := nsid_element.get(qn("w:val")))
                    }
                    nonce = 0
                    while True:
                        seed = (
                            etree.tostring(source_abstract_num)
                            + f":{target_abstract_id}:{nonce}".encode("ascii")
                        )
                        candidate_nsid = hashlib.sha256(seed).hexdigest()[:8]
                        if candidate_nsid.casefold() not in used_nsids:
                            nsid.set(qn("w:val"), candidate_nsid.upper())
                            break
                        nonce += 1

                self._materialize_inserted_theme_references(
                    doc,
                    copied_abstract_num,
                )
                insert_index = len(target_part.element)
                for index, child in enumerate(target_part.element):
                    if child.tag in {qn("w:num"), qn("w:numIdMacAtCleanup")}:
                        insert_index = index
                        break
                target_part.element.insert(insert_index, copied_abstract_num)

            abstract_reference = copied_num.find(qn("w:abstractNumId"))
            abstract_reference.set(qn("w:val"), str(target_abstract_id))
            self._materialize_inserted_theme_references(doc, copied_num)
            target_part.element._insert_num(copied_num)

            # Register and insert both IDs before walking numbering->style
            # references, because a referenced style can point back to this
            # same numbering definition.
            if copied_abstract_num is not None:
                self.add_styles(doc, copied_abstract_num)
                self._copy_numbering_picture_bullets(
                    doc,
                    copied_abstract_num,
                )
            self.add_styles(doc, copied_num)
            self._copy_numbering_picture_bullets(doc, copied_num)

        for numbering_reference in numbering_references:
            source_num_id = _parse_bounded_decimal(
                numbering_reference.get(qn("w:val")),
                MAX_OOXML_DECIMAL_NUMBER,
            )
            mapped_num_id = self.num_id_mapping.get(source_num_id)
            if mapped_num_id is not None:
                numbering_reference.set(qn("w:val"), str(mapped_num_id))

    @staticmethod
    def _font_definitions_by_name(fonts_root, label: str) -> dict[str, object]:
        definitions = {}
        for font in fonts_root.findall(qn("w:font")):
            name = font.get(qn("w:name"))
            if not isinstance(name, str) or not name:
                raise ValueError(f"{label} contains an invalid font name")
            folded_name = name.casefold()
            if folded_name in definitions:
                raise ValueError(f"{label} font names are not unique")
            definitions[folded_name] = font
        return definitions

    def _rewrite_embedded_font_relationships(
        self,
        source_part,
        target_part,
        element,
    ) -> None:
        for tag_name in _EMBEDDED_FONT_TAG_NAMES:
            for embedded_font in element.iter(qn(tag_name)):
                source_rid = embedded_font.get(qn("r:id"))
                relationship = source_part.rels.get(source_rid)
                if (
                    not source_rid
                    or relationship is None
                    or relationship.is_external
                    or relationship.reltype != RT.FONT
                ):
                    raise ValueError(
                        "inserted embedded font has an invalid relationship"
                    )
                copied_relationship = self.add_relationship(
                    source_part,
                    target_part,
                    relationship,
                )
                embedded_font.set(qn("r:id"), copied_relationship.rId)

    def _merge_inserted_font_table(self, source_doc) -> None:
        source_relationships = [
            relationship
            for relationship in source_doc.part.rels.values()
            if relationship.reltype == RT.FONT_TABLE
        ]
        if not source_relationships:
            return
        if len(source_relationships) != 1 or source_relationships[0].is_external:
            raise ValueError("inserted document has an invalid font table")

        source_relationship = source_relationships[0]
        source_part = source_relationship.target_part
        if (
            not isinstance(source_part.content_type, str)
            or source_part.content_type.lower() != CT.WML_FONT_TABLE.lower()
        ):
            raise ValueError("inserted font table has the wrong content type")
        source_root = _composer_xml_part_root(source_part)
        if source_root.tag != qn("w:fonts"):
            raise ValueError("inserted font table has the wrong root")

        target_relationships = [
            relationship
            for relationship in self.doc.part.rels.values()
            if relationship.reltype == RT.FONT_TABLE
        ]
        if not target_relationships:
            self.add_relationship(
                source_doc.part,
                self.doc.part,
                source_relationship,
            )
            return
        if len(target_relationships) != 1 or target_relationships[0].is_external:
            raise ValueError("target document has an invalid font table")

        target_part = target_relationships[0].target_part
        if (
            not isinstance(target_part.content_type, str)
            or target_part.content_type.lower() != CT.WML_FONT_TABLE.lower()
        ):
            raise ValueError("target font table has the wrong content type")
        target_root = _composer_xml_part_root(target_part)
        if target_root.tag != qn("w:fonts"):
            raise ValueError("target font table has the wrong root")

        source_fonts = self._font_definitions_by_name(
            source_root,
            "inserted font table",
        )
        target_fonts = self._font_definitions_by_name(
            target_root,
            "target font table",
        )
        changed = False
        for folded_name, source_font in source_fonts.items():
            target_font = target_fonts.get(folded_name)
            if target_font is None:
                copied_font = deepcopy(source_font)
                self._rewrite_embedded_font_relationships(
                    source_part,
                    target_part,
                    copied_font,
                )
                target_root.append(copied_font)
                target_fonts[folded_name] = copied_font
                changed = True
                continue

            for index, tag_name in enumerate(_EMBEDDED_FONT_TAG_NAMES):
                source_variants = source_font.findall(qn(tag_name))
                target_variants = target_font.findall(qn(tag_name))
                if len(source_variants) > 1 or len(target_variants) > 1:
                    raise ValueError("font table contains duplicate embedded variants")
                if not source_variants or target_variants:
                    continue

                copied_variant = deepcopy(source_variants[0])
                self._rewrite_embedded_font_relationships(
                    source_part,
                    target_part,
                    copied_variant,
                )
                _insert_xml_child_before(
                    target_font,
                    copied_variant,
                    _EMBEDDED_FONT_TAG_NAMES[index + 1 :],
                )
                changed = True

        if changed:
            _commit_composer_xml_part_root(target_part, target_root)

    def _get_or_add_body_image_part(self, source_image_part):
        """以全包唯一的 Part 名复制正文图片。"""
        mapping_key = id(source_image_part)
        mapped_part = self._academic_part_mapping.get(mapping_key)
        if mapped_part is not None:
            return mapped_part

        source_sha1 = getattr(source_image_part, "sha1", None)
        if not isinstance(source_sha1, str):
            source_sha1 = hashlib.sha1(source_image_part.blob).hexdigest()
        existing = self.pkg.image_parts._get_by_sha1(source_sha1)
        if existing is not None:
            self._academic_part_mapping[mapping_key] = existing
            return existing

        from docx.parts.image import ImagePart
        from docx.image.image import Image
        from docxcompose.image import ImageWrapper

        if hasattr(source_image_part, "sha1") and hasattr(
            source_image_part,
            "filename",
        ):
            image = ImageWrapper(source_image_part)
        else:
            image = Image.from_blob(source_image_part.blob)
        partname = self._allocate_copied_partname(source_image_part.partname)
        copied = ImagePart.from_image(image, partname)
        copied._package = self.pkg
        self.pkg.image_parts.append(copied)
        self._academic_part_mapping[mapping_key] = copied
        return copied

    def add_images(self, doc, element):
        """复制 DrawingML 图像、3D 模型与媒体关系。"""

        for node in element.iter():
            _validate_drawingml_relationship_attribute_shape(
                node,
                "inserted DrawingML",
            )

        for blip in element.iter():
            if (
                blip.tag not in _DRAWINGML_IMAGE_TAGS
                and blip.tag not in _DRAWINGML_EMBEDDED_IMAGE_ONLY_TAGS
            ):
                continue

            source_rid = blip.get(_RELATIONSHIP_EMBED_ATTRIBUTE)
            if source_rid:
                relationship = doc.part.rels.get(source_rid)
                if (
                    relationship is None
                    or relationship.is_external
                    or relationship.reltype != RT.IMAGE
                ):
                    raise ValueError("inserted image has a missing relationship")
                copied_part = self._get_or_add_body_image_part(
                    relationship.target_part
                )
                target_rid = self.doc.part.relate_to(copied_part, RT.IMAGE)
                blip.set(_RELATIONSHIP_EMBED_ATTRIBUTE, target_rid)

            if blip.tag in _DRAWINGML_EMBEDDED_IMAGE_ONLY_TAGS:
                continue

            linked_rid = blip.get(_RELATIONSHIP_LINK_ATTRIBUTE)
            if linked_rid:
                linked_relationship = doc.part.rels.get(linked_rid)
                if (
                    linked_relationship is None
                    or linked_relationship.reltype != RT.IMAGE
                ):
                    raise ValueError("inserted linked image has a missing relationship")
                copied_relationship = self.add_relationship(
                    doc.part,
                    self.doc.part,
                    linked_relationship,
                )
                blip.set(_RELATIONSHIP_LINK_ATTRIBUTE, copied_relationship.rId)

        model_tag = f"{{{_DRAWING_2017_MODEL3D_NAMESPACE}}}model3d"
        for model in element.iter(model_tag):
            for attribute_name, embedded in (
                (_RELATIONSHIP_EMBED_ATTRIBUTE, True),
                (_RELATIONSHIP_LINK_ATTRIBUTE, False),
            ):
                source_rid = model.get(attribute_name)
                if not source_rid:
                    continue
                relationship = doc.part.rels.get(source_rid)
                if (
                    relationship is None
                    or relationship.reltype != _MODEL3D_RELATIONSHIP_TYPE
                    or (embedded and relationship.is_external)
                    or (
                        not relationship.is_external
                        and str(
                            relationship.target_part.content_type
                        ).casefold()
                        != _MODEL3D_CONTENT_TYPE.casefold()
                    )
                ):
                    raise ValueError(
                        "inserted 3D model has an invalid relationship"
                    )
                copied_relationship = self.add_relationship(
                    doc.part,
                    self.doc.part,
                    relationship,
                )
                model.set(attribute_name, copied_relationship.rId)

        for tag, attribute_name, relationship_type, must_be_internal in (
            _DRAWINGML_MEDIA_RELATIONSHIP_SPECS
        ):
            for media in element.iter(tag):
                source_rid = media.get(attribute_name)
                if not source_rid:
                    continue
                relationship = doc.part.rels.get(source_rid)
                if (
                    relationship is None
                    or relationship.reltype != relationship_type
                    or (must_be_internal and relationship.is_external)
                ):
                    raise ValueError(
                        "inserted DrawingML media has an invalid relationship"
                    )
                copied_relationship = self.add_relationship(
                    doc.part,
                    self.doc.part,
                    relationship,
                )
                media.set(attribute_name, copied_relationship.rId)

    def add_diagrams(self, doc, element):
        """通过全包分配器复制 SmartArt/Diagram 的四类 Part。"""
        from docxcompose.utils import NS, xpath

        relationship_specs = (
            ("dm", RT.DIAGRAM_DATA),
            ("lo", RT.DIAGRAM_LAYOUT),
            ("qs", RT.DIAGRAM_QUICK_STYLE),
            ("cs", RT.DIAGRAM_COLORS),
        )
        for relation_ids in xpath(element, ".//dgm:relIds"):
            for attribute, expected_type in relationship_specs:
                attribute_name = f"{{{NS['r']}}}{attribute}"
                source_rid = relation_ids.get(attribute_name)
                relationship = doc.part.rels.get(source_rid)
                if (
                    not source_rid
                    or relationship is None
                    or relationship.is_external
                    or relationship.reltype != expected_type
                ):
                    raise ValueError("inserted diagram has an invalid relationship")
                copied_relationship = self.add_relationship(
                    doc.part,
                    self.doc.part,
                    relationship,
                )
                relation_ids.set(attribute_name, copied_relationship.rId)

    def add_shapes(self, doc, element):
        """复制 VML 图片，与通用关系图共用 Part 名分配器。"""
        from docxcompose.utils import NS, xpath

        for shape_resource in xpath(
            element,
            ".//v:imagedata | .//v:fill | .//v:stroke",
        ):
            relationship_attributes = [f"{{{NS['r']}}}id"]
            hyperlink_attribute = f"{{{NS['r']}}}href"
            if shape_resource.tag == qn("v:imagedata"):
                relationship_attributes.extend(
                    (
                        qn("o:relid"),
                        f"{{{NS['r']}}}pict",
                        hyperlink_attribute,
                    )
                )
            for attribute_name in relationship_attributes:
                source_rid = shape_resource.get(attribute_name)
                if not source_rid:
                    continue
                relationship = doc.part.rels.get(source_rid)
                if attribute_name == hyperlink_attribute:
                    if relationship is None:
                        raise ValueError(
                            "inserted shape has a missing relationship"
                        )
                    copied_relationship = self.add_relationship(
                        doc.part,
                        self.doc.part,
                        relationship,
                    )
                    shape_resource.set(
                        attribute_name,
                        copied_relationship.rId,
                    )
                    continue
                if (
                    relationship is None
                    or relationship.reltype != RT.IMAGE
                ):
                    raise ValueError(
                        "inserted shape has a missing image relationship"
                    )
                if relationship.is_external:
                    copied_relationship = self.add_relationship(
                        doc.part,
                        self.doc.part,
                        relationship,
                    )
                    target_rid = copied_relationship.rId
                else:
                    copied_part = self._get_or_add_body_image_part(
                        relationship.target_part
                    )
                    target_rid = self.doc.part.relate_to(
                        copied_part,
                        RT.IMAGE,
                    )
                shape_resource.set(attribute_name, target_rid)

    def _copy_custom_xml_parts(self, source_doc):
        if self._academic_custom_xml_parts_copied:
            return

        target_used_ids = set()
        for relationship in self.doc.part.rels.values():
            if relationship.reltype != RT.CUSTOM_XML:
                continue
            if relationship.is_external:
                raise ValueError("target custom XML relationship is external")
            metadata = _custom_xml_item_metadata(relationship.target_part)
            if metadata is None:
                continue
            target_id = metadata[0]
            folded_target_id = target_id.casefold()
            if folded_target_id in target_used_ids:
                raise ValueError("target custom XML item IDs are not unique")
            target_used_ids.add(folded_target_id)

        source_seen_ids = set()
        for relationship in source_doc.part.rels.values():
            if relationship.reltype != RT.CUSTOM_XML:
                continue
            if relationship.is_external:
                raise ValueError("inserted custom XML relationship is external")

            source_part = relationship.target_part
            metadata = _custom_xml_item_metadata(source_part)
            source_id = metadata[0] if metadata is not None else None
            target_id = source_id
            if source_id is not None:
                folded_source_id = source_id.casefold()
                if folded_source_id in source_seen_ids:
                    raise ValueError("inserted custom XML item IDs are not unique")
                source_seen_ids.add(folded_source_id)
                if folded_source_id in target_used_ids:
                    target_id = _allocate_custom_xml_item_id(
                        source_id,
                        source_part,
                        target_used_ids,
                    )

            copied_relationship = self.add_relationship(
                source_doc.part,
                self.doc.part,
                relationship,
            )
            if source_id is None:
                continue

            if target_id != source_id:
                copied_metadata = _custom_xml_item_metadata(
                    copied_relationship.target_part
                )
                if copied_metadata is None:  # pragma: no cover - graph copy invariant
                    raise ValueError("copied custom XML item lost its properties Part")
                _copied_id, copied_properties_part, copied_properties_root = (
                    copied_metadata
                )
                copied_properties_root.set(
                    _CUSTOM_XML_ITEM_ID_ATTRIBUTE,
                    target_id,
                )
                _commit_composer_xml_part_root(
                    copied_properties_part,
                    copied_properties_root,
                )

            self._academic_custom_xml_item_id_mapping[
                source_id.casefold()
            ] = target_id
            target_used_ids.add(target_id.casefold())

        self._academic_custom_xml_parts_copied = True

    def _rewrite_custom_xml_bindings(self, element):
        item_id_mapping = self._academic_custom_xml_item_id_mapping
        store_item_id_attribute = qn("w:storeItemID")
        for binding_tag in _CUSTOM_XML_DATA_BINDING_TAGS:
            for data_binding in element.iter(binding_tag):
                source_id = data_binding.get(store_item_id_attribute)
                if not isinstance(source_id, str):
                    raise ValueError("inserted data binding has no data store ID")
                target_id = item_id_mapping.get(source_id.casefold())
                if target_id is None:
                    raise ValueError("inserted data binding does not resolve")
                data_binding.set(store_item_id_attribute, target_id)

    def _copy_body_relationship_attribute(
        self,
        source_part,
        target_part,
        node,
        attribute_name,
        *,
        ignored_relationship_types=frozenset(),
    ):
        """复制 docxcompose 默认扫描未覆盖的正文关系属性。"""
        source_rid = node.get(attribute_name)
        if source_rid is None:
            return
        relationship = source_part.rels.get(source_rid)
        if relationship is None:
            raise ValueError("inserted body element has a missing relationship")
        if relationship.reltype in ignored_relationship_types:
            return
        copied_relationship = self.add_relationship(
            source_part,
            target_part,
            relationship,
        )
        node.set(attribute_name, copied_relationship.rId)

    def add_styles_from_other_parts(self, doc):
        """以安全解析后的 footnotes Part 迁移样式，避免 docxcompose 直接 parse。"""
        try:
            source_part = doc.part.rels.part_with_reltype(RT.FOOTNOTES)
        except (KeyError, ValueError):
            return
        source_root = _composer_xml_part_root(source_part)
        if source_root.tag != qn("w:footnotes"):
            raise ValueError("inserted footnotes Part has the wrong root")
        self.add_styles(doc, source_root)

    def add_styles(self, doc, element):
        """按完整依赖图复制冲突样式，并逐节点改写样式引用。"""
        from docxcompose.utils import xpath

        for reference in xpath(
            element,
            ".//w:tblStyle|.//w:pStyle|.//w:rStyle|"
            ".//w:styleLink|.//w:numStyleLink",
        ):
            source_style_id = reference.val
            reference.val = self._ensure_inserted_style_mapping(
                doc,
                source_style_id,
            )

    @staticmethod
    def _style_dependency_references(style_element):
        dependencies = []
        for tag_name in ("w:basedOn", "w:next", "w:link"):
            nodes = style_element.findall(qn(tag_name))
            if len(nodes) > 1:
                raise ValueError("style contains duplicate dependency references")
            if not nodes:
                continue
            style_id = nodes[0].get(qn("w:val"))
            if not style_id:
                raise ValueError("style contains an invalid dependency reference")
            dependencies.append((tag_name, nodes[0], style_id))

        internal_index = 0
        internal_tags = {
            qn("w:tblStyle"),
            qn("w:pStyle"),
            qn("w:rStyle"),
            qn("w:styleLink"),
            qn("w:numStyleLink"),
        }
        for node in style_element.iter():
            if node.tag not in internal_tags:
                continue
            style_id = node.get(qn("w:val"))
            if not style_id:
                raise ValueError(
                    "style contains an invalid internal style reference"
                )
            dependencies.append(
                (("internal", node.tag, internal_index), node, style_id)
            )
            internal_index += 1
        return tuple(dependencies)

    @classmethod
    def _style_dependency_ids(cls, style_element) -> tuple[tuple[object, str], ...]:
        return tuple(
            (edge_key, style_id)
            for edge_key, _node, style_id in cls._style_dependency_references(
                style_element
            )
        )

    def _build_style_index(self, document) -> dict[str | None, list]:
        """Index a styles Part once while retaining duplicate-ID validation."""
        index = {}
        for style in document.styles.element.findall(qn("w:style")):
            index.setdefault(style.get(qn("w:styleId")), []).append(style)
        return index

    def _style_index(self, document) -> dict[str | None, list]:
        styles_root = document.styles.element
        cache_key = id(styles_root)
        cached = self._academic_style_indexes.get(cache_key)
        if cached is None or cached[0] is not styles_root:
            cached = (styles_root, self._build_style_index(document))
            self._academic_style_indexes[cache_key] = cached
        return cached[1]

    def _register_target_style_element(self, style_element) -> None:
        """Keep the target index current as copied styles are appended."""
        styles_root = self.doc.styles.element
        cached = self._academic_style_indexes.get(id(styles_root))
        if cached is None or cached[0] is not styles_root:
            return
        style_id = style_element.get(qn("w:styleId"))
        cached[1].setdefault(style_id, []).append(style_element)

    def _style_element_by_id(self, document, style_id: str):
        if not isinstance(style_id, str) or not style_id:
            raise ValueError("style reference has an invalid style ID")
        matches = self._style_index(document).get(style_id, ())
        if len(matches) > 1:
            raise ValueError("style ID does not resolve uniquely")
        return matches[0] if matches else None

    def _build_target_style_names(self) -> set[str]:
        return {
            name.casefold()
            for style in self.doc.styles.element.findall(qn("w:style"))
            for name in (
                *(
                    name_node.get(qn("w:val"))
                    for name_node in style.findall(qn("w:name"))
                    if name_node.get(qn("w:val"))
                ),
                *self._style_aliases(style, "target style"),
            )
        }

    def _target_style_names(self) -> set[str]:
        if self._academic_target_style_names is None:
            self._academic_target_style_names = self._build_target_style_names()
        return self._academic_target_style_names

    def _remember_inserted_style_mapping(
        self,
        source_style_id: str,
        target_style_id: str,
    ) -> None:
        self._academic_style_id_mapping[source_style_id] = target_style_id
        self._academic_reserved_style_ids.add(target_style_id)

    _ACADEMIC_PARAGRAPH_PROPERTY_ORDER = (
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
    _ACADEMIC_RUN_PROPERTY_ORDER = (
        "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
        "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow",
        "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
        "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern",
        "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
        "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs",
        "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath",
        "w:rPrChange",
    )
    _ACADEMIC_RUN_DEFAULT_FALSE_TAGS = frozenset(
        qn(tag_name)
        for tag_name in (
            "w:b",
            "w:bCs",
            "w:i",
            "w:iCs",
            "w:caps",
            "w:smallCaps",
            "w:strike",
            "w:dstrike",
            "w:outline",
            "w:shadow",
            "w:emboss",
            "w:imprint",
            "w:noProof",
            "w:vanish",
            "w:webHidden",
            "w:rtl",
            "w:cs",
            "w:specVanish",
            "w:oMath",
        )
    )
    _ACADEMIC_PARAGRAPH_DEFAULT_FALSE_TAGS = frozenset(
        qn(tag_name)
        for tag_name in (
            "w:keepNext",
            "w:keepLines",
            "w:pageBreakBefore",
            "w:suppressLineNumbers",
            "w:suppressAutoHyphens",
            "w:bidi",
            "w:contextualSpacing",
            "w:mirrorIndents",
            "w:suppressOverlap",
        )
    )

    @staticmethod
    def _style_property_attribute_groups(property_tag) -> tuple[tuple, ...]:
        if property_tag == qn("w:rFonts"):
            return (
                (qn("w:ascii"), qn("w:asciiTheme")),
                (qn("w:hAnsi"), qn("w:hAnsiTheme")),
                (qn("w:eastAsia"), qn("w:eastAsiaTheme")),
                (qn("w:cs"), qn("w:cstheme")),
            )
        if property_tag == qn("w:ind"):
            return (
                (qn("w:left"), qn("w:start")),
                (qn("w:right"), qn("w:end")),
                (qn("w:leftChars"), qn("w:startChars")),
                (qn("w:rightChars"), qn("w:endChars")),
                (
                    qn("w:firstLine"),
                    qn("w:hanging"),
                    qn("w:firstLineChars"),
                    qn("w:hangingChars"),
                ),
            )
        if property_tag == qn("w:spacing"):
            return (
                (
                    qn("w:before"),
                    qn("w:beforeLines"),
                    qn("w:beforeAutospacing"),
                ),
                (
                    qn("w:after"),
                    qn("w:afterLines"),
                    qn("w:afterAutospacing"),
                ),
            )
        return ()

    @staticmethod
    def _merge_missing_style_property_attributes(
        target_property,
        source_property,
    ) -> None:
        if target_property.tag not in {
            qn("w:rFonts"),
            qn("w:lang"),
            qn("w:ind"),
            qn("w:spacing"),
        }:
            return
        grouped_attributes = (
            _StructurePreservingComposerMixin._style_property_attribute_groups(
                target_property.tag
            )
        )
        grouped_by_attribute = {
            attribute: group
            for group in grouped_attributes
            for attribute in group
        }
        original_target_attributes = set(target_property.attrib)
        for attribute, value in source_property.attrib.items():
            if attribute in target_property.attrib:
                continue
            group = grouped_by_attribute.get(attribute)
            if group is not None and any(
                grouped_attribute in original_target_attributes
                for grouped_attribute in group
            ):
                continue
            target_property.set(attribute, value)

    @staticmethod
    def _insert_missing_style_properties(
        target_properties,
        source_properties,
        property_order,
    ) -> None:
        source_tags = [property_element.tag for property_element in source_properties]
        if len(source_tags) != len(set(source_tags)):
            raise ValueError("default style contains duplicate formatting properties")
        existing_tags = {child.tag for child in target_properties}
        ordered_tags = tuple(qn(tag_name) for tag_name in property_order)
        order_indexes = {
            tag_name: index for index, tag_name in enumerate(ordered_tags)
        }
        existing_by_tag = {}
        for child in target_properties:
            existing_by_tag.setdefault(child.tag, []).append(child)
        for source_property in source_properties:
            if source_property.tag in existing_tags:
                matches = existing_by_tag[source_property.tag]
                if len(matches) == 1:
                    _StructurePreservingComposerMixin._merge_missing_style_property_attributes(
                        matches[0],
                        source_property,
                    )
                continue
            copied_property = deepcopy(source_property)
            property_index = order_indexes.get(source_property.tag)
            if property_index is None:
                target_properties.append(copied_property)
            else:
                successor = next(
                    (
                        child
                        for child in target_properties
                        if order_indexes.get(child.tag, -1) > property_index
                    ),
                    None,
                )
                if successor is None:
                    target_properties.append(copied_property)
                else:
                    successor.addprevious(copied_property)
            existing_tags.add(source_property.tag)
            existing_by_tag[source_property.tag] = [copied_property]

    def _style_chain_run_property_tags(self, document, style_id: str) -> set[str]:
        property_tags = set()
        visited = set()
        current_id = style_id
        while current_id not in visited:
            visited.add(current_id)
            style = self._style_element_by_id(document, current_id)
            if style is None:
                break
            run_properties = style.find(qn("w:rPr"))
            if run_properties is not None:
                property_tags.update(child.tag for child in run_properties)
            based_on = [
                dependency_id
                for tag_name, dependency_id in self._style_dependency_ids(style)
                if tag_name == "w:basedOn"
            ]
            if not based_on:
                break
            current_id = based_on[0]
        return property_tags

    @staticmethod
    def _document_default_property_elements(document) -> tuple[list, list]:
        styles_root = document.styles.element
        defaults = styles_root.findall(qn("w:docDefaults"))
        if len(defaults) > 1:
            raise ValueError("styles Part contains duplicate document defaults")
        if not defaults:
            return [], []

        def properties(default_tag: str, property_tag: str) -> list:
            default_nodes = defaults[0].findall(qn(default_tag))
            if len(default_nodes) > 1:
                raise ValueError(
                    "styles Part contains duplicate document-default containers"
                )
            if not default_nodes:
                return []
            property_nodes = default_nodes[0].findall(qn(property_tag))
            if len(property_nodes) > 1:
                raise ValueError(
                    "styles Part contains duplicate document-default properties"
                )
            if not property_nodes:
                return []
            children = list(property_nodes[0])
            tags = [child.tag for child in children]
            if len(tags) != len(set(tags)):
                raise ValueError(
                    "document defaults contain duplicate formatting properties"
                )
            return children

        return (
            properties("w:pPrDefault", "w:pPr"),
            properties("w:rPrDefault", "w:rPr"),
        )

    @classmethod
    def _different_document_default_properties(
        cls,
        source_properties,
        target_properties,
    ) -> tuple[list, list, frozenset[str]]:
        source_by_tag = {
            property_element.tag: property_element
            for property_element in source_properties
        }
        target_by_tag = {
            property_element.tag: property_element
            for property_element in target_properties
        }
        different_tags = {
            tag
            for tag in source_by_tag.keys() | target_by_tag.keys()
            if tag not in source_by_tag
            or tag not in target_by_tag
            or _xml_value_signature(source_by_tag[tag])
            != _xml_value_signature(target_by_tag[tag])
        }

        # These composite leaf properties merge attribute-by-attribute across
        # style levels. A target-only slot would therefore leak through a
        # promoted source default unless the source supplies the same slot (or
        # its mutually-exclusive alias). Keep those cases fail-closed.
        mergeable_tags = {
            qn("w:rFonts"),
            qn("w:lang"),
            qn("w:ind"),
            qn("w:spacing"),
        }
        unsafe_target_gaps = set()
        for tag in different_tags & source_by_tag.keys() & target_by_tag.keys():
            if tag not in mergeable_tags:
                continue
            source_property = source_by_tag[tag]
            target_property = target_by_tag[tag]
            source_attributes = set(source_property.attrib)
            grouped_attributes = cls._style_property_attribute_groups(tag)
            grouped_by_attribute = {
                attribute: group
                for group in grouped_attributes
                for attribute in group
            }
            for target_attribute in target_property.attrib:
                if target_attribute in source_attributes:
                    continue
                group = grouped_by_attribute.get(target_attribute)
                if group is not None and any(
                    attribute in source_attributes for attribute in group
                ):
                    continue
                unsafe_target_gaps.add(tag)
                break
            if len(target_property) and not len(source_property):
                unsafe_target_gaps.add(tag)
            if target_property.text and not source_property.text:
                unsafe_target_gaps.add(tag)

        return (
            [
                property_element
                for property_element in source_properties
                if property_element.tag in different_tags
            ],
            [
                property_element
                for property_element in target_properties
                if property_element.tag in different_tags
            ],
            frozenset(unsafe_target_gaps),
        )

    def _inserted_doc_defaults_context(self, source_doc) -> dict:
        cache_key = id(source_doc.part)
        cached = self._academic_doc_defaults_cache.get(cache_key)
        if cached is not None:
            return cached

        source_p_properties, source_r_properties = (
            self._document_default_property_elements(source_doc)
        )
        target_p_properties, target_r_properties = (
            self._document_default_property_elements(self.doc)
        )

        source_p_root = OxmlElement("w:pPr")
        source_r_root = OxmlElement("w:rPr")
        target_p_root = OxmlElement("w:pPr")
        target_r_root = OxmlElement("w:rPr")
        for source, target, source_properties, target_properties in (
            (
                source_p_root,
                target_p_root,
                source_p_properties,
                target_p_properties,
            ),
            (
                source_r_root,
                target_r_root,
                source_r_properties,
                target_r_properties,
            ),
        ):
            for property_element in source_properties:
                source.append(deepcopy(property_element))
            for property_element in target_properties:
                target.append(deepcopy(property_element))
            self._materialize_inserted_theme_references(source_doc, source)

        source_p_differences, target_p_differences, unsafe_p_gaps = (
            self._different_document_default_properties(
                list(source_p_root),
                list(target_p_root),
            )
        )
        source_r_differences, target_r_differences, unsafe_r_gaps = (
            self._different_document_default_properties(
                list(source_r_root),
                list(target_r_root),
            )
        )
        context = {
            "differs": bool(
                source_p_differences
                or target_p_differences
                or source_r_differences
                or target_r_differences
            ),
            "source_p": tuple(
                deepcopy(child) for child in source_p_differences
            ),
            "source_r": tuple(
                deepcopy(child) for child in source_r_differences
            ),
            "target_p": tuple(
                deepcopy(child) for child in target_p_differences
            ),
            "target_r": tuple(
                deepcopy(child) for child in target_r_differences
            ),
            "unsafe_p": unsafe_p_gaps,
            "unsafe_r": unsafe_r_gaps,
        }
        self._academic_doc_defaults_cache[cache_key] = context
        return context

    @classmethod
    def _neutral_document_default_property(cls, property_element, kind: str):
        false_tags = (
            cls._ACADEMIC_RUN_DEFAULT_FALSE_TAGS
            if kind == "run"
            else cls._ACADEMIC_PARAGRAPH_DEFAULT_FALSE_TAGS
        )
        neutral = deepcopy(property_element)
        neutral.clear()
        neutral.tail = None
        if property_element.tag in false_tags:
            neutral.set(qn("w:val"), "0")
            return neutral
        if kind == "run":
            simple_values = {
                qn("w:color"): "auto",
                qn("w:highlight"): "none",
                qn("w:u"): "none",
                qn("w:vertAlign"): "baseline",
                qn("w:spacing"): "0",
                qn("w:position"): "0",
                qn("w:kern"): "0",
                qn("w:w"): "100",
                qn("w:bdr"): "nil",
                qn("w:shd"): "nil",
                qn("w:effect"): "none",
                qn("w:em"): "none",
            }
        else:
            simple_values = {
                qn("w:outlineLvl"): "9",
                qn("w:shd"): "nil",
            }
        value = simple_values.get(property_element.tag)
        if value is None:
            return None
        neutral.set(qn("w:val"), value)
        return neutral

    def _materialize_copied_paragraph_style_doc_defaults(
        self,
        source_doc,
        source_style,
        copied_style,
    ) -> None:
        context = self._inserted_doc_defaults_context(source_doc)
        if not context["differs"]:
            return
        if source_style.get(qn("w:type")) != "paragraph":
            return
        if source_style.find(qn("w:basedOn")) is not None:
            return

        for (
            kind,
            property_name,
            source_key,
            target_key,
            unsafe_key,
            property_order,
        ) in (
            (
                "paragraph",
                "pPr",
                "source_p",
                "target_p",
                "unsafe_p",
                self._ACADEMIC_PARAGRAPH_PROPERTY_ORDER,
            ),
            (
                "run",
                "rPr",
                "source_r",
                "target_r",
                "unsafe_r",
                self._ACADEMIC_RUN_PROPERTY_ORDER,
            ),
        ):
            source_templates = [deepcopy(item) for item in context[source_key]]
            source_tags = {item.tag for item in source_templates}
            target_only = [
                item
                for item in context[target_key]
                if item.tag not in source_tags
            ]
            if not source_templates and not target_only:
                continue
            if context[unsafe_key]:
                raise ValueError(
                    "inserted document defaults contain a target-only "
                    "composite formatting slot"
                )
            properties = getattr(copied_style, f"get_or_add_{property_name}")()
            existing_tags = {child.tag for child in properties}
            neutral_templates = []
            for target_property in target_only:
                if target_property.tag in existing_tags:
                    continue
                neutral = self._neutral_document_default_property(
                    target_property,
                    kind,
                )
                if neutral is None:
                    raise ValueError(
                        "inserted document defaults cannot safely override "
                        "a target-only formatting property"
                    )
                neutral_templates.append(neutral)
            self._insert_missing_style_properties(
                properties,
                (*source_templates, *neutral_templates),
                property_order,
            )

    def _inserted_default_paragraph_style_differs(self, source_doc) -> bool:
        if self._inserted_doc_defaults_context(source_doc)["differs"]:
            return True
        source_default = source_doc.styles.default(WD_STYLE_TYPE.PARAGRAPH)
        target_default = self.doc.styles.default(WD_STYLE_TYPE.PARAGRAPH)
        if source_default is None:
            return target_default is not None
        if target_default is None:
            return True
        return not self._styles_are_semantically_equal(
            source_doc,
            source_default.style_id,
            target_default.style_id,
        )

    def _inserted_default_table_style_differs(self, source_doc) -> bool:
        source_default = source_doc.styles.default(WD_STYLE_TYPE.TABLE)
        target_default = self.doc.styles.default(WD_STYLE_TYPE.TABLE)
        if source_default is None:
            return target_default is not None
        if target_default is None:
            return True
        return not self._styles_are_semantically_equal(
            source_doc,
            source_default.style_id,
            target_default.style_id,
        )

    @staticmethod
    def _word_on_off_attribute(value: str | None) -> bool | None:
        if not isinstance(value, str):
            return None
        normalized = value.strip().casefold()
        if normalized in {"1", "true", "on", "yes"}:
            return True
        if normalized in {"0", "false", "off", "no"}:
            return False
        return None

    @classmethod
    def _table_style_condition_may_apply(
        cls,
        table_properties,
        condition_type: str | None,
    ) -> bool:
        requirements = {
            "firstRow": (("firstRow", True),),
            "lastRow": (("lastRow", True),),
            "firstCol": (("firstColumn", True),),
            "lastCol": (("lastColumn", True),),
            "band1Horz": (("noHBand", False),),
            "band2Horz": (("noHBand", False),),
            "band1Vert": (("noVBand", False),),
            "band2Vert": (("noVBand", False),),
            "nwCell": (("firstRow", True), ("firstColumn", True)),
            "neCell": (("firstRow", True), ("lastColumn", True)),
            "swCell": (("lastRow", True), ("firstColumn", True)),
            "seCell": (("lastRow", True), ("lastColumn", True)),
        }.get(condition_type)
        if requirements is None:
            # wholeTable, missing types, and future conditions remain
            # conservatively active.
            return True

        table_looks = (
            table_properties.findall(qn("w:tblLook"))
            if table_properties is not None
            else []
        )
        table_look = table_looks[0] if len(table_looks) == 1 else None
        bits = None
        if table_look is not None:
            bit_mask = table_look.get(qn("w:val"))
            if (
                isinstance(bit_mask, str)
                and re.fullmatch(r"[0-9A-Fa-f]{4}", bit_mask)
            ):
                bits = int(bit_mask, 16)
        bit_by_flag = {
            "firstRow": 0x0020,
            "lastRow": 0x0040,
            "firstColumn": 0x0080,
            "lastColumn": 0x0100,
            "noHBand": 0x0200,
            "noVBand": 0x0400,
        }
        flags = {}
        for flag_name, bit in bit_by_flag.items():
            named_value = (
                table_look.get(qn(f"w:{flag_name}"))
                if table_look is not None
                else None
            )
            if named_value is not None:
                flags[flag_name] = cls._word_on_off_attribute(named_value)
            elif bits is not None:
                flags[flag_name] = bool(bits & bit)
            else:
                flags[flag_name] = None

        return not any(
            flags[flag_name] is not None
            and flags[flag_name] != expected_value
            for flag_name, expected_value in requirements
        )

    @classmethod
    def _active_table_style_property_containers(
        cls,
        style,
        table_properties,
        property_tag,
    ):
        yield from style.findall(property_tag)
        for conditional_style in style.findall(qn("w:tblStylePr")):
            if not cls._table_style_condition_may_apply(
                table_properties,
                conditional_style.get(qn("w:type")),
            ):
                continue
            yield from conditional_style.findall(property_tag)

    def _table_style_conflicts_with_document_defaults(
        self,
        source_doc,
        table,
        context,
    ) -> bool:
        table_properties = table.find(qn("w:tblPr"))
        table_style = (
            table_properties.find(qn("w:tblStyle"))
            if table_properties is not None
            else None
        )
        style_id = (
            table_style.get(qn("w:val"))
            if table_style is not None
            else None
        )
        if not style_id:
            default_style = source_doc.styles.default(WD_STYLE_TYPE.TABLE)
            style_id = (
                default_style.style_id
                if default_style is not None
                else None
            )
        if not style_id:
            return False

        paragraph_default_tags = {
            property_element.tag
            for key in ("source_p", "target_p")
            for property_element in context[key]
        }
        run_default_tags = {
            property_element.tag
            for key in ("source_r", "target_r")
            for property_element in context[key]
        }
        visited = set()
        while style_id not in visited:
            visited.add(style_id)
            style = self._style_element_by_id(source_doc, style_id)
            if style is None or style.get(qn("w:type")) != "table":
                return True
            if any(
                child.tag in paragraph_default_tags
                for properties in self._active_table_style_property_containers(
                    style,
                    table_properties,
                    qn("w:pPr"),
                )
                for child in properties
            ) or any(
                child.tag in run_default_tags
                for properties in self._active_table_style_property_containers(
                    style,
                    table_properties,
                    qn("w:rPr"),
                )
                for child in properties
            ):
                return True
            based_on = style.findall(qn("w:basedOn"))
            if len(based_on) > 1:
                raise ValueError(
                    "table style contains duplicate inheritance references"
                )
            if not based_on:
                return False
            style_id = based_on[0].get(qn("w:val"))
            if not style_id:
                raise ValueError(
                    "table style has an invalid inheritance reference"
                )
        return True

    def _retain_inserted_default_formatting(self, source_doc, root) -> None:
        doc_defaults_context = self._inserted_doc_defaults_context(source_doc)
        doc_defaults_differ = doc_defaults_context["differs"]
        if doc_defaults_differ and any(
            self._table_style_conflicts_with_document_defaults(
                source_doc,
                table,
                doc_defaults_context,
            )
            for table in root.iter(qn("w:tbl"))
        ):
            # Moving document defaults into paragraph styles would raise their
            # cascade priority above table-style and conditional formatting.
            # Reject only when those properties overlap; layout-only table
            # styles remain safe and keep the common cover-table path working.
            raise ValueError(
                "inserted table text styles cannot be isolated across "
                "different document defaults"
            )
        paragraph_default_differs = (
            self._inserted_default_paragraph_style_differs(source_doc)
        )
        table_default_differs = self._inserted_default_table_style_differs(
            source_doc
        )
        if not paragraph_default_differs and not table_default_differs:
            return
        source_default = source_doc.styles.default(WD_STYLE_TYPE.PARAGRAPH)
        default_run_properties = (
            source_default.element.find(qn("w:rPr"))
            if source_default is not None
            else None
        )
        run_templates = (
            list(default_run_properties)
            if default_run_properties is not None
            else []
        )

        if table_default_differs:
            source_table_default = source_doc.styles.default(
                WD_STYLE_TYPE.TABLE
            )
            for table in root.iter(qn("w:tbl")):
                table_properties = table.find(qn("w:tblPr"))
                table_style = (
                    table_properties.find(qn("w:tblStyle"))
                    if table_properties is not None
                    else None
                )
                if table_style is not None:
                    continue
                if source_table_default is None:
                    raise ValueError(
                        "inserted table has no resolvable default table style"
                    )
                if table_properties is None:
                    table_properties = OxmlElement("w:tblPr")
                    table.insert(0, table_properties)
                table_style = OxmlElement("w:tblStyle")
                table_style.set(
                    qn("w:val"),
                    source_table_default.style_id,
                )
                table_properties.insert(0, table_style)

        if not paragraph_default_differs:
            return

        for paragraph in root.iter(qn("w:p")):
            paragraph_properties = paragraph.find(qn("w:pPr"))
            paragraph_style_id = None
            if paragraph_properties is not None:
                paragraph_style = paragraph_properties.find(qn("w:pStyle"))
                if paragraph_style is not None:
                    paragraph_style_id = paragraph_style.get(qn("w:val"))
            if paragraph_style_id is None:
                if source_default is None:
                    raise ValueError(
                        "inserted paragraph has no resolvable default paragraph style"
                    )
                if paragraph_properties is None:
                    paragraph_properties = OxmlElement("w:pPr")
                    paragraph.insert(0, paragraph_properties)
                paragraph_style = OxmlElement("w:pStyle")
                paragraph_style.set(
                    qn("w:val"),
                    source_default.style_id,
                )
                paragraph_properties.insert(0, paragraph_style)
                continue

            inherited_run_tags = (
                self._style_chain_run_property_tags(
                    source_doc,
                    paragraph_style_id,
                )
                if paragraph_style_id
                else set()
            )
            for run in paragraph.iter(qn("w:r")):
                nearest_paragraph = run.getparent()
                while (
                    nearest_paragraph is not None
                    and nearest_paragraph.tag != qn("w:p")
                ):
                    nearest_paragraph = nearest_paragraph.getparent()
                if nearest_paragraph is not paragraph:
                    continue
                run_properties = run.find(qn("w:rPr"))
                character_style_id = None
                if run_properties is not None:
                    character_style = run_properties.find(qn("w:rStyle"))
                    if character_style is not None:
                        character_style_id = character_style.get(qn("w:val"))
                run_inherited_tags = set(inherited_run_tags)
                if character_style_id:
                    run_inherited_tags.update(
                        self._style_chain_run_property_tags(
                            source_doc,
                            character_style_id,
                        )
                    )
                missing_templates = [
                    template
                    for template in run_templates
                    if template.tag not in run_inherited_tags
                ]
                if not missing_templates:
                    continue
                if run_properties is None:
                    run_properties = OxmlElement("w:rPr")
                    run.insert(0, run_properties)
                self._insert_missing_style_properties(
                    run_properties,
                    missing_templates,
                    self._ACADEMIC_RUN_PROPERTY_ORDER,
                )

    def retain_formatting_from_default_styles(self, doc):
        """物化源默认格式，但绝不追加会覆盖直接格式的重复属性。"""
        self._retain_inserted_default_formatting(doc, doc.element)

    def _style_node_signature(self, document, style_element, *, source: bool):
        normalized = deepcopy(style_element)
        normalized.attrib.pop(qn("w:styleId"), None)
        for tag_name in ("w:name", "w:rsid", "w:basedOn", "w:next", "w:link"):
            for child in normalized.findall(qn(tag_name)):
                normalized.remove(child)
        if source:
            self._materialize_inserted_theme_references(document, normalized)
        has_numbering = next(normalized.iter(qn("w:numId")), None) is not None
        return _xml_value_signature(normalized), has_numbering

    def _styles_are_semantically_equal(
        self,
        source_doc,
        source_style_id: str,
        target_style_id: str,
    ) -> bool:
        if self._inserted_doc_defaults_context(source_doc)["differs"]:
            return False
        # Compare the two deterministic, edge-labelled graphs iteratively.
        # Keeping a bijection distinguishes a self-cycle from a two-node cycle
        # even when every local formatting node happens to be identical.
        pending = [(source_style_id, target_style_id)]
        source_to_target = {}
        target_to_source = {}
        while pending:
            current_source_id, current_target_id = pending.pop()
            paired_target = source_to_target.get(current_source_id)
            paired_source = target_to_source.get(current_target_id)
            if paired_target is not None or paired_source is not None:
                if (
                    paired_target != current_target_id
                    or paired_source != current_source_id
                ):
                    return False
                continue
            source_to_target[current_source_id] = current_target_id
            target_to_source[current_target_id] = current_source_id

            source_style = self._style_element_by_id(
                source_doc,
                current_source_id,
            )
            target_style = self._style_element_by_id(
                self.doc,
                current_target_id,
            )
            if source_style is None or target_style is None:
                if not (
                    source_style is None
                    and target_style is None
                    and current_source_id == current_target_id
                ):
                    return False
                continue

            source_signature, source_has_numbering = self._style_node_signature(
                source_doc,
                source_style,
                source=True,
            )
            target_signature, target_has_numbering = self._style_node_signature(
                self.doc,
                target_style,
                source=False,
            )
            # Numeric IDs from separate numbering Parts have no cross-package
            # identity.  Copy such a graph so the real source definition is
            # validated and remapped.
            if source_has_numbering or target_has_numbering:
                return False
            if source_signature != target_signature:
                return False

            source_dependencies = self._style_dependency_ids(source_style)
            target_dependencies = self._style_dependency_ids(target_style)
            if any(
                isinstance(edge_key, tuple)
                and edge_key
                and edge_key[0] == "internal"
                for edge_key, _style_id in (
                    *source_dependencies,
                    *target_dependencies,
                )
            ):
                return False
            if tuple(edge_key for edge_key, _style_id in source_dependencies) != tuple(
                edge_key for edge_key, _style_id in target_dependencies
            ):
                return False
            pending.extend(
                (
                    source_dependency_id,
                    target_dependency_id,
                )
                for (
                    _source_tag,
                    source_dependency_id,
                ), (
                    _target_tag,
                    target_dependency_id,
                ) in zip(source_dependencies, target_dependencies)
            )
        return True

    def _allocate_inserted_style_identity(self, source_style) -> tuple[str, str | None]:
        from docxcompose.utils import increment_name

        used_ids = self._style_index(self.doc)
        source_id = source_style.get(qn("w:styleId"))
        if not source_id:
            raise ValueError("inserted style has no style ID")
        target_id = source_id
        while (
            target_id in used_ids
            or target_id in self._academic_reserved_style_ids
        ):
            target_id = increment_name(target_id)

        name_element = source_style.find(qn("w:name"))
        source_name = (
            name_element.get(qn("w:val"))
            if name_element is not None
            else None
        )
        if not source_name:
            return target_id, source_name
        used_names = self._target_style_names()
        target_name = source_name
        while (
            target_name.casefold() in used_names
            or target_name.casefold() in self._academic_reserved_style_names
        ):
            target_name = increment_name(target_name)
        return target_id, target_name

    @staticmethod
    def _style_aliases(style_element, label: str) -> tuple[str, ...]:
        alias_nodes = style_element.findall(qn("w:aliases"))
        if len(alias_nodes) > 1:
            raise ValueError(f"{label} contains duplicate alias lists")
        if not alias_nodes:
            return ()
        value = alias_nodes[0].get(qn("w:val"))
        if value is None:
            raise ValueError(f"{label} has an invalid alias list")
        aliases = tuple(alias.strip() for alias in value.split(","))
        if any(not alias for alias in aliases):
            raise ValueError(f"{label} has an empty style alias")
        folded = [alias.casefold() for alias in aliases]
        if len(folded) != len(set(folded)):
            raise ValueError(f"{label} has duplicate style aliases")
        return aliases

    def _sanitize_copied_style_aliases(
        self,
        copied_style,
        target_name: str | None,
    ) -> None:
        alias_nodes = copied_style.findall(qn("w:aliases"))
        if not alias_nodes:
            return
        aliases = self._style_aliases(copied_style, "inserted style")
        target_style_names = self._target_style_names()
        retained = []
        for alias in aliases:
            folded = alias.casefold()
            if (
                folded in target_style_names
                or folded in self._academic_reserved_style_names
                or (target_name and folded == target_name.casefold())
            ):
                continue
            retained.append(alias)
            self._academic_reserved_style_names.add(folded)
        if retained:
            alias_nodes[0].set(qn("w:val"), ",".join(retained))
        else:
            copied_style.remove(alias_nodes[0])

    def _ensure_inserted_style_mapping(self, source_doc, source_style_id: str) -> str:
        mapped = self._academic_style_id_mapping.get(source_style_id)
        if mapped is not None:
            return mapped

        source_style = self._style_element_by_id(source_doc, source_style_id)
        candidate_id = self.mapped_style_id(source_style_id)
        candidate_style = self._style_element_by_id(self.doc, candidate_id)
        if source_style is None:
            if candidate_style is not None:
                raise ValueError(
                    "missing inserted style would bind to a target style"
                )
            self._remember_inserted_style_mapping(
                source_style_id,
                candidate_id,
            )
            return candidate_id
        if candidate_style is not None and self._styles_are_semantically_equal(
            source_doc,
            source_style_id,
            candidate_id,
        ):
            self._remember_inserted_style_mapping(
                source_style_id,
                candidate_id,
            )
            return candidate_id

        # Once a root conflicts, keep its entire source dependency closure
        # isolated.  This avoids mixing packages halfway through a chain and
        # lets arbitrarily deep/cyclic graphs be copied without Python
        # recursion.
        pending_style_ids = [source_style_id]
        copied_styles = {}
        while pending_style_ids:
            pending_style_id = pending_style_ids.pop()
            if pending_style_id in self._academic_style_id_mapping:
                continue
            pending_source_style = self._style_element_by_id(
                source_doc,
                pending_style_id,
            )
            if pending_source_style is None:
                if self._inserted_doc_defaults_context(source_doc)["differs"]:
                    raise ValueError(
                        "inserted style dependency is missing while document "
                        "defaults require isolated inheritance"
                    )
                pending_candidate_id = self.mapped_style_id(pending_style_id)
                if self._style_element_by_id(
                    self.doc,
                    pending_candidate_id,
                ) is not None:
                    raise ValueError(
                        "missing inserted style dependency would bind to a target style"
                    )
                self._remember_inserted_style_mapping(
                    pending_style_id,
                    pending_candidate_id,
                )
                continue

            target_id, target_name = self._allocate_inserted_style_identity(
                pending_source_style
            )
            self._remember_inserted_style_mapping(
                pending_style_id,
                target_id,
            )
            if target_name:
                self._academic_reserved_style_names.add(
                    target_name.casefold()
                )
            copied_styles[pending_style_id] = (
                deepcopy(pending_source_style),
                target_id,
                target_name,
            )
            dependencies = self._style_dependency_ids(pending_source_style)
            pending_style_ids.extend(
                dependency_id
                for _tag_name, dependency_id in reversed(dependencies)
                if dependency_id not in self._academic_style_id_mapping
            )

        if self._inserted_doc_defaults_context(source_doc)["differs"]:
            for pending_style_id, (
                _copied_style,
                _target_id,
                _target_name,
            ) in copied_styles.items():
                source_cursor = self._style_element_by_id(
                    source_doc,
                    pending_style_id,
                )
                if (
                    source_cursor is None
                    or source_cursor.get(qn("w:type")) != "paragraph"
                ):
                    continue
                visited = set()
                while True:
                    cursor_id = source_cursor.get(qn("w:styleId"))
                    if not cursor_id or cursor_id in visited:
                        raise ValueError(
                            "cyclic paragraph style inheritance cannot be "
                            "isolated across different document defaults"
                        )
                    visited.add(cursor_id)
                    based_on = source_cursor.findall(qn("w:basedOn"))
                    if len(based_on) > 1:
                        raise ValueError(
                            "style contains duplicate dependency references"
                        )
                    if not based_on:
                        break
                    base_id = based_on[0].get(qn("w:val"))
                    source_cursor = self._style_element_by_id(
                        source_doc,
                        base_id,
                    )
                    if (
                        source_cursor is None
                        or source_cursor.get(qn("w:type")) != "paragraph"
                    ):
                        raise ValueError(
                            "paragraph style inheritance cannot be resolved "
                            "while document defaults differ"
                        )

        prepared_styles = []
        for pending_style_id, (
            copied_style,
            target_id,
            target_name,
        ) in copied_styles.items():
            copied_style.set(qn("w:styleId"), target_id)
            # The inserted document's defaults cannot be scoped to its
            # section.  Retain the master defaults and use this copy only via
            # explicit/dependency references.
            copied_style.attrib.pop(qn("w:default"), None)
            name_element = copied_style.find(qn("w:name"))
            if name_element is not None and target_name:
                name_element.set(qn("w:val"), target_name)
            self._sanitize_copied_style_aliases(
                copied_style,
                target_name,
            )
            for _edge_key, dependency, dependency_id in self._style_dependency_references(
                copied_style
            ):
                dependency.set(
                    qn("w:val"),
                    self._academic_style_id_mapping[dependency_id],
                )
            source_style = self._style_element_by_id(
                source_doc,
                pending_style_id,
            )
            self._materialize_copied_paragraph_style_doc_defaults(
                source_doc,
                source_style,
                copied_style,
            )
            self._materialize_inserted_theme_references(
                source_doc,
                copied_style,
            )
            prepared_styles.append(copied_style)

        for copied_style in prepared_styles:
            self.doc.styles.element.append(copied_style)
            self._register_target_style_element(copied_style)
        for copied_style in prepared_styles:
            self.add_numberings(source_doc, copied_style)
        return self._academic_style_id_mapping[source_style_id]

    def renumber_bookmarks(self):
        """按起止配对重编号书签，保留嵌套书签的闭合关系。"""
        start_tag = qn("w:bookmarkStart")
        end_tag = qn("w:bookmarkEnd")
        id_attribute = qn("w:id")
        open_bookmarks: dict[str, list[int]] = {}
        next_id = 0

        for node in self.doc.element.body.iter():
            if node.tag == start_tag:
                source_id = _normalize_ooxml_integer(node.get(id_attribute))
                if source_id is None:
                    raise ValueError("bookmark start has an invalid ID")
                if next_id > MAX_OOXML_DECIMAL_NUMBER:  # pragma: no cover
                    raise ValueError("document has no available bookmark ID")
                node.set(id_attribute, str(next_id))
                open_bookmarks.setdefault(source_id, []).append(next_id)
                next_id += 1
                continue

            if node.tag != end_tag:
                continue
            source_id = _normalize_ooxml_integer(node.get(id_attribute))
            matching_starts = (
                open_bookmarks.get(source_id)
                if source_id is not None
                else None
            )
            if not matching_starts:
                raise ValueError("bookmark end does not match a start")
            node.set(id_attribute, str(matching_starts.pop()))
            if not matching_starts:
                open_bookmarks.pop(source_id, None)

        if open_bookmarks:
            raise ValueError("bookmark start does not match an end")

    def add_referenced_parts(self, src_part, dst_part, element):
        source_doc = src_part.document
        self._materialize_inserted_theme_references(source_doc, element)
        self._copy_custom_xml_parts(source_doc)
        self._rewrite_custom_xml_bindings(element)
        self._rewrite_bookmark_names_and_references(element)

        # Composer 的 ``.//*[@r:id]`` XPath 不包含传入元素本身，
        # 因此作为 body 顶层子元素的 w:altChunk 会留下悬空 r:id。
        self._copy_body_relationship_attribute(
            src_part,
            dst_part,
            element,
            qn("r:id"),
            ignored_relationship_types=frozenset(
                {RT.IMAGE, RT.HEADER, RT.FOOTER}
            ),
        )

        # 旧式嵌入对象使用 o:relid，Composer 仅扫描 r:id，会使
        # OLEObject 节点保留却丢失 embeddings Part。
        legacy_relationship_attribute = qn("o:relid")
        for node in element.iter():
            if node.tag == qn("v:imagedata"):
                continue
            self._copy_body_relationship_attribute(
                src_part,
                dst_part,
                node,
                legacy_relationship_attribute,
            )

        super().add_referenced_parts(src_part, dst_part, element)
        _copy_composer_story_references(
            self,
            source_doc,
            element,
            reference_tags=(
                "w:commentRangeStart",
                "w:commentRangeEnd",
                "w:commentReference",
            ),
            relationship_type=RT.COMMENTS,
            content_type=CT.WML_COMMENTS,
            root_tag="w:comments",
            item_tag="w:comment",
            default_partname="/word/comments.xml",
            next_partname_template="/word/comments%d.xml",
        )
        _copy_composer_story_references(
            self,
            source_doc,
            element,
            reference_tags=("w:endnoteReference",),
            relationship_type=RT.ENDNOTES,
            content_type=CT.WML_ENDNOTES,
            root_tag="w:endnotes",
            item_tag="w:endnote",
            default_partname="/word/endnotes.xml",
            next_partname_template="/word/endnotes%d.xml",
            copy_special_notes=True,
        )

    def add_footnotes(self, doc, element):
        _copy_composer_story_references(
            self,
            doc,
            element,
            reference_tags=("w:footnoteReference",),
            relationship_type=RT.FOOTNOTES,
            content_type=CT.WML_FOOTNOTES,
            root_tag="w:footnotes",
            item_tag="w:footnote",
            default_partname="/word/footnotes.xml",
            next_partname_template="/word/footnotes%d.xml",
            copy_special_notes=True,
        )


def _create_structure_preserving_composer(master_doc, composer_class=None):
    """创建保留批注/脚注/尾注关联的 docxcompose Composer。"""
    if composer_class is None:
        from docxcompose.composer import Composer

        composer_class = Composer

    class StructurePreservingComposer(
        _StructurePreservingComposerMixin,
        composer_class,
    ):
        pass

    return StructurePreservingComposer(master_doc, preserve_styles=True)


_DOCPROPERTY_FIELD_NAME_RE = re.compile(
    r'(?P<prefix>\bDOCPROPERTY\s+)'
    r'(?:(?:"(?P<quoted>(?:""|[^"])*)")|(?P<plain>[^\s\\]+))',
    flags=re.IGNORECASE,
)
_CUSTOM_PROPERTIES_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
)
_CUSTOM_PROPERTIES_ROOT_TAG = f"{{{_CUSTOM_PROPERTIES_NAMESPACE}}}Properties"
_CUSTOM_PROPERTY_TAG = f"{{{_CUSTOM_PROPERTIES_NAMESPACE}}}property"
_CUSTOM_PROPERTY_NAME_LIMIT = 255


def _xml_value_signature(element) -> tuple:
    """生成与命名空间前缀无关、保留标量值语义的递归 XML 签名。"""
    attributes = tuple(sorted(element.attrib.items()))
    children = tuple(_xml_value_signature(child) for child in element)
    text = element.text
    if children and (text is None or not text.strip()):
        text = None
    return element.tag, attributes, text, children


def _custom_property_signature(property_element) -> tuple:
    """返回除名称和 PID 外可用于判断属性语义是否相同的稳定签名。"""
    attributes = tuple(
        sorted(
            (name, value)
            for name, value in property_element.attrib.items()
            if name not in {"name", "pid"}
        )
    )
    children = tuple(
        _xml_value_signature(child)
        for child in property_element
    )
    return attributes, children


def _validated_custom_property_elements(properties_root, label: str) -> list:
    """校验并返回名称唯一、值节点完整的自定义属性。"""
    properties = list(properties_root.findall(_CUSTOM_PROPERTY_TAG))
    used_names = set()
    for property_element in properties:
        name = property_element.get("name")
        if not name or len(name) > _CUSTOM_PROPERTY_NAME_LIMIT:
            raise ValueError(f"{label} custom property has an invalid name")
        folded_name = name.casefold()
        if folded_name in used_names:
            raise ValueError(f"{label} custom properties contain duplicate names")
        if len(property_element) != 1:
            raise ValueError(f"{label} custom property has an invalid value")
        used_names.add(folded_name)
    return properties


def _allocate_inserted_custom_property_name(
    source_name: str,
    used_names: set[str],
) -> str:
    """为同名异值的插入属性生成确定且不区分大小写唯一的名称。"""
    suffix_index = 1
    while True:
        suffix = f"__inserted_{suffix_index}"
        base = source_name[: max(1, _CUSTOM_PROPERTY_NAME_LIMIT - len(suffix))]
        candidate = f"{base}{suffix}"
        if candidate.casefold() not in used_names:
            return candidate
        suffix_index += 1


def _rewrite_docproperty_instruction(
    instruction: str | None,
    renamed_properties: dict[str, str],
) -> str | None:
    """改写字段指令中的属性名，同时保留其余开关和格式。"""
    if not instruction or not renamed_properties:
        return instruction

    def replace(match):
        source_name = match.group("quoted")
        if source_name is not None:
            source_name = source_name.replace('""', '"')
        else:
            source_name = match.group("plain")
        target_name = renamed_properties.get(source_name.casefold())
        if target_name is None:
            return match.group(0)
        if match.group("quoted") is None:
            return f'{match.group("prefix")}{target_name}'
        escaped_name = target_name.replace('"', '""')
        return f'{match.group("prefix")}"{escaped_name}"'

    return _DOCPROPERTY_FIELD_NAME_RE.sub(replace, instruction)


def _rewrite_split_docproperty_instruction(
    instruction_nodes,
    renamed_properties: dict[str, str],
) -> None:
    """把跨多个 ``instrText`` 的复杂字段作为一个逻辑指令改写。"""
    if not instruction_nodes:
        return
    original_fragments = [node.text or "" for node in instruction_nodes]
    original = "".join(original_fragments)
    rewritten = _rewrite_docproperty_instruction(original, renamed_properties)
    if rewritten == original:
        return

    remaining = rewritten
    for index, (node, fragment) in enumerate(
        zip(instruction_nodes, original_fragments)
    ):
        if index == len(instruction_nodes) - 1:
            node.text = remaining
            break
        fragment_length = min(len(fragment), len(remaining))
        node.text = remaining[:fragment_length]
        remaining = remaining[fragment_length:]


def _rewrite_inserted_docproperty_fields(
    source_doc,
    renamed_properties: dict[str, str],
) -> None:
    """在文档各 story Part 中同步改写冲突的 DOCPROPERTY 字段。"""
    if not renamed_properties:
        return

    roots_and_parts = [(source_doc.element.body, None)]
    seen_parts = set()
    for relationship in source_doc.part.rels.values():
        if (
            relationship.is_external
            or relationship.reltype
            not in {
                RT.HEADER,
                RT.FOOTER,
                RT.FOOTNOTES,
                RT.ENDNOTES,
                RT.COMMENTS,
            }
        ):
            continue
        part = relationship.target_part
        if id(part) in seen_parts:
            continue
        seen_parts.add(id(part))
        roots_and_parts.append((_composer_xml_part_root(part), part))

    for root, part in roots_and_parts:
        for simple_field in root.iter(qn("w:fldSimple")):
            instruction_name = qn("w:instr")
            instruction = simple_field.get(instruction_name)
            rewritten = _rewrite_docproperty_instruction(
                instruction,
                renamed_properties,
            )
            if rewritten != instruction:
                simple_field.set(instruction_name, rewritten)

        processed_instruction_nodes = set()
        field_stack = []
        for node in root.iter():
            if node.tag == qn("w:fldChar"):
                field_type = node.get(qn("w:fldCharType"))
                if field_type == "begin":
                    field_stack.append([])
                elif field_type in {"separate", "end"} and field_stack:
                    instruction_nodes = field_stack[-1]
                    _rewrite_split_docproperty_instruction(
                        instruction_nodes,
                        renamed_properties,
                    )
                    processed_instruction_nodes.update(instruction_nodes)
                    field_stack[-1] = []
                    if field_type == "end":
                        field_stack.pop()
                continue

            if node.tag == qn("w:instrText") and field_stack:
                field_stack[-1].append(node)

        for instruction_nodes in field_stack:
            _rewrite_split_docproperty_instruction(
                instruction_nodes,
                renamed_properties,
            )
            processed_instruction_nodes.update(instruction_nodes)

        for instruction_text in root.iter(qn("w:instrText")):
            if instruction_text in processed_instruction_nodes:
                continue
            rewritten = _rewrite_docproperty_instruction(
                instruction_text.text,
                renamed_properties,
            )
            if rewritten != instruction_text.text:
                instruction_text.text = rewritten

        if part is not None:
            _commit_composer_xml_part_root(part, root)


def _merge_inserted_custom_properties(source_doc, target_doc) -> None:
    """合并插入文档的自定义属性，并保留两边可更新的属性字段。

    自定义属性名在 Office 中按不区分大小写处理。同名同值时复用 master
    属性；同名异值时为插入侧属性分配新名称，并同步改写其字段指令。
    """
    try:
        source_part = source_doc.part.package.part_related_by(RT.CUSTOM_PROPERTIES)
    except KeyError:
        return
    if source_part.content_type != CT.OFC_CUSTOM_PROPERTIES:
        raise ValueError("custom properties relationship has the wrong content type")

    source_root = _parse_untrusted_ooxml_part(source_part.blob)
    if source_root.tag != _CUSTOM_PROPERTIES_ROOT_TAG:
        raise ValueError("custom properties Part has the wrong root")
    source_properties = _validated_custom_property_elements(
        source_root,
        "source",
    )

    target_package = target_doc.part.package
    try:
        target_part = target_package.part_related_by(RT.CUSTOM_PROPERTIES)
    except KeyError:
        partname = _allocate_package_partname(
            target_package,
            PackURI("/docProps/custom.xml"),
            "/docProps/custom%d.xml",
        )
        target_root = etree.Element(
            _CUSTOM_PROPERTIES_ROOT_TAG,
            nsmap=source_root.nsmap,
        )
        target_part = Part(
            partname,
            CT.OFC_CUSTOM_PROPERTIES,
            _serialize_composer_xml_root(target_root),
            target_package,
        )
        target_package.relate_to(target_part, RT.CUSTOM_PROPERTIES)
    else:
        if target_part.content_type != CT.OFC_CUSTOM_PROPERTIES:
            raise ValueError(
                "target custom properties relationship has the wrong content type"
            )
        target_root = _parse_untrusted_ooxml_part(target_part.blob)

    if target_root.tag != _CUSTOM_PROPERTIES_ROOT_TAG:
        raise ValueError("target custom properties Part has the wrong root")
    target_properties = _validated_custom_property_elements(
        target_root,
        "target",
    )

    target_by_name = {}
    used_names = set()
    for property_element in target_properties:
        name = property_element.get("name")
        folded_name = name.casefold()
        used_names.add(folded_name)
        target_by_name[folded_name] = property_element

    renamed_properties = {}
    changed = False
    for source_property in source_properties:
        source_name = source_property.get("name")
        folded_name = source_name.casefold()
        existing = target_by_name.get(folded_name)
        if (
            existing is not None
            and _custom_property_signature(existing)
            == _custom_property_signature(source_property)
        ):
            continue

        copied_property = deepcopy(source_property)
        target_name = source_name
        if existing is not None:
            target_name = _allocate_inserted_custom_property_name(
                source_name,
                used_names,
            )
            renamed_properties.setdefault(folded_name, target_name)

        copied_property.set("name", target_name)
        target_root.append(copied_property)
        target_by_name[target_name.casefold()] = copied_property
        used_names.add(target_name.casefold())
        changed = True

    for index, property_element in enumerate(
        target_root.findall(_CUSTOM_PROPERTY_TAG),
        start=2,
    ):
        if index > MAX_OOXML_DECIMAL_NUMBER:  # pragma: no cover - XML 大小已先受限
            raise ValueError("custom properties Part has no available PID")
        expected_pid = str(index)
        if property_element.get("pid") != expected_pid:
            property_element.set("pid", expected_pid)
            changed = True

    if changed:
        serialized_target_root = _serialize_composer_xml_root(target_root)
        if getattr(target_part, "element", None) is not None:
            target_part._element = parse_xml(serialized_target_root)
            if hasattr(target_part, "_comments"):
                target_part._comments = target_part._element
        else:
            target_part._blob = serialized_target_root
    _rewrite_inserted_docproperty_fields(source_doc, renamed_properties)


def _migrate_composer_story_references_in_part(
    composer,
    source_doc,
    source_part,
) -> None:
    """迁移页眉/页脚等独立 story 中的批注、尾注和脚注引用。"""
    source_root = getattr(source_part, "element", None)
    if source_root is None:
        return

    _copy_composer_story_references(
        composer,
        source_doc,
        source_root,
        reference_tags=(
            "w:commentRangeStart",
            "w:commentRangeEnd",
            "w:commentReference",
        ),
        relationship_type=RT.COMMENTS,
        content_type=CT.WML_COMMENTS,
        root_tag="w:comments",
        item_tag="w:comment",
        default_partname="/word/comments.xml",
        next_partname_template="/word/comments%d.xml",
    )
    _copy_composer_story_references(
        composer,
        source_doc,
        source_root,
        reference_tags=("w:endnoteReference",),
        relationship_type=RT.ENDNOTES,
        content_type=CT.WML_ENDNOTES,
        root_tag="w:endnotes",
        item_tag="w:endnote",
        default_partname="/word/endnotes.xml",
        next_partname_template="/word/endnotes%d.xml",
        copy_special_notes=True,
    )
    _copy_composer_story_references(
        composer,
        source_doc,
        source_root,
        reference_tags=("w:footnoteReference",),
        relationship_type=RT.FOOTNOTES,
        content_type=CT.WML_FOOTNOTES,
        root_tag="w:footnotes",
        item_tag="w:footnote",
        default_partname="/word/footnotes.xml",
        next_partname_template="/word/footnotes%d.xml",
        copy_special_notes=True,
    )


def _restore_inserted_header_footer_parts(
    composer,
    source_doc,
    target_doc,
    inserted_section_count: int,
) -> None:
    """恢复 docxcompose 默认移除的前插文档页眉/页脚定义。"""
    if inserted_section_count <= 0:
        return

    source_sections = list(source_doc.sections)[:inserted_section_count]
    target_sections = list(target_doc.sections)[:inserted_section_count]
    relationship_ids = {}
    reference_specs = {
        qn("w:headerReference"): RT.HEADER,
        qn("w:footerReference"): RT.FOOTER,
    }
    migrated_story_parts = set()

    for source_section, target_section in zip(source_sections, target_sections):
        _strip_header_footer_references(target_section)
        source_references = [
            child
            for child in source_section._sectPr
            if child.tag in reference_specs
        ]
        for insert_index, source_reference in enumerate(source_references):
            source_rid = source_reference.get(qn("r:id"))
            relationship = source_section._document_part.rels.get(source_rid)
            if (
                not source_rid
                or relationship is None
                or relationship.is_external
                or relationship.reltype != reference_specs[source_reference.tag]
            ):
                raise ValueError("invalid header/footer relationship in inserted document")

            source_story_part = relationship.target_part
            if id(source_story_part) not in migrated_story_parts:
                source_story_root = getattr(source_story_part, "element", None)
                if source_story_root is not None:
                    composer._retain_inserted_default_formatting(
                        source_doc,
                        source_story_root,
                    )
                    composer._materialize_inserted_theme_references(
                        source_doc,
                        source_story_root,
                    )
                    composer.add_styles(source_doc, source_story_root)
                    composer.add_numberings(source_doc, source_story_root)
                _rewrite_composer_custom_xml_bindings_in_part(
                    composer,
                    source_story_part,
                )
                _migrate_composer_story_references_in_part(
                    composer,
                    source_doc,
                    source_story_part,
                )
                migrated_story_parts.add(id(source_story_part))

            target_rid = relationship_ids.get(source_rid)
            if target_rid is None:
                target_relationship = composer.add_relationship(
                    source_section._document_part,
                    target_section._document_part,
                    relationship,
                )
                target_rid = target_relationship.rId
                relationship_ids[source_rid] = target_rid

            target_reference = deepcopy(source_reference)
            target_reference.set(qn("r:id"), target_rid)
            target_section._sectPr.insert(insert_index, target_reference)


def _materialize_even_header_footer_from_defaults(document) -> None:
    """在启用全局奇偶页分离前，把原 default story 物化为 even。

    对原本关闭 odd/even 的文档，默认页眉/页脚同时用于奇偶页。
    合并另一份开启该开关的文档时，若不显式物化，该文档的偶数
    页可能变成空白或继承到边界另一侧的 even story。
    """
    for section in document.sections:
        section_properties = section._sectPr
        document_part = section._document_part
        for default_story, relationship_type, add_reference, get_reference in (
            (
                section.header,
                RT.HEADER,
                section_properties.add_headerReference,
                section_properties.get_headerReference,
            ),
            (
                section.footer,
                RT.FOOTER,
                section_properties.add_footerReference,
                section_properties.get_footerReference,
            ),
        ):
            if get_reference(WD_HEADER_FOOTER.EVEN_PAGE) is not None:
                continue
            default_part = default_story.part
            relationship_id = document_part.relate_to(
                default_part,
                relationship_type,
            )
            add_reference(WD_HEADER_FOOTER.EVEN_PAGE, relationship_id)


_CONCAT_DOCUMENT_LOAD_ERRORS = (
    InvalidInputDocxError,
    PackageNotFoundError,
    BadZipFile,
    InvalidXmlError,
    OxmlInvalidXmlError,
    etree.XMLSyntaxError,
    KeyError,
    TypeError,
    ValueError,
)
_CONCAT_DOCUMENT_STRUCTURE_ERRORS = (
    InvalidXmlError,
    OxmlInvalidXmlError,
    etree.XMLSyntaxError,
    KeyError,
    ValueError,
)


def _load_concatenation_document(path: str | Path, label: str):
    """读取一个拼接输入，并把已知包/OOXML 错误转换为稳定的用户错误。"""
    try:
        return _load_validated_document(path)
    except _CONCAT_DOCUMENT_LOAD_ERRORS as exc:
        raise DocumentConcatError(
            f"{label}无法打开，可能已损坏或不是有效的 .docx 文件，"
            "请用 Word 重新导出后再试。"
        ) from exc


def _compose_concatenation_documents(
    cover_doc,
    body_doc,
    *,
    composer_class,
    progress_callback=None,
    restart_body_page_number: bool,
):
    """执行只依赖两个输入文档的组合步骤，返回 composer 与页码状态。"""
    body_section_count = len(body_doc.sections)

    emit_progress(progress_callback, 2, "正在把封面封装为独立的一节")
    # 先清掉封面尾部多余的空段落，再加“下一页”分节符，避免衔接处多出空白页；
    # 分节符让封面成为独立一节、正文据此从新的一页开始。
    _strip_trailing_blank_paragraphs(cover_doc)
    cover_doc.add_section(WD_SECTION.NEW_PAGE)

    if not restart_body_page_number:
        cover_uses_even_stories = (
            cover_doc.settings.odd_and_even_pages_header_footer
        )
        body_uses_even_stories = (
            body_doc.settings.odd_and_even_pages_header_footer
        )
        if cover_uses_even_stories or body_uses_even_stories:
            if not cover_uses_even_stories:
                _materialize_even_header_footer_from_defaults(cover_doc)
            if not body_uses_even_stories:
                _materialize_even_header_footer_from_defaults(body_doc)
            body_doc.settings.odd_and_even_pages_header_footer = True

    emit_progress(progress_callback, 3, "正在把封面拼接到正文之前")
    _merge_inserted_custom_properties(cover_doc, body_doc)
    composer = _create_structure_preserving_composer(
        body_doc,
        composer_class=composer_class,
    )
    composer.insert(0, cover_doc, remove_property_fields=False)

    # 正文的各节始终排在合并文档的末尾，据此定位正文首节的位置。
    merged_sections = list(body_doc.sections)
    body_first_index = len(merged_sections) - body_section_count

    page_number_restarted = False
    if restart_body_page_number and 0 <= body_first_index < len(merged_sections):
        emit_progress(progress_callback, 4, "正在让封面不计页码、正文从第 1 页重新编号")
        # 封面所在的各节：去掉页眉/页脚引用，使封面不显示页码。
        for index in range(body_first_index):
            _strip_header_footer_references(merged_sections[index])
        # 正文首节：页码从第 1 页开始。
        _set_section_page_number_start(merged_sections[body_first_index], 1)
        page_number_restarted = True
    elif 0 <= body_first_index < len(merged_sections):
        # docxcompose 会移除被插入文档的页眉/页脚引用。
        # 未请求“封面无页码”时，必须把第一份文档的定义恢复。
        _restore_inserted_header_footer_parts(
            composer,
            cover_doc,
            body_doc,
            body_first_index,
        )

    return composer, page_number_restarted


def concatenate_documents(
    first_path: str,
    second_path: str,
    output_path: str,
    progress_callback=None,
    restart_body_page_number: bool = True,
    max_output_bytes=None,
):
    """
    纯拼接两个 Word 文档，完整保留正文原有排版形态。

    典型场景：把封面（first）接到已排版好的正文（second）前面，合成一个文档。
    与 merge_cover_and_body 不同，这里不会对任何一个文档做排版处理。

    实现要点：
    - 以“正文”为主文档（docxcompose 仅会简化被插入文档的节属性，主文档的
      页边距、页眉页脚、页码字段等会完整保留），再把封面插入到最前面；
    - 在封面末尾加一个“下一页”分节符，使封面成为独立的一节，正文据此从
      新的一页开始；
    - 用 docxcompose 在文档级别合并主体、样式、编号与图片关系，避免直接
      复制 XML（或让模型手动拼接）导致的样式错乱、乱码。

    Args:
        first_path: 排在前面的文档（例如封面）路径
        second_path: 排在后面、需完整保留排版的文档（例如已排版好的正文）路径
        output_path: 拼接结果输出路径
        progress_callback: 进度回调（用于 SSE 推送）
        restart_body_page_number: 是否让封面不计入页码、正文从第 1 页重新
            编号（默认 True，契合“封面 + 正文”场景）。仅调整页码与封面页眉
            页脚引用，不改动正文任何排版形态。
        max_output_bytes: 可选的最终输出 ZIP 硬上限

    Raises:
        DocumentConcatError: 任一文档损坏或不是有效的 .docx 文件时抛出
            （便于上层返回明确的提示）。
        OutputSizeLimitExceeded: 最终文档超过配置的单文件上限

    Returns:
        拼接结果 dict（成功）或 False（缺少前置条件 / 未知失败）
    """
    first_file = Path(first_path)
    second_file = Path(second_path)

    if not first_file.exists():
        logger.error(f"第一个文档不存在：{format_log_path(first_path)}")
        return False

    if not second_file.exists():
        logger.error(f"第二个文档不存在：{format_log_path(second_path)}")
        return False

    try:
        from docxcompose.composer import Composer
    except ImportError:
        logger.error("未安装 docxcompose，无法拼接文档")
        return False

    try:
        emit_progress(progress_callback, 1, "正在读取两个文档")
        # 以正文为主文档，完整保留其页边距、页眉页脚与页码字段。
        body_doc = _load_concatenation_document(second_path, "第二个文档")
        cover_doc = _load_concatenation_document(first_path, "第一个文档")
        try:
            composer, page_number_restarted = _compose_concatenation_documents(
                cover_doc,
                body_doc,
                composer_class=Composer,
                progress_callback=progress_callback,
                restart_body_page_number=restart_body_page_number,
            )
        except _CONCAT_DOCUMENT_STRUCTURE_ERRORS as exc:
            raise DocumentConcatError(
                "其中一个文档包含无效的 Word 结构，请用 Word 分别重新导出后再试。"
            ) from exc

        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        emit_progress(progress_callback, 4, "正在生成拼接后的文档")
        _save_with_output_limit(
            composer.save,
            output_file,
            max_output_bytes,
            validate_callable=is_valid_generated_docx,
        )

        logger.info(
            f"拼接完成！{format_log_path(first_path)} + "
            f"{format_log_path(second_path)} → {format_log_path(output_path)}"
        )
        return {
            "concatenated": True,
            "restart_body_page_number": bool(restart_body_page_number),
            "page_number_restarted": page_number_restarted,
        }

    except DocumentConcatError:
        raise
    except OutputSizeLimitExceeded:
        raise
    except Exception as e:
        logger.error(
            f"拼接文档失败: {format_log_exception(e, first_path, second_path, output_path)}",
            exc_info=logger.isEnabledFor(logging.DEBUG),
        )
        return False


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
