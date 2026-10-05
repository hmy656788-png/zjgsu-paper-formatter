#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
学术论文自动排版工具 - Web 服务端
基于 Flask 提供文件上传、排版处理、结果下载功能。
"""

import codecs
import errno
import re
import os
import json
import math
import secrets
import shutil
import socket
import stat
import time
import threading
import logging
import uuid
from collections.abc import Iterable, Mapping
from contextlib import contextmanager, suppress
from dataclasses import replace
from itertools import islice
from importlib.metadata import PackageNotFoundError, version as get_package_version
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urljoin, urlsplit

import docx_validation as _docx_validation

from flask import (
    Flask,
    Response,
    g,
    has_request_context,
    jsonify,
    request,
    send_file,
    send_from_directory,
)
from werkzeug.exceptions import HTTPException
from werkzeug.exceptions import RequestEntityTooLarge, RequestedRangeNotSatisfiable
from werkzeug.wsgi import FileWrapper

from format_paper import (
    ParagraphType,
    DocumentConcatError,
    MAX_TEXT_PARAGRAPHS,
    OutputSizeLimitExceeded,
    format_log_path,
    format_log_exception,
    format_academic_paper,
    format_academic_paper_from_text,
    get_active_temp_paths,
    remove_expired_inactive_temp_path,
    text_exceeds_paragraph_limit,
    merge_cover_and_body,
    concatenate_documents,
)


def _safe_urlsplit(value):
    """Return a parsed URL or ``None`` for malformed/non-string input.

    ``urllib.parse.urlsplit`` raises for a handful of attacker-controlled
    values (for example an unmatched IPv6 bracket).  Requirements metadata is
    optional diagnostic input, so one malformed include should never bubble a
    ``TypeError``/``ValueError`` into application startup or a health check.
    """
    if not isinstance(value, str):
        return None
    # Embedded control characters can be interpreted differently by URL
    # consumers (and may become header-injection material for ``Request``).
    if any(not character.isprintable() for character in value):
        return None
    try:
        return urlsplit(value)
    except (TypeError, ValueError, UnicodeError):
        return None


def _validate_remote_response_target(response) -> None:
    """Reject unsafe redirect targets returned by an HTTP client.

    ``urllib`` follows redirects automatically.  A requirements URL may be
    supplied by deployment configuration, so a redirect to ``file:``, an
    embedded-credential URL, or another non-HTTP scheme must never be treated
    as a valid requirements source.  Lightweight response doubles used by
    callers/tests may not expose ``geturl``; in that case the original
    validated request target remains the only available authority.
    """
    get_url = getattr(response, "geturl", None)
    if not callable(get_url):
        return
    try:
        final_url = get_url()
    except Exception as exc:
        raise ValueError("remote requirements response target is unavailable") from exc
    # Real urllib responses return a concrete string.  A few compatible
    # response adapters expose a dynamic ``geturl`` attribute but return
    # ``None`` or another sentinel when no final URL is available; in that
    # case the already validated request target remains the authority.
    if not isinstance(final_url, str):
        return
    parsed = _safe_urlsplit(final_url)
    if (
        parsed is None
        or parsed.scheme.lower() not in {"http", "https"}
        or not parsed.netloc
        or parsed.username is not None
        or parsed.password is not None
    ):
        raise ValueError("remote requirements redirect target is unsafe")


def _strip_surrounding_quotes(value: str) -> str:
    """去除成对出现的单引号或双引号。"""
    if not isinstance(value, str):
        return ""
    value = value.strip()
    if len(value) >= 2 and (value[0] == value[-1]) and value[0] in {"'", '"'}:
        return value[1:-1]
    return value


def _is_remote_requirements_source(requirement_path: str) -> bool:
    """判断字符串是否看起来像 HTTP(S) requirements 文件来源。"""
    parsed = _safe_urlsplit(requirement_path)
    return bool(
        parsed is not None
        and parsed.scheme.lower() in {"http", "https"}
        and bool(parsed.netloc)
    )


def _is_github_contents_api_source(requirement_path: str) -> bool:
    """判断是否为需要 GitHub raw media type 的 Contents API 文件地址。"""
    parsed = _safe_urlsplit(requirement_path)
    if parsed is None:
        return False
    try:
        hostname = (parsed.hostname or "").casefold()
        # Credentials and non-standard ports are not part of the canonical
        # GitHub API source.  Rejecting them also prevents accidentally
        # forwarding deployment secrets to a URL that merely resembles the
        # API host.
        has_credentials = bool(parsed.username or parsed.password)
        port = parsed.port
    except (TypeError, ValueError, UnicodeError):
        return False
    return (
        parsed.scheme.lower() in {"http", "https"}
        and hostname == "api.github.com"
        and not has_credentials
        and port in {None, 80, 443}
        and re.fullmatch(r"/repos/[^/]+/[^/]+/contents/.+", parsed.path) is not None
        and not parsed.path.endswith("/")
    )


def _normalize_remote_requirements_source(requirement_path: str) -> str:
    """将常见 GitHub 仓库与 gist 链接转换为可直接读取的 raw 链接。"""
    if not isinstance(requirement_path, str):
        return ""
    parsed = _safe_urlsplit(requirement_path)
    if parsed is None:
        return ""
    if parsed.scheme.lower() not in {"http", "https"}:
        return requirement_path
    normalized = parsed._replace(
        scheme=parsed.scheme.lower(),
        netloc=parsed.netloc.lower(),
        fragment="",
    )
    normalized_base = normalized.geturl()

    try:
        hostname = (normalized.hostname or "").casefold()
        has_credentials = bool(normalized.username or normalized.password)
        port = normalized.port
    except (TypeError, ValueError, UnicodeError):
        return ""
    # Never turn a URL carrying credentials into a request target.  For
    # ordinary non-GitHub hosts preserve the original normalized spelling for
    # backwards compatibility; the caller may still choose to reject it.
    if has_credentials:
        return normalized_base
    netloc = hostname
    if port is not None:
        netloc = f"{netloc}:{port}"
    if netloc in {"github.com", "www.github.com"}:
        match = re.match(r"^/([^/]+)/([^/]+)/(?:blob|raw|tree)/(.*)$", normalized.path)
        if match is None:
            return normalized_base

        owner, repo, tail = match.groups()
        if not tail or "/" not in tail or tail.endswith("/"):
            return normalized_base

        if repo.lower().endswith(".git"):
            repo = repo[:-4]

        return f"{normalized.scheme}://raw.githubusercontent.com/{owner}/{repo}/{tail}"

    if netloc == "api.github.com":
        match = re.match(r"^/repos/([^/]+)/([^/]+)/contents/(.+)$", normalized.path)
        if match is None:
            return normalized_base

        owner, repo, content_path = match.groups()
        if not content_path or content_path.endswith("/"):
            return normalized_base

        if repo.lower().endswith(".git"):
            repo = repo[:-4]

        api_url = (
            f"{normalized.scheme}://api.github.com/repos/"
            f"{owner}/{repo}/contents/{content_path}"
        )
        ref_values = parse_qs(parsed.query or "").get("ref")
        if not ref_values:
            return api_url

        ref = ref_values[0].strip()
        return f"{api_url}?ref={quote(ref, safe='')}" if ref else api_url

    if netloc in {"gist.github.com", "www.gist.github.com"}:
        match = re.match(r"^/([^/]+)/([^/]+)/raw/(.+)$", normalized.path)
        if match is not None:
            owner, gist_id, tail = match.groups()
            if not tail or tail.endswith("/"):
                return normalized_base

            return f"{normalized.scheme}://gist.githubusercontent.com/{owner}/{gist_id}/raw/{tail}"

        raw_root_match = re.match(
            r"^/([^/]+)/([^/]+)/raw/?$",
            normalized.path,
        )
        if raw_root_match is not None:
            owner, gist_id = raw_root_match.groups()
            return (
                f"{normalized.scheme}://gist.githubusercontent.com/"
                f"{owner}/{gist_id}/raw"
            )

        query = parse_qs(parsed.query or "")
        file_values = query.get("file")
        page_match = re.match(r"^/([^/]+)/([^/]+)$", normalized.path.rstrip("/"))
        if page_match is None:
            return normalized_base

        owner, gist_id = page_match.groups()
        if not file_values:
            if unquote(parsed.fragment or "").casefold().startswith("file-"):
                return normalized_base
            return (
                f"{normalized.scheme}://gist.githubusercontent.com/"
                f"{owner}/{gist_id}/raw"
            )

        gist_file = file_values[0].strip()
        if not gist_file:
            return (
                f"{normalized.scheme}://gist.githubusercontent.com/"
                f"{owner}/{gist_id}/raw"
            )

        gist_file = quote(gist_file, safe="/")
        return f"{normalized.scheme}://gist.githubusercontent.com/{owner}/{gist_id}/raw/{gist_file}"

    if netloc in {"raw.githubusercontent.com", "gist.githubusercontent.com"}:
        return f"{normalized.scheme}://{normalized.netloc}{normalized.path}"

    return normalized_base


_REMOTE_REQUIREMENTS_CACHE: dict[str, str] = {}
# Keep the callable used to populate each cache entry alongside the public
# string cache.  This is intentionally separate so callers/tests can continue
# to clear or inspect ``_REMOTE_REQUIREMENTS_CACHE`` as a plain mapping.  It
# also prevents a monkey-patched/rotated fetcher from accidentally reusing a
# response produced by an earlier transport (important for long-lived worker
# processes and hot-reload environments).
_REMOTE_REQUIREMENTS_CACHE_FETCHERS: dict[str, object] = {}
_GIST_FILE_SOURCE_CACHE: dict[tuple[str, str], str] = {}

# Remote requirements are only used for a small dependency-version probe, but
# this code runs in a long-lived web worker and can be reached concurrently by
# health checks.  Keep the public cache as a regular ``dict`` (a few callers
# and tests intentionally clear/inspect it) while protecting compound cache
# operations with a lock.  The in-flight table coalesces concurrent requests
# for the same URL without serialising unrelated URLs.
_REMOTE_REQUIREMENTS_CACHE_LOCK = threading.RLock()
# Requirements files are deployment metadata rather than arbitrary user
# uploads.  Keep a conservative bound nevertheless: a process that follows a
# large number of remote includes must not accumulate unbounded text or Gist
# metadata for its entire lifetime.
MAX_REMOTE_REQUIREMENTS_PAYLOAD_BYTES = 4 * 1024 * 1024
MAX_GIST_METADATA_PAYLOAD_BYTES = 512 * 1024
MAX_REMOTE_REQUIREMENTS_CACHE_ENTRIES = 64
MAX_REMOTE_REQUIREMENTS_CACHE_BYTES = 8 * 1024 * 1024
MAX_GIST_FILE_SOURCE_CACHE_ENTRIES = 128
# A failed or wedged transport must not keep every concurrent health check
# blocked forever behind one in-flight fetch.  This is deliberately longer
# than the normal five-second urllib timeout, while still bounding damage from
# a custom transport that ignores its timeout.
REMOTE_REQUIREMENTS_INFLIGHT_WAIT_SECONDS = 15.0
_GIST_FILE_SOURCE_CACHE_LOCK = threading.RLock()
_REMOTE_REQUIREMENTS_CACHE_INFLIGHT: dict[tuple[str, int], threading.Event] = {}

_REQUIREMENTS_BOM_ENCODINGS = (
    (codecs.BOM_UTF32_BE, "utf-32-be"),
    (codecs.BOM_UTF32_LE, "utf-32-le"),
    (codecs.BOM_UTF8, "utf-8"),
    (codecs.BOM_UTF16_BE, "utf-16-be"),
    (codecs.BOM_UTF16_LE, "utf-16-le"),
)
_REQUIREMENTS_ENCODING_COOKIE_RE = re.compile(br"coding[:=]\s*([-\w.]+)")


def _decode_requirements_payload(
    payload: bytes,
    response_encoding: str | None = None,
) -> str:
    """按 BOM、文件头或响应头编码解码，未知编码继续沿用 UTF-8 容错策略。"""
    if not isinstance(payload, (bytes, bytearray, memoryview)):
        return ""
    # Work with an immutable byte string so the checks below cannot observe a
    # caller mutating a bytearray while decoding it.
    payload = bytes(payload)
    # UTF-32 LE 的 BOM 以 UTF-16 LE 的 BOM 开头，因此长 BOM 必须优先匹配。
    for bom, encoding in _REQUIREMENTS_BOM_ENCODINGS:
        if payload.startswith(bom):
            return payload[len(bom):].decode(encoding, errors="replace")

    # ``payload.splitlines()[:2]`` still materializes *all* lines before the
    # slice.  Split only at the first two newlines to keep memory/time bounded
    # when a remote endpoint returns a very large single requirements file.
    cookie_probe = payload.split(b"\n", 2)[:2]
    for line in cookie_probe:
        # PEP 263 allows horizontal whitespace before the comment marker;
        # accepting it is important for generated/indented requirements files.
        # Do not call ``strip()`` here because a coding cookie embedded in a
        # non-comment token must remain ignored (the regex below is anchored
        # by the leading ``#`` after whitespace removal).
        comment_line = line.lstrip(b" \t")
        if not comment_line.startswith(b"#"):
            continue
        encoding_match = _REQUIREMENTS_ENCODING_COOKIE_RE.search(comment_line)
        if encoding_match is None:
            continue
        declared_encoding = encoding_match.group(1).decode("ascii")
        try:
            codecs.lookup(declared_encoding)
            return payload.decode(declared_encoding, errors="replace")
        except (LookupError, UnicodeError):
            break

    if response_encoding:
        try:
            return payload.decode(response_encoding, errors="replace")
        except (LookupError, UnicodeError):
            pass

    return payload.decode("utf-8", errors="replace")


def _read_response_payload(response, max_bytes: int) -> bytes:
    """Read a remote response with a hard byte ceiling.

    ``urllib`` responses accept a size argument, while several lightweight
    test doubles (and some compatible HTTP clients) expose a zero-argument
    ``read`` method.  Prefer bounded chunk reads and fall back to the latter
    only when necessary; either way reject non-byte payloads and oversized
    data before it reaches the decoder/cache.
    """
    if (
        isinstance(max_bytes, bool)
        or not isinstance(max_bytes, int)
        or max_bytes < 0
    ):
        raise ValueError("invalid remote payload limit")
    reader = getattr(response, "read", None)
    if not callable(reader):
        raise TypeError("remote response has no readable body")

    # A trustworthy Content-Length lets us fail before allocating a body.  Do
    # not rely on it exclusively: chunked responses and malicious headers are
    # common enough that the bounded reader below remains mandatory.
    headers = getattr(response, "headers", None)
    content_length = None
    for header_name in ("Content-Length", "content-length"):
        try:
            raw_length = headers.get(header_name) if headers is not None else None
        except Exception:
            raw_length = None
        if raw_length is not None and not isinstance(raw_length, (dict, list, tuple)):
            try:
                parsed_length = int(str(raw_length).strip())
            except (TypeError, ValueError, OverflowError):
                parsed_length = None
            if parsed_length is not None and parsed_length >= 0:
                content_length = parsed_length
                break
    if content_length is not None and content_length > max_bytes:
        raise ValueError("remote requirements payload exceeds limit")

    # Keep the initial request reasonably small for very large limits.  A
    # response that implements the usual ``read(size)`` contract may return a
    # short chunk before EOF; when Content-Length is available we therefore
    # continue until that declared length has been consumed.  Without a
    # declaration, however, a short first chunk is the only EOF signal exposed
    # by many lightweight response doubles (and by a few file-like adapters).
    # Treat it as complete instead of repeatedly asking a stateless ``read``
    # mock for the same bytes forever.
    initial_request_size = min(max_bytes + 1, 1024 * 1024)
    try:
        first_chunk = reader(initial_request_size)
        bounded_reader = True
    except TypeError:
        # Compatibility fallback for simple ``read()`` test doubles.  The
        # returned object is still checked before caching.
        first_chunk = reader()
        bounded_reader = False

    def coerce_chunk(chunk):
        if not isinstance(chunk, (bytes, bytearray, memoryview)):
            raise TypeError("remote response body is not bytes")
        return bytes(chunk)

    first_chunk = coerce_chunk(first_chunk)
    if len(first_chunk) > max_bytes:
        raise ValueError("remote requirements payload exceeds limit")
    chunks = [first_chunk]
    total = len(first_chunk)

    if bounded_reader and first_chunk:
        # Continue until EOF when the server tells us how many bytes to
        # expect.  Ask for one extra byte at the ceiling to detect overflow
        # without retaining a body larger than the configured limit.
        #
        # If there is no trustworthy Content-Length, only a full initial
        # chunk proves that more data may remain.  Stopping on a short chunk
        # keeps compatibility with zero-state ``read(size)`` test doubles and
        # avoids an unbounded series of duplicate reads; urllib's HTTPResponse
        # normally fills the requested amount unless EOF is reached.
        should_continue = content_length is not None or len(first_chunk) >= initial_request_size
        while should_continue and total <= max_bytes:
            remaining = max_bytes - total
            if remaining < 0:
                raise ValueError("remote requirements payload exceeds limit")
            chunk = coerce_chunk(reader(remaining + 1))
            if not chunk:
                break
            total += len(chunk)
            if total > max_bytes:
                raise ValueError("remote requirements payload exceeds limit")
            chunks.append(chunk)
            if content_length is not None and total >= content_length:
                break
            # With no declaration, another short read is treated as EOF.  A
            # full read means the response may have more bytes, so continue.
            should_continue = content_length is not None or len(chunk) >= remaining + 1

    return b"".join(chunks)


def _cache_entry_size(value: str) -> int:
    if not isinstance(value, str):
        return 0
    try:
        return len(value.encode("utf-8", errors="replace"))
    except (UnicodeError, TypeError):
        return len(value)


def _prune_remote_requirements_cache_locked() -> None:
    """Evict oldest/invalid remote cache entries; caller holds cache lock."""
    # Keep the side map in sync even when callers clear the public cache
    # mapping directly (a supported testing/diagnostic operation).
    for key in list(_REMOTE_REQUIREMENTS_CACHE_FETCHERS):
        if key not in _REMOTE_REQUIREMENTS_CACHE:
            _REMOTE_REQUIREMENTS_CACHE_FETCHERS.pop(key, None)

    def total_bytes() -> int:
        return sum(_cache_entry_size(value) for value in _REMOTE_REQUIREMENTS_CACHE.values())

    while len(_REMOTE_REQUIREMENTS_CACHE) > MAX_REMOTE_REQUIREMENTS_CACHE_ENTRIES:
        oldest_key = next(iter(_REMOTE_REQUIREMENTS_CACHE), None)
        if oldest_key is None:
            break
        _REMOTE_REQUIREMENTS_CACHE.pop(oldest_key, None)
        _REMOTE_REQUIREMENTS_CACHE_FETCHERS.pop(oldest_key, None)

    while _REMOTE_REQUIREMENTS_CACHE and total_bytes() > MAX_REMOTE_REQUIREMENTS_CACHE_BYTES:
        oldest_key = next(iter(_REMOTE_REQUIREMENTS_CACHE))
        _REMOTE_REQUIREMENTS_CACHE.pop(oldest_key, None)
        _REMOTE_REQUIREMENTS_CACHE_FETCHERS.pop(oldest_key, None)


def _remote_cache_get(key: str, fetcher):
    with _REMOTE_REQUIREMENTS_CACHE_LOCK:
        _prune_remote_requirements_cache_locked()
        if key not in _REMOTE_REQUIREMENTS_CACHE:
            return None
        cached_fetcher = _REMOTE_REQUIREMENTS_CACHE_FETCHERS.get(key)
        if cached_fetcher is not fetcher:
            _REMOTE_REQUIREMENTS_CACHE.pop(key, None)
            _REMOTE_REQUIREMENTS_CACHE_FETCHERS.pop(key, None)
            return None
        # Dict insertion order provides a tiny, dependency-free LRU: refresh
        # the key on every hit so frequently used includes survive eviction.
        value = _REMOTE_REQUIREMENTS_CACHE.pop(key)
        _REMOTE_REQUIREMENTS_CACHE[key] = value
        return value


def _remote_cache_put(key: str, value: str, fetcher) -> None:
    if not isinstance(value, str):
        return
    with _REMOTE_REQUIREMENTS_CACHE_LOCK:
        _REMOTE_REQUIREMENTS_CACHE.pop(key, None)
        _REMOTE_REQUIREMENTS_CACHE_FETCHERS.pop(key, None)
        _REMOTE_REQUIREMENTS_CACHE[key] = value
        _REMOTE_REQUIREMENTS_CACHE_FETCHERS[key] = fetcher
        _prune_remote_requirements_cache_locked()


def _claim_remote_requirements_fetch(key: str, fetcher):
    """Return ``(cached_value, event, owner)`` for a URL/fetcher pair.

    Cache reads and the decision to become the fetch owner must be atomic;
    otherwise two simultaneous health checks can both miss and issue the same
    network request.  ``event`` is shared with waiters and is signalled by the
    owner in ``_read_remote_requirements_text``'s ``finally`` block.  The
    helper deliberately keeps the public cache as a normal mapping so direct
    ``clear()`` calls remain safe.
    """
    fetcher_key = (key, id(fetcher))
    wait_deadline = time.monotonic() + REMOTE_REQUIREMENTS_INFLIGHT_WAIT_SECONDS
    while True:
        with _REMOTE_REQUIREMENTS_CACHE_LOCK:
            _prune_remote_requirements_cache_locked()
            if key in _REMOTE_REQUIREMENTS_CACHE:
                cached_fetcher = _REMOTE_REQUIREMENTS_CACHE_FETCHERS.get(key)
                if cached_fetcher is fetcher:
                    value = _REMOTE_REQUIREMENTS_CACHE.pop(key)
                    _REMOTE_REQUIREMENTS_CACHE[key] = value
                    return value, None, False
                _REMOTE_REQUIREMENTS_CACHE.pop(key, None)
                _REMOTE_REQUIREMENTS_CACHE_FETCHERS.pop(key, None)

            event = _REMOTE_REQUIREMENTS_CACHE_INFLIGHT.get(fetcher_key)
            if event is None:
                event = threading.Event()
                _REMOTE_REQUIREMENTS_CACHE_INFLIGHT[fetcher_key] = event
                return None, event, True

        # Do not hold the cache lock while waiting for network I/O.  The
        # normal transport timeout is five seconds; this slightly larger
        # bound also covers a slow test double or a short-lived connection
        # retry.  Once signalled, loop and observe the owner's cache write.
        remaining = wait_deadline - time.monotonic()
        if remaining <= 0 or not event.wait(timeout=min(10.0, remaining)):
            # Do not leave a request hanging indefinitely when the owner is a
            # broken/custom transport that never returns.  The owner remains
            # responsible for removing the in-flight marker in its finally
            # block; this waiter simply fails closed and lets its caller use
            # the normal optional-metadata fallback path.
            if time.monotonic() >= wait_deadline:
                raise TimeoutError("remote requirements fetch is still in flight")


def _gist_cache_get(key):
    with _GIST_FILE_SOURCE_CACHE_LOCK:
        value = _GIST_FILE_SOURCE_CACHE.get(key)
        if value is None:
            return None
        # Refresh insertion order for bounded LRU behaviour.
        _GIST_FILE_SOURCE_CACHE.pop(key, None)
        _GIST_FILE_SOURCE_CACHE[key] = value
        return value


def _gist_cache_put(key, value: str) -> None:
    if not isinstance(value, str):
        return
    with _GIST_FILE_SOURCE_CACHE_LOCK:
        _GIST_FILE_SOURCE_CACHE.pop(key, None)
        _GIST_FILE_SOURCE_CACHE[key] = value
        while len(_GIST_FILE_SOURCE_CACHE) > MAX_GIST_FILE_SOURCE_CACHE_ENTRIES:
            oldest_key = next(iter(_GIST_FILE_SOURCE_CACHE), None)
            if oldest_key is None:
                break
            _GIST_FILE_SOURCE_CACHE.pop(oldest_key, None)


def _github_api_request_headers(accept: str) -> dict[str, str]:
    """构造仅用于 api.github.com 的标准请求头。"""
    headers = {
        "Accept": accept,
        "User-Agent": "zjgsu-paper-formatter",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    for token_name in ("GH_TOKEN", "GITHUB_TOKEN"):
        raw_token = os.getenv(token_name, "")
        github_token = raw_token.strip()
        if github_token and "\r" not in raw_token and "\n" not in raw_token:
            headers["Authorization"] = f"Bearer {github_token}"
            break
    return headers


_GIST_FILE_LINE_FRAGMENT_RE = re.compile(
    r"-L(?:C)?\d+(?:-L(?:C)?\d+)?$",
    flags=re.IGNORECASE,
)


def _gist_filename_fragment(filename: str) -> str:
    """返回 GitHub Gist 页面用于文件容器的 fragment id。"""
    if not isinstance(filename, str):
        return "file-"
    filename_slug = re.sub(r"[^a-z0-9]+", "-", filename.casefold()).strip("-")
    return f"file-{filename_slug}"


def _resolve_gist_file_fragment_source(url: str) -> str:
    """将多文件 Gist 的文件 fragment 解析为唯一 raw URL。"""
    if not isinstance(url, str):
        return ""
    parsed = _safe_urlsplit(url)
    if parsed is None:
        return ""
    # DNS host names are case-insensitive.  ``urlsplit().hostname`` preserves
    # the spelling supplied by the caller, so normalize it before checking the
    # allow-list; otherwise an otherwise-valid ``GIST.GITHUB.COM`` URL would
    # bypass fragment resolution and fetch the HTML page as requirements text.
    try:
        hostname = (parsed.hostname or "").casefold()
    except (TypeError, ValueError, UnicodeError):
        return url
    if hostname not in {"gist.github.com", "www.gist.github.com"}:
        return url
    if parse_qs(parsed.query or "").get("file"):
        return url

    page_match = re.fullmatch(r"/([^/]+)/([^/]+)/?", parsed.path)
    fragment = unquote(parsed.fragment or "")
    if page_match is None or not fragment.casefold().startswith("file-"):
        return url

    _owner, gist_id = page_match.groups()
    file_fragment = _GIST_FILE_LINE_FRAGMENT_RE.sub("", fragment).casefold()
    cache_key = (gist_id.casefold(), file_fragment)
    cached_source = _gist_cache_get(cache_key)
    if cached_source is not None:
        return cached_source

    from urllib.request import Request, urlopen

    api_request = Request(
        f"https://api.github.com/gists/{quote(gist_id, safe='')}",
        headers=_github_api_request_headers("application/vnd.github+json"),
    )
    with urlopen(api_request, timeout=5) as response:  # noqa: S310 - 固定 GitHub API 主机
        # GitHub's API client follows redirects.  Metadata is used to select a
        # subsequent raw URL, so accepting a redirect to another host would
        # turn a trusted GitHub lookup into an attacker-controlled SSRF and
        # could make us parse arbitrary JSON as Gist metadata.
        _validate_remote_response_target(response)
        get_url = getattr(response, "geturl", None)
        if callable(get_url):
            final_url = get_url()
            parsed_final = _safe_urlsplit(final_url) if isinstance(final_url, str) else None
            if parsed_final is not None:
                try:
                    final_host = (parsed_final.hostname or "").casefold()
                    final_port = parsed_final.port
                except (TypeError, ValueError, UnicodeError) as exc:
                    raise ValueError("GitHub Gist API redirect target is unsafe") from exc
                if (
                    parsed_final.scheme.lower() not in {"http", "https"}
                    or final_host != "api.github.com"
                    or final_port not in {None, 80, 443}
                    or parsed_final.username is not None
                    or parsed_final.password is not None
                ):
                    raise ValueError("GitHub Gist API redirect target is unsafe")
        payload = _read_response_payload(response, MAX_GIST_METADATA_PAYLOAD_BYTES)

    try:
        metadata = json.loads(payload)
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("GitHub Gist metadata is not valid JSON") from exc
    files = metadata.get("files") if isinstance(metadata, Mapping) else None
    if not isinstance(files, Mapping):
        raise ValueError("GitHub Gist metadata does not contain files")

    matching_sources: list[str] = []
    for file_metadata in files.values():
        if not isinstance(file_metadata, Mapping):
            continue
        filename = file_metadata.get("filename")
        raw_url = file_metadata.get("raw_url")
        if not isinstance(filename, str) or not isinstance(raw_url, str):
            continue
        if _gist_filename_fragment(filename) != file_fragment:
            continue

        normalized_raw_url = _normalize_remote_requirements_source(raw_url)
        raw = _safe_urlsplit(normalized_raw_url)
        raw_path_parts = raw.path.split("/") if raw is not None else []
        try:
            raw_hostname = (raw.hostname or "").casefold() if raw is not None else ""
            raw_port = raw.port if raw is not None else None
            raw_username = raw.username if raw is not None else None
            raw_password = raw.password if raw is not None else None
        except (TypeError, ValueError, UnicodeError):
            raw_hostname = ""
            raw_port = object()
            raw_username = raw_password = None
        if (
            raw is None
            or raw.scheme.lower() not in {"http", "https"}
            or raw_hostname != "gist.githubusercontent.com"
            or raw_username is not None
            or raw_password is not None
            or raw_port not in {None, 80, 443}
            or len(raw_path_parts) < 4
            or raw_path_parts[2].casefold() != gist_id.casefold()
            or raw_path_parts[3] != "raw"
        ):
            continue
        matching_sources.append(normalized_raw_url)

    if len(set(matching_sources)) != 1:
        raise ValueError("Gist file fragment does not identify exactly one file")

    resolved_source = matching_sources[0]
    _gist_cache_put(cache_key, resolved_source)
    return resolved_source


def _read_remote_requirements_text(url: str) -> str:
    """从 HTTP(S) 获取 requirements 内容。"""
    if not isinstance(url, str):
        return ""
    # Reject credentials before Gist fragment resolution as well: resolving a
    # credential-bearing Gist page would otherwise perform an unnecessary API
    # request before the normalized URL guard below gets a chance to run.
    source_url = _safe_urlsplit(url)
    if source_url is None or source_url.username or source_url.password:
        return ""
    normalized_url = _normalize_remote_requirements_source(
        _resolve_gist_file_fragment_source(url)
    )
    # Credentials embedded in a requirements URL are both unnecessary (the
    # supported public sources use their own authentication mechanism) and
    # unsafe to forward through urllib, redirects, logs, or cache keys.  Drop
    # such inputs before any network request is attempted.
    parsed_url = _safe_urlsplit(normalized_url)
    if parsed_url is None or parsed_url.username or parsed_url.password:
        return ""
    # Resolve urlopen at call time so test doubles, hot-reloaded transports,
    # and a changed network implementation cannot inherit stale bytes from a
    # different fetcher.  Within one transport (the normal production case),
    # equivalent URL variants still share the cache as intended.
    import urllib.request

    current_fetcher = urllib.request.urlopen
    cached_value, fetch_event, fetch_owner = _claim_remote_requirements_fetch(
        normalized_url,
        current_fetcher,
    )
    if not fetch_owner:
        return cached_value

    from urllib.request import Request

    response_encoding = None
    request_target = normalized_url
    if _is_github_contents_api_source(normalized_url):
        request_target = Request(
            normalized_url,
            headers=_github_api_request_headers(
                "application/vnd.github.raw+json"
            ),
        )
    try:
        try:
            with current_fetcher(request_target, timeout=5) as response:  # noqa: S310 - 需要兼容环境的明确远端读取
                _validate_remote_response_target(response)
                headers = getattr(response, "headers", None)
                get_content_charset = getattr(headers, "get_content_charset", None)
                if callable(get_content_charset):
                    charset = get_content_charset()
                    if isinstance(charset, str):
                        response_encoding = charset
                payload = _read_response_payload(
                    response,
                    MAX_REMOTE_REQUIREMENTS_PAYLOAD_BYTES,
                )
        except Exception as exc:
            logger.debug("读取远端 requirements 失败：%s", normalized_url)
            if str(exc):
                logger.debug("失败原因：%s", exc)
            raise

        decoded = _decode_requirements_payload(payload, response_encoding)
        decoded = decoded.lstrip("\ufeff")
        _remote_cache_put(normalized_url, decoded, current_fetcher)
        return decoded
    finally:
        # Wake all waiters even when transport/decoding fails.  They may retry
        # independently, but must never remain blocked behind a failed owner.
        with _REMOTE_REQUIREMENTS_CACHE_LOCK:
            owner_event = _REMOTE_REQUIREMENTS_CACHE_INFLIGHT.pop(
                (normalized_url, id(current_fetcher)),
                None,
            )
            if owner_event is not None:
                owner_event.set()


def _file_requirements_url_to_path(requirement_url: str) -> Path | None:
    """按 pip/RFC 8089 规则将本地 file URL 转换为路径。"""
    parsed = urlsplit(requirement_url)
    if parsed.scheme.lower() != "file":
        return None

    netloc = parsed.netloc
    if not netloc or netloc.casefold() == "localhost":
        netloc = ""
    elif os.name == "nt":
        netloc = "\\\\" + netloc
    else:
        return None

    from urllib.request import url2pathname

    path_text = url2pathname(netloc + parsed.path)
    if (
        os.name == "nt"
        and not netloc
        and re.match(r"^/[A-Za-z]:(?:/|$)", path_text) is not None
    ):
        path_text = path_text[1:]
    return Path(path_text)


def _join_remote_requirements_source(base_url: str, requirement_path: str) -> str:
    """拼接远端 include，并为同仓库 Contents API 子路径继承 ref。"""
    # pip treats a leading slash in a VCS-hosted requirements file as a path
    # relative to the repository, rather than as a path at the HTTP origin.
    # ``urllib.parse.urljoin`` would incorrectly drop the owner/repository
    # components for raw.githubusercontent.com URLs.  Preserve those two
    # components explicitly while retaining the normal URL-join behaviour for
    # parent/child relative paths.
    base = urlsplit(base_url)
    if (
        requirement_path.startswith("/")
        and base.hostname is not None
        and base.hostname.casefold() == "raw.githubusercontent.com"
    ):
        base_parts = base.path.strip("/").split("/")
        if len(base_parts) >= 2 and all(base_parts[:2]):
            repository_root = "/".join(base_parts[:2])
            joined_url = base._replace(
                path=f"/{repository_root}/{requirement_path.lstrip('/')}",
                query="",
                fragment="",
            ).geturl()
        else:
            joined_url = urljoin(base_url, requirement_path)
    else:
        joined_url = urljoin(base_url, requirement_path)
    reference = urlsplit(requirement_path)
    if reference.scheme or reference.netloc:
        return joined_url

    if not (
        _is_github_contents_api_source(base_url)
        and _is_github_contents_api_source(joined_url)
    ):
        return joined_url

    base = urlsplit(base_url)
    joined = urlsplit(joined_url)
    base_repo = re.match(r"^/repos/([^/]+)/([^/]+)/contents/", base.path)
    joined_repo = re.match(r"^/repos/([^/]+)/([^/]+)/contents/", joined.path)
    if base_repo is None or joined_repo is None:
        return joined_url
    if tuple(value.casefold() for value in base_repo.groups()) != tuple(
        value.casefold() for value in joined_repo.groups()
    ):
        return joined_url

    joined_refs = parse_qs(joined.query or "").get("ref")
    if joined_refs and joined_refs[0].strip():
        return joined_url

    base_refs = parse_qs(base.query or "").get("ref")
    if not base_refs:
        return joined_url
    inherited_ref = base_refs[0].strip()
    if not inherited_ref:
        return joined_url

    return joined._replace(query=f"ref={quote(inherited_ref, safe='')}").geturl()


def _resolve_requirements_source(requirement_path: str, base_dir: Path | str | None) -> tuple[str, bool] | None:
    """解析 include 参数到目标来源。"""
    path_text = _strip_surrounding_quotes(requirement_path)
    if not path_text:
        return None

    if _is_remote_requirements_source(path_text):
        return _normalize_remote_requirements_source(path_text), True

    parsed_path = urlsplit(path_text)
    if parsed_path.scheme.lower() == "file":
        # 不允许远端 requirements 借 file URL 读取服务器本地文件。
        if isinstance(base_dir, str) and _is_remote_requirements_source(base_dir):
            return None
        include_path = _file_requirements_url_to_path(path_text)
        if include_path is None:
            return None
        if not include_path.is_absolute():
            local_base = Path.cwd() if base_dir is None else Path(base_dir)
            include_path = local_base / include_path
        try:
            return str(include_path.resolve()), False
        except OSError:
            return None

    if isinstance(base_dir, str) and _is_remote_requirements_source(base_dir):
        return (
            _normalize_remote_requirements_source(
                _join_remote_requirements_source(base_dir, path_text)
            ),
            True,
        )

    base_path = Path.cwd() if base_dir is None else Path(base_dir)
    include_path = Path(path_text)
    if not include_path.is_absolute():
        include_path = base_path / include_path
    try:
        include_path = include_path.resolve()
    except OSError:
        return None
    return str(include_path), False

try:
    from flask_cors import CORS
except ImportError:  # pragma: no cover - optional dependency fallback
    CORS = None

try:
    from docxcompose.composer import Composer as _DocxComposer  # noqa: F401
    HAS_DOCXCOMPOSE = True
except ImportError:  # pragma: no cover - optional dependency fallback
    HAS_DOCXCOMPOSE = False


DOCXCOMPOSE_MIN_VERSION = "2.2.0"
_DOCXCOMPOSE_MIN_VERSION_FROM_REQUIREMENTS = None
_DOCXCOMPOSE_MIN_VERSION_FROM_LOCK = None
_DOCXCOMPOSE_MIN_VERSION_WARNING_EMITTED = False
DOCXCOMPOSE_MIN_VERSION_SOURCE_LABELS = {
    "runtime": "运行时配置",
    "requirements_in": "requirements.in",
    "requirements_txt": "requirements.txt",
    "unavailable": "未配置",
}

_REQUIREMENTS_CONTROL_OPTIONS_NO_ARG = frozenset(
    {
        "--no-deps",
        "--no-index",
        "--prefer-binary",
        "--pre",
        "--require-hashes",
        "--no-require-hashes",
    }
)
_REQUIREMENTS_CONTROL_OPTIONS_WITH_ARG = frozenset(
    {
        "--all-releases",
        "--config-settings",
        "--editable",
        "--extra-index-url",
        "--find-links",
        "--hash",
        "--index-url",
        "--only-final",
        "--pypi-url",
        "--trusted-host",
        "--constraint",
        "--requirement",
        "--no-binary",
        "--only-binary",
        "--use-feature",
        "-r",
        "-c",
        "-i",
        "-f",
        "-e",
        "-C",
    }
)
_REQUIREMENTS_CONTROL_SHORT_OPTIONS_WITH_ARG = (
    "-r",
    "-c",
    "-i",
    "-f",
    "-e",
    "-C",
)
_REQUIREMENTS_CONTROL_INCLUDE_OPTIONS = frozenset(
    {
        "-r",
        "-c",
        "--requirement",
        "--constraint",
    }
)
_REQUIREMENTS_INCLUDE_MAX_DEPTH = 8
_REQUIREMENTS_ENV_VAR_RE = re.compile(
    r"(?P<var>\$\{(?P<name>[A-Z0-9_]+)\})"
)


def _expand_requirements_environment_variables(line: str) -> str:
    """按 pip 规则展开 requirements 中的 POSIX 风格环境变量。"""
    for env_var, variable_name in _REQUIREMENTS_ENV_VAR_RE.findall(line):
        value = os.getenv(variable_name)
        if value:
            line = line.replace(env_var, value)
    return line


def _iter_joined_requirements_lines(requirements_text: str):
    """按 requirements 的续行规则合并物理行。

    ``pip`` 的 ``join_lines`` historically only recognised a backslash when
    it was the very last source character.  Real-world lock files commonly
    put trailing spaces (or an inline comment) after the escape, however, so
    we accept those forms as well.  A backslash attached directly to the
    requirement token is deliberately *not* an escape: this avoids turning a
    literal value such as ``package\\   `` into a continuation by accident.

    Comment-only lines need a little care while a continuation is pending.  A
    convention used by pip-tools is to indent metadata comments; those lines
    are ignored and the pending requirement remains open.  An unindented
    comment starts a new logical line and therefore flushes the pending text.
    Blank lines are ignored while pending, matching pip's treatment of empty
    lines in a hash block.
    """

    # ``#`` is a comment marker only when separated from the preceding token
    # by whitespace, matching ``_strip_requirement_comment`` below.  Keeping
    # this helper local avoids accidentally treating VCS URL fragments such as
    # ``#subdirectory=src`` as comments.
    inline_comment_re = re.compile(r"\s+#")

    def split_continuation(line: str) -> tuple[bool, str]:
        """Return ``(is_continuation, code_without_escape)`` for one line."""
        comment_match = inline_comment_re.search(line)
        if comment_match is not None:
            code = line[: comment_match.start()].rstrip()
        else:
            code = line

        # The escape must be preceded by whitespace.  It may be followed by
        # spaces/tabs and, optionally, the comment we removed above.  Checking
        # the whitespace-stripped source (rather than the original line)
        # prevents a backslash inside an inline comment from being mistaken
        # for the escape marker while still accepting ``\\   ``.
        code_without_trailing_space = code.rstrip(" \t")
        if not re.search(r"(?<=\s)\\$", code_without_trailing_space):
            return False, line
        return True, code_without_trailing_space[:-1].rstrip()

    joined_parts: list[str] = []

    def flush_joined() -> str | None:
        if not joined_parts:
            return None
        # Continuation indentation is formatting, not part of the package
        # specification.  Joining stripped segments also keeps repeated hash
        # continuations stable (one separator instead of indentation spaces).
        value = " ".join(part.strip() for part in joined_parts if part.strip())
        joined_parts.clear()
        return value

    for raw_line in requirements_text.splitlines():
        physical_line = raw_line.lstrip("\ufeff")
        leading = physical_line[: len(physical_line) - len(physical_line.lstrip())]
        stripped = physical_line.lstrip()

        # Empty physical lines are ignored while a continuation is pending;
        # otherwise yielding them is harmless and lets the caller apply its
        # normal blank-line filtering.
        if not stripped:
            if joined_parts:
                continue
            yield physical_line
            continue

        is_comment_line = stripped.startswith("#")
        if is_comment_line:
            if joined_parts:
                if leading:
                    # Indented metadata comment: keep waiting for the next
                    # continuation segment.
                    continue
                # A column-zero comment terminates the pending logical line;
                # the comment itself is yielded so the caller can ignore it.
                pending = flush_joined()
                if pending is not None:
                    yield pending
            yield physical_line
            continue

        is_continuation, code = split_continuation(physical_line)
        if is_continuation:
            joined_parts.append(code)
            continue

        if joined_parts:
            joined_parts.append(physical_line)
            pending = flush_joined()
            if pending is not None:
                yield pending
        else:
            yield physical_line

    pending = flush_joined()
    if pending is not None:
        yield pending


def _parse_version_tuple(version: str) -> tuple[int, ...]:
    """将版本号转为整数元组，便于进行大小比较。"""
    raw_parts = re.findall(r"\d+", version)
    parts = [int(part) for part in raw_parts[:3]]
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def _is_version_at_least(actual: str, minimum: str) -> bool:
    """按版本语义判断 actual 是否不低于 minimum。"""
    try:
        from packaging.specifiers import SpecifierSet
        from packaging.version import Version

        if minimum.strip().startswith(("<", ">", "=", "!", "~")):
            return SpecifierSet(minimum).contains(Version(actual))
        return Version(actual) >= Version(minimum)
    except Exception:
        normalized_minimum = minimum.strip()
        if normalized_minimum.startswith(">="):
            return _parse_version_tuple(actual) >= _parse_version_tuple(
                normalized_minimum[2:]
            )
        if normalized_minimum.startswith(">"):
            return _parse_version_tuple(actual) > _parse_version_tuple(
                normalized_minimum[1:]
            )
        return _parse_version_tuple(actual) >= _parse_version_tuple(minimum)


def _extract_docxcompose_archive_version(requirement_source: str) -> str | None:
    """从结构有效的 wheel/sdist 文件名提取 docxcompose 精确版本。"""
    try:
        from packaging.utils import (
            canonicalize_name,
            parse_sdist_filename,
            parse_wheel_filename,
        )
    except Exception:
        return None

    requirement_source = re.split(
        r"\s+;\s*",
        requirement_source,
        maxsplit=1,
    )[0].strip()
    requirement_source = _strip_surrounding_quotes(requirement_source)
    if not requirement_source:
        return None

    parsed = urlsplit(requirement_source)
    if parsed.scheme.lower() in {"http", "https", "file"}:
        archive_path = parsed.path
    else:
        archive_path = requirement_source.split("?", 1)[0].split("#", 1)[0]
    filename = unquote(archive_path).replace("\\", "/").rsplit("/", 1)[-1]
    if not filename:
        return None

    try:
        if filename.casefold().endswith(".whl"):
            distribution_name, version, _build, _tags = parse_wheel_filename(
                filename
            )
        elif filename.casefold().endswith((".tar.gz", ".zip")):
            distribution_name, version = parse_sdist_filename(filename)
        else:
            return None
    except Exception:
        return None

    if canonicalize_name(distribution_name) != "docxcompose":
        return None
    return str(version)


def _extract_docxcompose_min_supported_version(requirement_line: str) -> str | None:
    """从单条 requirements 语句提取 docxcompose 的版本约束。"""
    try:
        from packaging.requirements import Requirement
        from packaging.utils import canonicalize_name
        from packaging.version import Version
    except Exception:
        return None

    requirement_line = _strip_requirement_comment(requirement_line)
    requirement_line = requirement_line.split(" --", 1)[0].strip()
    if not requirement_line:
        return None

    try:
        requirement = Requirement(requirement_line)
    except Exception:
        return _extract_docxcompose_archive_version(requirement_line)
    if canonicalize_name(requirement.name) != "docxcompose":
        if requirement.url is not None:
            return None
        return _extract_docxcompose_archive_version(requirement_line)
    if requirement.marker is not None:
        try:
            if not requirement.marker.evaluate():
                return None
        except Exception:
            return None

    exact_versions: list[Version] = []
    lower_bound_versions: list[tuple[Version, bool]] = []
    for spec in requirement.specifier:
        if spec.version is None:
            continue
        candidate = spec.version.strip()
        if spec.operator == "==" and candidate.endswith(".*"):
            candidate = candidate.removesuffix(".*")
        if not candidate:
            continue
        try:
            parsed_candidate = Version(candidate)
        except Exception:
            continue
        if spec.operator in {">", ">=", "~="}:
            lower_bound_versions.append(
                (parsed_candidate, spec.operator == ">")
            )
        elif spec.operator in {"==", "==="}:
            exact_versions.append(parsed_candidate)

    if exact_versions:
        return str(max(exact_versions))

    if lower_bound_versions:
        strongest_version = max(
            version for version, _strict in lower_bound_versions
        )
        strict = any(
            version == strongest_version and is_strict
            for version, is_strict in lower_bound_versions
        )
        return f">{strongest_version}" if strict else str(strongest_version)

    return _extract_docxcompose_archive_version(
        requirement.url or requirement_line
    )


def _strip_requirement_comment(requirement_line: str) -> str:
    """去除 requirements 内的行尾注释，但保留 URL 的片段标记。"""
    if "#" not in requirement_line:
        return requirement_line
    if requirement_line.lstrip().startswith("#"):
        return ""
    return re.sub(r"\s+#.*$", "", requirement_line)


def _is_requirements_control_line(requirement_line: str) -> bool:
    """判断是否为 requirements 的控制指令行。"""
    stripped = requirement_line.lstrip()
    if not stripped.startswith("-"):
        return False
    token = stripped.split(None, 1)[0]
    if token in _REQUIREMENTS_CONTROL_OPTIONS_NO_ARG:
        return True
    if token in _REQUIREMENTS_CONTROL_OPTIONS_WITH_ARG:
        return True
    if token in _REQUIREMENTS_CONTROL_SHORT_OPTIONS_WITH_ARG:
        return True
    if any(
        token.startswith(short_option + "=") or (
            token.startswith(short_option) and token != short_option
        )
        for short_option in _REQUIREMENTS_CONTROL_SHORT_OPTIONS_WITH_ARG
    ):
        return True
    return any(
        token.startswith(prefix + "=")
        for prefix in _REQUIREMENTS_CONTROL_OPTIONS_WITH_ARG
        if prefix.startswith("--")
    )


def _iter_requirement_file_entries(
    requirement_path: str,
    *,
    base_dir: Path | str | None,
    seen_include_files: set[Path | str] | None,
    include_depth: int,
) -> Iterable[str] | None:
    """读取并展开 include 的子 requirements 文件。"""
    if include_depth > _REQUIREMENTS_INCLUDE_MAX_DEPTH:
        return None

    resolved = _resolve_requirements_source(requirement_path, base_dir)
    if resolved is None:
        return None
    include_source, is_remote = resolved

    if seen_include_files is None:
        seen_include_files = set()
    # Use one representation for every source in the active include stack.
    # ``_resolve_requirements_source`` returns local paths as strings so that
    # they can share a stable key with callers that seed ``seen_include_files``
    # themselves.  Older callers may still provide ``Path`` members; accept
    # those for cycle detection and remove the equivalent value on unwind.
    include_key = (
        include_source
        if is_remote
        else str(Path(include_source))
    )
    equivalent_path_key = None if is_remote else Path(include_key)
    if include_key in seen_include_files or (
        equivalent_path_key is not None
        and equivalent_path_key in seen_include_files
    ):
        return None
    if is_remote:
        try:
            include_text = _read_remote_requirements_text(include_source)
        except Exception as exc:
            logger.debug("远端 requirements include 解析失败，已跳过：%s", include_source)
            if str(exc):
                logger.debug("原因：%s", exc)
            return None
    else:
        include_path = Path(include_source)
        if not include_path.is_file():
            return None
        try:
            # Match remote requirements decoding (BOM, PEP-263 coding cookie,
            # and a UTF-8 replacement fallback) for local include files too.
            # ``Path.read_text(encoding="utf-8")`` rejects UTF-16/CP1252
            # lockfiles and lets ``UnicodeDecodeError`` escape the include
            # iterator, turning an optional include into a whole health-check
            # failure.  Reading bytes also lets us strip a BOM consistently.
            include_text = _decode_requirements_payload(include_path.read_bytes())
        except (OSError, UnicodeError, TypeError) as exc:
            logger.debug("本地 requirements include 不可读，已跳过：%s", include_path)
            logger.debug("原因：%s", exc)
            return None

    seen_include_files.add(include_key)
    try:
        for child_entry in _iter_requirement_entries(
            include_text,
            # Keep the source URL itself as the base.  Besides making ordinary
            # relative joins work, this preserves a GitHub Contents API
            # ``?ref=...`` query so nested includes inherit the same branch.
            include_base_dir=include_source if is_remote else include_path.parent,
            seen_include_files=seen_include_files,
            include_depth=include_depth + 1,
        ):
            yield child_entry
    finally:
        # ``discard`` keeps cleanup idempotent if a nested parser exits early
        # after detecting the same cycle.  Also discard a legacy ``Path`` key
        # if a caller supplied one while this include was being expanded.
        seen_include_files.discard(include_key)
        if equivalent_path_key is not None:
            seen_include_files.discard(equivalent_path_key)


def _extract_requirements_control_option_and_arg(
    stripped_for_comment: str,
) -> tuple[str, str | None, list[str]]:
    stripped_for_comment = stripped_for_comment.lstrip()
    split_tokens = stripped_for_comment.split(None, 1)
    if not split_tokens:
        return "", None, split_tokens
    token = split_tokens[0]

    if "=" in stripped_for_comment:
        for prefix in _REQUIREMENTS_CONTROL_OPTIONS_WITH_ARG:
            if stripped_for_comment.startswith(prefix + "="):
                return prefix, stripped_for_comment[len(prefix) + 1 :].strip(), split_tokens
        for prefix in _REQUIREMENTS_CONTROL_OPTIONS_NO_ARG:
            if stripped_for_comment.startswith(prefix + "="):
                return prefix, None, split_tokens

    for short_option in _REQUIREMENTS_CONTROL_SHORT_OPTIONS_WITH_ARG:
        if token.startswith(short_option) and token != short_option:
            short_arg = token[len(short_option) :]
            if short_arg.startswith("="):
                short_arg = short_arg[1:]
            return short_option, short_arg, split_tokens

    if token in _REQUIREMENTS_CONTROL_OPTIONS_WITH_ARG:
        if len(split_tokens) >= 2:
            return token, split_tokens[1], split_tokens
        return token, None, split_tokens
    return token, None, split_tokens


def _iter_requirement_entries(
    requirements_text: str,
    *,
    include_base_dir: Path | str | None = None,
    seen_include_files: set[Path | str] | None = None,
    include_depth: int = 0,
    expand_includes: bool = True,
):
    """逐条产出可解析的依赖条目，支持 `\\` 续行。

    ``expand_includes`` is primarily useful to callers inspecting a lock
    file: direct pins in that file should be considered before a broad
    ``-r``/``-c`` include.  Include directives are still consumed (including
    their next-line argument) when expansion is disabled, so the argument is
    never accidentally emitted as a package requirement.
    """
    waiting_for_control_arg = False
    waiting_for_control_include = False
    if include_base_dir is None:
        include_base_dir = Path.cwd()
    elif isinstance(include_base_dir, str) and _is_remote_requirements_source(include_base_dir):
        pass
    else:
        include_base_dir = Path(include_base_dir)
    if seen_include_files is None:
        seen_include_files = set()
    for physical_line in _iter_joined_requirements_lines(requirements_text):
        text_line = physical_line.rstrip()
        stripped_for_comment = _strip_requirement_comment(text_line).rstrip()
        stripped_for_comment = _expand_requirements_environment_variables(
            stripped_for_comment
        )
        if not stripped_for_comment:
            # Keep waiting through indented metadata comments, but let a
            # column-zero comment terminate a pending split control argument.
            # This mirrors the continuation rule above and prevents a URL on
            # the following line from being swallowed unexpectedly.
            if (
                waiting_for_control_arg
                and physical_line.lstrip().startswith("#")
                and not physical_line[:1].isspace()
            ):
                waiting_for_control_arg = False
                waiting_for_control_include = False
            continue

        if waiting_for_control_arg:
            waiting_for_control_arg = False
            if waiting_for_control_include and stripped_for_comment:
                for child_entry in _iter_requirement_file_entries(
                    stripped_for_comment,
                    base_dir=include_base_dir,
                    seen_include_files=seen_include_files,
                    include_depth=include_depth,
                ):
                    yield child_entry
            waiting_for_control_include = False
            if stripped_for_comment:
                continue

        if _is_requirements_control_line(stripped_for_comment):
            option, option_arg, split_tokens = _extract_requirements_control_option_and_arg(
                stripped_for_comment
            )
            if option in _REQUIREMENTS_CONTROL_INCLUDE_OPTIONS:
                if option_arg is not None and expand_includes:
                    for child_entry in _iter_requirement_file_entries(
                        option_arg,
                        base_dir=include_base_dir,
                        seen_include_files=seen_include_files,
                        include_depth=include_depth,
                    ):
                        yield child_entry
                elif option_arg is None:
                    waiting_for_control_arg = True
                    waiting_for_control_include = expand_includes
                continue

            inline_short_arg = any(
                option != short_token and option.startswith(short_token)
                for short_token in _REQUIREMENTS_CONTROL_SHORT_OPTIONS_WITH_ARG
            )
            waiting_for_control_arg = (
                option in _REQUIREMENTS_CONTROL_OPTIONS_WITH_ARG
                and "=" not in stripped_for_comment.split(None, 1)[0]
                and len(split_tokens) == 1
                and not inline_short_arg
            )
            continue

        line = stripped_for_comment.strip()
        if line:
            yield line


def _read_docxcompose_min_supported_version_from_requirements() -> str | None:
    """从 requirements.in 读取 docxcompose 的最小支持版本约束。"""
    requirements_path = Path(__file__).resolve().parent / "requirements.in"
    try:
        requirements_text = requirements_path.read_text(encoding="utf-8")
    except OSError:
        return None

    for line in _iter_requirement_entries(
        requirements_text,
        include_base_dir=requirements_path.parent,
    ):
        docxcompose_min_version = _extract_docxcompose_min_supported_version(line)
        if docxcompose_min_version is None:
            continue
        return docxcompose_min_version
    return None


def _read_docxcompose_min_supported_version_from_lock() -> str | None:
    """从 requirements.txt 读取锁定的 docxcompose 版本（第一条可解析的可用于比较约束）。"""
    requirements_lock_path = Path(__file__).resolve().parent / "requirements.txt"
    try:
        requirements_lock_text = requirements_lock_path.read_text(encoding="utf-8")
    except OSError:
        return None

    # A compiled lock file can retain the source ``-r requirements.in`` line
    # while also carrying its own concrete pin.  Expanding includes in one
    # pass would encounter the looser source constraint first (for example
    # ``>=2.1``) and incorrectly report it instead of the lock's exact
    # ``==2.2`` pin.  Inspect direct entries first, then fall back to included
    # files only when the lock itself has no usable docxcompose constraint.
    direct_entries = _iter_requirement_entries(
        requirements_lock_text,
        include_base_dir=requirements_lock_path.parent,
        expand_includes=False,
    )
    for line in direct_entries:
        if not line:
            continue
        docxcompose_min_version = _extract_docxcompose_min_supported_version(line)
        if docxcompose_min_version is not None:
            return docxcompose_min_version

    expanded_entries = _iter_requirement_entries(
        requirements_lock_text,
        include_base_dir=requirements_lock_path.parent,
    )
    for line in expanded_entries:
        if not line:
            continue
        docxcompose_min_version = _extract_docxcompose_min_supported_version(line)
        if docxcompose_min_version is not None:
            return docxcompose_min_version
    return None



_DOCXCOMPOSE_MIN_VERSION_FROM_REQUIREMENTS = (
    _read_docxcompose_min_supported_version_from_requirements()
)
_DOCXCOMPOSE_MIN_VERSION_FROM_LOCK = None


_DOCXCOMPOSE_MIN_VERSION_FROM_LOCK = _read_docxcompose_min_supported_version_from_lock()


def _resolve_docxcompose_min_supported_version() -> tuple[str, str]:
    """返回 (`min_version`, `source`) 元组。

    source:
      - ``runtime``: 显式配置的 DOCXCOMPOSE_MIN_VERSION 生效
      - ``requirements_in``: 解析 requirements.in 成功
      - ``requirements_txt``: 回退到 requirements.txt 解析结果
      - ``unavailable``: 未能解析到可用版本约束
    """
    configured = (DOCXCOMPOSE_MIN_VERSION or "").strip()
    requirements_min_version = (
        _DOCXCOMPOSE_MIN_VERSION_FROM_REQUIREMENTS.strip()
        if isinstance(_DOCXCOMPOSE_MIN_VERSION_FROM_REQUIREMENTS, str)
        else None
    )
    lock_min_version = (
        _DOCXCOMPOSE_MIN_VERSION_FROM_LOCK.strip()
        if isinstance(_DOCXCOMPOSE_MIN_VERSION_FROM_LOCK, str)
        else None
    )

    if configured and requirements_min_version and configured != requirements_min_version:
        global _DOCXCOMPOSE_MIN_VERSION_WARNING_EMITTED
        if not _DOCXCOMPOSE_MIN_VERSION_WARNING_EMITTED:
            _DOCXCOMPOSE_MIN_VERSION_WARNING_EMITTED = True
            logger.warning(
                "docxcompose 最低支持版本配置与 requirements.in 不一致："
                "配置为 %s，requirements 为 %s。",
                configured,
                requirements_min_version,
            )
        return configured, "runtime"

    if configured:
        return configured, "runtime"
    if requirements_min_version:
        return requirements_min_version, "requirements_in"
    if lock_min_version:
        return lock_min_version, "requirements_txt"
    return "", "unavailable"


def get_docxcompose_min_supported_version() -> str:
    """返回可比对的最小支持版本（已清理空白）。"""
    return _resolve_docxcompose_min_supported_version()[0]


def get_docxcompose_min_supported_version_source_label(source: str | None) -> str:
    """返回文档拼接最小版本来源的中文标签。"""
    normalized_source = (source or "").strip().lower()
    if normalized_source in DOCXCOMPOSE_MIN_VERSION_SOURCE_LABELS:
        return DOCXCOMPOSE_MIN_VERSION_SOURCE_LABELS[normalized_source]
    if normalized_source == "unavailable":
        return "未配置"
    return "未知来源"


def is_docxcompose_version_supported(version: str | None) -> bool:
    """判断安装的 docxcompose 版本是否达到最低运行要求。"""
    version = version.strip() if isinstance(version, str) else version
    if not HAS_DOCXCOMPOSE or not version:
        return False
    min_version = get_docxcompose_min_supported_version()
    if not min_version:
        return False
    try:
        return _is_version_at_least(version, min_version)
    except Exception:  # pragma: no cover - 版本信息异常时保守降级
        return False


def get_docxcompose_support_state() -> dict[str, bool | str | None]:
    """返回 docxcompose 的可用状态与诊断信息。"""
    version = get_docxcompose_version()
    supported = is_docxcompose_version_supported(version)
    min_version, min_version_source = _resolve_docxcompose_min_supported_version()
    if not HAS_DOCXCOMPOSE:
        reason = "dependency_missing"
    elif not version:
        reason = "version_unknown"
    elif not supported:
        reason = "version_too_old"
    else:
        reason = None
    return {
        "supported": supported,
        "version": version,
        "reason": reason,
        "min_version": min_version,
        "min_version_source": min_version_source,
        "min_version_source_label": get_docxcompose_min_supported_version_source_label(
            min_version_source
        ),
    }


def get_docxcompose_version() -> str | None:
    """返回 docxcompose 的分发版版本（不可用则返回 None）。"""
    if not HAS_DOCXCOMPOSE:
        return None
    try:
        version = get_package_version("docxcompose")
        return version.strip() if isinstance(version, str) and version.strip() else None
    except PackageNotFoundError:
        return None
    except Exception:  # pragma: no cover - 环境意外问题兜底
        return None


def is_docxcompose_supported() -> bool:
    """当前环境下 docxcompose 是否可用于合并/拼接。"""
    return get_docxcompose_support_state()["supported"]


# ============================================================
# 应用配置
# ============================================================
def parse_cors_allowed_origins(value) -> tuple[str, ...]:
    if not isinstance(value, str):
        return ()

    origins = []
    for raw_origin in value.split(","):
        origin = raw_origin.strip()
        if not origin or origin in {"*", "null"}:
            continue
        if any(ch.isspace() or not ch.isprintable() for ch in origin):
            continue

        try:
            parsed = urlsplit(origin)
            hostname = parsed.hostname
            port = parsed.port
        except ValueError:
            continue

        if (
            parsed.scheme not in {"http", "https"}
            or not parsed.netloc
            or not hostname
            or parsed.username
            or parsed.password
            or parsed.query
            or parsed.fragment
            or parsed.path.rstrip("/")
            or (port is None and parsed.netloc.endswith(":"))
        ):
            continue

        normalized = f"{parsed.scheme.lower()}://{parsed.netloc.lower()}"
        if normalized not in origins:
            origins.append(normalized)
    return tuple(origins)


def parse_bounded_int(value, *, default: int, minimum: int, maximum: int) -> int:
    """解析受限整数配置；无效或越界值回退到安全默认值。"""
    if isinstance(value, bool):
        return default

    try:
        parsed = int(str(value).strip())
    except (TypeError, ValueError):
        return default

    if parsed < minimum or parsed > maximum:
        return default
    return parsed


app = Flask(__name__, static_folder="static", static_url_path="/static")

# 在 Serverless 环境中，通常只有 /tmp 目录有写入权限
UPLOAD_FOLDER = Path("/tmp") / "uploads"
OUTPUT_FOLDER = Path("/tmp") / "outputs"
UPLOAD_FOLDER.mkdir(exist_ok=True, parents=True)
OUTPUT_FOLDER.mkdir(exist_ok=True, parents=True)

app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB 上传限制
UPLOAD_STORAGE_RESERVATION_MULTIPLIER = 2
MAX_UPLOAD_STORAGE_RESERVATION_BYTES = (
    app.config["MAX_CONTENT_LENGTH"] * UPLOAD_STORAGE_RESERVATION_MULTIPLIER
)

ALLOWED_EXTENSIONS = {".docx"}
# Re-export the shared validation policy for compatibility with existing callers.
DocxValidationLimits = _docx_validation.DocxValidationLimits
INPUT_DOCX_LIMITS = _docx_validation.INPUT_DOCX_LIMITS
GENERATED_DOCX_LIMITS = _docx_validation.GENERATED_DOCX_LIMITS
MAX_INPUT_DOCX_ARCHIVE_BYTES = _docx_validation.MAX_INPUT_DOCX_ARCHIVE_BYTES
REQUIRED_DOCX_MEMBERS = _docx_validation.REQUIRED_DOCX_MEMBERS
MAX_DOCX_ARCHIVE_MEMBERS = _docx_validation.MAX_DOCX_ARCHIVE_MEMBERS
MAX_DOCX_MEMBER_UNCOMPRESSED_BYTES = _docx_validation.MAX_DOCX_MEMBER_UNCOMPRESSED_BYTES
MAX_DOCX_TOTAL_UNCOMPRESSED_BYTES = _docx_validation.MAX_DOCX_TOTAL_UNCOMPRESSED_BYTES
MAX_DOCX_XML_MEMBER_UNCOMPRESSED_BYTES = _docx_validation.MAX_DOCX_XML_MEMBER_UNCOMPRESSED_BYTES
MAX_DOCX_TOTAL_XML_UNCOMPRESSED_BYTES = _docx_validation.MAX_DOCX_TOTAL_XML_UNCOMPRESSED_BYTES
MAX_DOCX_CONTENT_TYPES_BYTES = _docx_validation.MAX_DOCX_CONTENT_TYPES_BYTES
MAX_DOCX_RELATIONSHIP_PART_BYTES = _docx_validation.MAX_DOCX_RELATIONSHIP_PART_BYTES
MAX_DOCX_TOTAL_RELATIONSHIP_BYTES = _docx_validation.MAX_DOCX_TOTAL_RELATIONSHIP_BYTES
MAX_DOCX_RELATIONSHIPS_PER_PART = _docx_validation.MAX_DOCX_RELATIONSHIPS_PER_PART
MAX_DOCX_TOTAL_RELATIONSHIPS = _docx_validation.MAX_DOCX_TOTAL_RELATIONSHIPS
MAX_DOCX_RELATIONSHIP_GRAPH_DEPTH = _docx_validation.MAX_DOCX_RELATIONSHIP_GRAPH_DEPTH
FORBIDDEN_XML_DECLARATION_MARKERS = _docx_validation.FORBIDDEN_XML_DECLARATION_MARKERS
MAX_DOCX_XML_ELEMENTS = _docx_validation.MAX_DOCX_XML_ELEMENTS
MAX_DOCX_XML_DEPTH = _docx_validation.MAX_DOCX_XML_DEPTH
MAX_DOCX_TOTAL_XML_ELEMENTS = _docx_validation.MAX_DOCX_TOTAL_XML_ELEMENTS
MAX_DOCX_TOTAL_PARAGRAPHS = _docx_validation.MAX_DOCX_TOTAL_PARAGRAPHS
MAX_DOCX_TOTAL_RUNS = _docx_validation.MAX_DOCX_TOTAL_RUNS
MAX_DOCX_NON_XML_COMPRESSION_RATIO = _docx_validation.MAX_DOCX_NON_XML_COMPRESSION_RATIO
MAX_DOCX_MEMBER_NAME_LENGTH = _docx_validation.MAX_DOCX_MEMBER_NAME_LENGTH
MAX_DOCX_TABLE_GRID_COLUMNS = _docx_validation.MAX_DOCX_TABLE_GRID_COLUMNS
MAX_DOCX_TABLE_LOGICAL_CELLS = _docx_validation.MAX_DOCX_TABLE_LOGICAL_CELLS
DOCX_MAIN_DOCUMENT_CONTENT_TYPE = _docx_validation.DOCX_MAIN_DOCUMENT_CONTENT_TYPE
MAX_GENERATED_DOCX_ARCHIVE_MEMBERS = _docx_validation.MAX_GENERATED_DOCX_ARCHIVE_MEMBERS
MAX_GENERATED_DOCX_MEMBER_UNCOMPRESSED_BYTES = _docx_validation.MAX_GENERATED_DOCX_MEMBER_UNCOMPRESSED_BYTES
MAX_GENERATED_DOCX_TOTAL_UNCOMPRESSED_BYTES = _docx_validation.MAX_GENERATED_DOCX_TOTAL_UNCOMPRESSED_BYTES
MAX_GENERATED_DOCX_CONTENT_TYPES_BYTES = _docx_validation.MAX_GENERATED_DOCX_CONTENT_TYPES_BYTES
MAX_GENERATED_DOCX_XML_MEMBER_UNCOMPRESSED_BYTES = _docx_validation.MAX_GENERATED_DOCX_XML_MEMBER_UNCOMPRESSED_BYTES
MAX_GENERATED_DOCX_TOTAL_XML_UNCOMPRESSED_BYTES = _docx_validation.MAX_GENERATED_DOCX_TOTAL_XML_UNCOMPRESSED_BYTES
MAX_GENERATED_DOCX_RELATIONSHIP_PART_BYTES = _docx_validation.MAX_GENERATED_DOCX_RELATIONSHIP_PART_BYTES
MAX_GENERATED_DOCX_TOTAL_RELATIONSHIP_BYTES = _docx_validation.MAX_GENERATED_DOCX_TOTAL_RELATIONSHIP_BYTES
MAX_GENERATED_DOCX_RELATIONSHIPS_PER_PART = _docx_validation.MAX_GENERATED_DOCX_RELATIONSHIPS_PER_PART
MAX_GENERATED_DOCX_TOTAL_RELATIONSHIPS = _docx_validation.MAX_GENERATED_DOCX_TOTAL_RELATIONSHIPS
MAX_GENERATED_DOCX_DOCUMENT_XML_ELEMENTS = _docx_validation.MAX_GENERATED_DOCX_DOCUMENT_XML_ELEMENTS
MAX_GENERATED_DOCX_TABLE_LOGICAL_CELLS = _docx_validation.MAX_GENERATED_DOCX_TABLE_LOGICAL_CELLS
MAX_GENERATED_DOCX_TOTAL_PARAGRAPHS = (
    _docx_validation.MAX_GENERATED_DOCX_TOTAL_PARAGRAPHS
)
MAX_GENERATED_DOCX_TOTAL_RUNS = _docx_validation.MAX_GENERATED_DOCX_TOTAL_RUNS
WORDPROCESSINGML_NAMESPACE = _docx_validation.WORDPROCESSINGML_NAMESPACE
PACKAGE_RELATIONSHIPS_NAMESPACE = _docx_validation.PACKAGE_RELATIONSHIPS_NAMESPACE
OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE = _docx_validation.OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE
STRICT_OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE = _docx_validation.STRICT_OFFICE_DOCUMENT_RELATIONSHIPS_NAMESPACE
DRAWINGML_NAMESPACE = _docx_validation.DRAWINGML_NAMESPACE
CORE_PROPERTIES_NAMESPACE = _docx_validation.CORE_PROPERTIES_NAMESPACE
EXTENDED_PROPERTIES_NAMESPACE = _docx_validation.EXTENDED_PROPERTIES_NAMESPACE
CUSTOM_PROPERTIES_NAMESPACE = _docx_validation.CUSTOM_PROPERTIES_NAMESPACE
WORD_DOCUMENT_TAG = _docx_validation.WORD_DOCUMENT_TAG
WORD_BACKGROUND_TAG = _docx_validation.WORD_BACKGROUND_TAG
WORD_BODY_TAG = _docx_validation.WORD_BODY_TAG
WORD_STYLES_TAG = _docx_validation.WORD_STYLES_TAG
WORD_NUMBERING_TAG = _docx_validation.WORD_NUMBERING_TAG
WORD_SETTINGS_TAG = _docx_validation.WORD_SETTINGS_TAG
WORD_FOOTNOTES_TAG = _docx_validation.WORD_FOOTNOTES_TAG
WORD_ENDNOTES_TAG = _docx_validation.WORD_ENDNOTES_TAG
WORD_HEADER_TAG = _docx_validation.WORD_HEADER_TAG
WORD_FOOTER_TAG = _docx_validation.WORD_FOOTER_TAG
WORD_COMMENTS_TAG = _docx_validation.WORD_COMMENTS_TAG
WORD_GLOSSARY_DOCUMENT_TAG = _docx_validation.WORD_GLOSSARY_DOCUMENT_TAG
WORD_FONTS_TAG = _docx_validation.WORD_FONTS_TAG
WORD_WEB_SETTINGS_TAG = _docx_validation.WORD_WEB_SETTINGS_TAG
PACKAGE_RELATIONSHIPS_TAG = _docx_validation.PACKAGE_RELATIONSHIPS_TAG
DRAWING_THEME_TAG = _docx_validation.DRAWING_THEME_TAG
CORE_PROPERTIES_TAG = _docx_validation.CORE_PROPERTIES_TAG
EXTENDED_PROPERTIES_TAG = _docx_validation.EXTENDED_PROPERTIES_TAG
CUSTOM_PROPERTIES_TAG = _docx_validation.CUSTOM_PROPERTIES_TAG
WORD_TABLE_TAG = _docx_validation.WORD_TABLE_TAG
WORD_TABLE_ROW_TAG = _docx_validation.WORD_TABLE_ROW_TAG
WORD_TABLE_CELL_TAG = _docx_validation.WORD_TABLE_CELL_TAG
WORD_TABLE_CELL_PROPERTIES_TAG = _docx_validation.WORD_TABLE_CELL_PROPERTIES_TAG
WORD_GRID_SPAN_TAG = _docx_validation.WORD_GRID_SPAN_TAG
WORD_VALUE_ATTRIBUTE = _docx_validation.WORD_VALUE_ATTRIBUTE
ASCII_PART_NAME_CASE_TABLE = _docx_validation.ASCII_PART_NAME_CASE_TABLE
PACKAGE_RELATIONSHIP_CONTENT_TYPE = _docx_validation.PACKAGE_RELATIONSHIP_CONTENT_TYPE
DOCX_XML_CONTENT_TYPE_ROOTS = _docx_validation.DOCX_XML_CONTENT_TYPE_ROOTS
OFFICE_DOCUMENT_RELATIONSHIP_TYPES = _docx_validation.OFFICE_DOCUMENT_RELATIONSHIP_TYPES
PACKAGE_ROOT_RELATIONSHIP_EXPECTED_CONTENT_TYPES = _docx_validation.PACKAGE_ROOT_RELATIONSHIP_EXPECTED_CONTENT_TYPES
DOCX_RELATIONSHIP_EXPECTED_CONTENT_TYPES = _docx_validation.DOCX_RELATIONSHIP_EXPECTED_CONTENT_TYPES
CASE_INSENSITIVE_PACKAGE_ROOT_RELATIONSHIP_CATEGORIES = _docx_validation.CASE_INSENSITIVE_PACKAGE_ROOT_RELATIONSHIP_CATEGORIES
CASE_INSENSITIVE_DOCX_RELATIONSHIP_CATEGORIES = _docx_validation.CASE_INSENSITIVE_DOCX_RELATIONSHIP_CATEGORIES
KNOWN_CONSUMED_DOCX_RELATIONSHIP_TYPES = _docx_validation.KNOWN_CONSUMED_DOCX_RELATIONSHIP_TYPES
KNOWN_CONSUMED_DOCX_RELATIONSHIP_TYPES_ASCII_LOWER = _docx_validation.KNOWN_CONSUMED_DOCX_RELATIONSHIP_TYPES_ASCII_LOWER
STRICT_RELATIONSHIP_PREFIX_ASCII_LOWER = _docx_validation.STRICT_RELATIONSHIP_PREFIX_ASCII_LOWER
TEMP_FILE_TTL_SECONDS = 6 * 60 * 60
JOB_TTL_SECONDS = TEMP_FILE_TTL_SECONDS
JOB_ID_BYTES = 16
JOB_ID_LENGTH = JOB_ID_BYTES * 2
JOB_ID_RE = re.compile(rf"^[0-9a-f]{{{JOB_ID_LENGTH}}}$")
RENDER_GIT_COMMIT_RE = re.compile(r"^[0-9a-f]{40}$")
CLIENT_JOB_ID_HEADER = "X-Job-ID"
GENERATED_OUTPUT_RE = re.compile(rf"^[0-9a-f]{{{JOB_ID_LENGTH}}}_output\.docx$")
SSE_EVENT_NAME_RE = re.compile(r"^[A-Za-z0-9_-]{1,32}$")
TERMINAL_JOB_STATUSES = {"done", "error"}
JOB_EXPIRED_MESSAGE = "任务已超时，请重新提交任务。"
USER_MESSAGE_PATH_TOKEN_RE = re.compile(r"(?<![\w:/])(?:[A-Za-z]:)?[\\/][^\s\"'<>|,;:]+")
JOB_ID_GENERATION_ATTEMPTS = 32
JOB_TEMP_SUFFIXES = ("input", "cover", "body", "first", "second")
MAX_FILENAME_TEXT_LENGTH = 512
MAX_LOG_FILENAME_LENGTH = 160
MAX_DOWNLOAD_STEM_LENGTH = 120
MAX_LAST_EVENT_ID_LENGTH = 32
MAX_SUMMARY_OUTLINE_ITEMS = 80
MAX_PREVIEW_OUTLINE_ITEMS = 8
MAX_OUTLINE_TEXT_LENGTH = 120
MAX_SUMMARY_TITLE_LENGTH = 240
MAX_PAGE_SETUP_TEXT_LENGTH = 120
MAX_PROGRESS_MESSAGE_LENGTH = 160
MAX_PROGRESS_DETAIL_LENGTH = 240
MAX_PROGRESS_EVENTS_PER_JOB = 80
MAX_PROGRESS_JOBS = 128
MIN_UNREAD_TERMINAL_RETENTION_SECONDS = 10 * 60
MIN_REQUESTED_TERMINAL_RETENTION_SECONDS = 60
MAX_OUTPUT_STORAGE_BYTES = 256 * 1024 * 1024
MAX_OUTPUT_STORAGE_FILES = 256
MAX_OUTPUT_FILE_BYTES = 64 * 1024 * 1024
# 标准预留不得小于单文件硬上限；保持别名可防止两者独立漂移。
OUTPUT_STORAGE_RESERVATION_BYTES = MAX_OUTPUT_FILE_BYTES
MIN_OUTPUT_FREE_BYTES = 64 * 1024 * 1024
MIN_OUTPUT_RETENTION_SECONDS = 30 * 60
OUTPUT_STORAGE_BUSY_MESSAGE = "服务器临时存储空间不足，请稍后重试。"
OUTPUT_TOO_LARGE_MESSAGE = "生成文件超过服务器单次输出上限，请精简文档后重试。"
DEFAULT_CONCURRENT_PROCESSING_JOBS = 2
MIN_CONCURRENT_PROCESSING_JOBS = 1
MAX_ALLOWED_CONCURRENT_PROCESSING_JOBS = 8
MAX_CONCURRENT_PROCESSING_JOBS = parse_bounded_int(
    os.environ.get("MAX_CONCURRENT_PROCESSING_JOBS"),
    default=DEFAULT_CONCURRENT_PROCESSING_JOBS,
    minimum=MIN_CONCURRENT_PROCESSING_JOBS,
    maximum=MAX_ALLOWED_CONCURRENT_PROCESSING_JOBS,
)
DEFAULT_CONCURRENT_SSE_CONNECTIONS = 4
MIN_CONCURRENT_SSE_CONNECTIONS = 1
MAX_ALLOWED_CONCURRENT_SSE_CONNECTIONS = 4
MAX_CONCURRENT_SSE_CONNECTIONS = parse_bounded_int(
    os.environ.get("MAX_CONCURRENT_SSE_CONNECTIONS"),
    default=DEFAULT_CONCURRENT_SSE_CONNECTIONS,
    minimum=MIN_CONCURRENT_SSE_CONNECTIONS,
    maximum=MAX_ALLOWED_CONCURRENT_SSE_CONNECTIONS,
)
SSE_RETRY_AFTER_SECONDS = 3
JOB_RESULT_RETRY_AFTER_SECONDS = 3
SSE_CAPACITY_MESSAGE = "实时进度连接较多，请稍后重试。"
SERVICE_DRAINING_MESSAGE = "服务正在更新，暂不接收新任务，请稍后重试。"
MAX_JOB_ERROR_MESSAGE_LENGTH = 200
MAX_TEXT_INPUT_BYTES = 2 * 1024 * 1024
MAX_TEXT_REQUEST_BYTES = MAX_TEXT_INPUT_BYTES * 3 + 64 * 1024
MAX_COVER_FIELD_LENGTH = 120
MAX_JSON_KEY_LENGTH = 120
MAX_JSON_COLLECTION_ITEMS = 200
MAX_JSON_STRING_LENGTH = 2000
TEXT_INPUT_TOO_LARGE_MESSAGE = "文本内容超过服务器单次处理上限（约 2MB），请拆分后再试。"
TEXT_INVALID_CHARACTERS_MESSAGE = "文本包含无法写入 Word 的控制字符，请删除后重试。"
TEXT_PARAGRAPH_LIMIT_MESSAGE = (
    f"文本段落数量超过服务器单次处理上限（{MAX_TEXT_PARAGRAPHS} 段），请删除多余空行或拆分后再试。"
)
TEXT_REQUEST_TOO_LARGE_MESSAGE = "请求内容超过文本排版接口上限（约 6MB），请移除无关字段或拆分文本后再试。"
CORS_ALLOWED_ORIGINS = parse_cors_allowed_origins(os.environ.get("CORS_ALLOWED_ORIGINS", ""))
PERMISSIONS_POLICY = "camera=(), microphone=(), geolocation=()"
CONTENT_SECURITY_POLICY = "; ".join(
    (
        "default-src 'self'",
        "script-src 'self'",
        "style-src 'self' 'unsafe-inline'",
        "font-src 'self'",
        "img-src 'self' data:",
        "connect-src 'self'",
        "object-src 'none'",
        "base-uri 'self'",
        "form-action 'self'",
        "frame-ancestors 'none'",
    )
)
PAGE_SETUP_TEXT_FIELDS = ("page_size", "header_text", "page_number_position")
PAGE_SETUP_FLOAT_FIELDS = (
    "page_width_cm",
    "page_height_cm",
    "header_distance_cm",
    "footer_distance_cm",
)
PAGE_MARGIN_FIELDS = ("top", "bottom", "left", "right")
RESULT_PAYLOAD_PRIORITY_KEYS = (
    "success",
    "download_url",
    "download_name",
    "original_name",
    "stats",
    "format_summary",
    "format_result",
    "preview",
)
OUTLINE_LEVELS = {
    "title",
    "h1",
    "h2",
    "h3",
    "section",
    "references",
    "english_abstract_heading",
    "abstract",
}
PROCESSING_STEP_LABELS = {
    1: "解析文档结构",
    2: "识别标题层级",
    3: "应用排版规则",
    4: "生成输出文档",
}
MAX_PROCESSING_STEP = max(PROCESSING_STEP_LABELS)
COVER_FORM_FIELDS = (
    "title",
    "cover_title",
    "course_title",
    "college",
    "teacher",
    "class_name",
    "student_name",
    "student_id",
    "school_name",
)

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)
PROGRESS_JOBS = {}
# 结果领取与输出淘汰必须先经保护锁串行化；其后的全局锁序只能是
# 任务表锁 → 单任务 condition → 输出存储锁，禁止反向获取。
OUTPUT_RESULT_GUARD_LOCK = threading.Lock()
PROGRESS_JOBS_LOCK = threading.Lock()
PROCESSING_JOB_SLOTS = threading.BoundedSemaphore(MAX_CONCURRENT_PROCESSING_JOBS)
SSE_CONNECTION_SLOTS = threading.BoundedSemaphore(MAX_CONCURRENT_SSE_CONNECTIONS)
OUTPUT_STORAGE_LOCK = threading.Lock()
ACTIVE_OUTPUT_RESERVATIONS = {}
ACTIVE_UPLOAD_RESERVATIONS = {}
ACTIVE_OUTPUT_DOWNLOADS = {}
ACTIVE_TEMP_FILES = set()
RECENT_OUTPUT_REQUESTS = {}
SHUTDOWN_EVENT = threading.Event()
MULTIPART_PROCESSING_REQUEST_ENDPOINTS = frozenset(
    {
        "api_format",
        "api_format_merge",
        "api_concat",
        "api_format_async",
        "api_format_merge_async",
        "api_concat_async",
    }
)
PROCESSING_REQUEST_ENDPOINTS = frozenset(
    MULTIPART_PROCESSING_REQUEST_ENDPOINTS
    | {"api_format_text", "api_format_text_async"}
)
ASYNC_PROCESSING_JOB_KINDS = {
    "api_format_async": "format",
    "api_format_text_async": "format_text",
    "api_format_merge_async": "format_merge",
    "api_concat_async": "concat",
}
REQUEST_PROCESSING_SLOT_ATTR = "_processing_job_slots"
REQUEST_UPLOAD_RESERVATION_ATTR = "_upload_storage_reservation_release"
REQUEST_MULTIPART_PARSED_ATTR = "_multipart_body_parsed"
REQUEST_BODY_CONSUMED_ATTR = "_request_body_consumed"
REQUEST_CLIENT_JOB_ID_ATTR = "_client_job_id"
REQUEST_REUSED_PROGRESS_JOB_ATTR = "_reused_progress_job"
REQUEST_FAILED_PROGRESS_JOB_ATTR = "_failed_progress_job"
REQUEST_ID_ATTR = "_request_id"
REQUEST_METHODS_WITH_BODIES = frozenset({"POST", "PUT", "PATCH", "DELETE"})


def configure_cors(flask_app, allowed_origins, cors_impl=CORS) -> str:
    if cors_impl is None:
        logger.warning("未检测到 Flask-Cors，已跳过 CORS 配置；同源部署不受影响。")
        return "disabled"

    if allowed_origins:
        # Flask-CORS 会把含正则元字符的来源字符串当作模式解析。环境变量属于
        # 不可信部署配置，因此逐项转义并锚定，确保 allowlist 始终是精确匹配。
        exact_origin_patterns = tuple(
            re.compile(rf"\A{re.escape(origin)}\Z")
            for origin in allowed_origins
            if isinstance(origin, str) and origin
        )
        if not exact_origin_patterns:
            logger.info("CORS 来源白名单为空，默认仅支持同源 API 访问。")
            return "same-origin"
        cors_impl(
            flask_app,
            resources={r"/api/*": {"origins": exact_origin_patterns}},
            # Retry-After drives bounded polling for queued jobs and SSE
            # reconnects.  It is not a CORS-safelisted response header, so
            # browser clients on an explicitly allowed origin must be able
            # to read it just like the request correlation ID.
            expose_headers=("X-Request-ID", "Retry-After"),
        )
        return "allowlist"

    logger.info("未配置 CORS_ALLOWED_ORIGINS，默认仅支持同源 API 访问。")
    return "same-origin"


configure_cors(app, CORS_ALLOWED_ORIGINS)


class JobProcessingError(Exception):
    """包装可直接返回给用户的处理失败信息。"""

    def __init__(self, message: str, status_code: int = 500):
        message = sanitize_user_facing_message(message, "处理失败")
        super().__init__(message)
        self.message = message
        self.status_code = status_code


def begin_graceful_shutdown() -> None:
    """进入排空状态：拒绝新处理任务，让已启动的非 daemon 任务自然完成。"""
    if SHUTDOWN_EVENT.is_set():
        return

    SHUTDOWN_EVENT.set()
    # 唤醒所有可能正在等待进度的 SSE 生成器，使长连接在排空时主动
    # 收尾，而不是一直占用 Gunicorn 线程直到客户端断开。
    with PROGRESS_JOBS_LOCK:
        for job in PROGRESS_JOBS.values():
            with job["condition"]:
                job["condition"].notify_all()
    logger.info("服务进入排空状态，停止接收新的文档处理任务。")


def acquire_processing_job_slot():
    """非阻塞地预留一个文档处理槽位，避免请求堆积耗尽线程与内存。"""
    if SHUTDOWN_EVENT.is_set():
        raise JobProcessingError(SERVICE_DRAINING_MESSAGE, 503)

    slots = PROCESSING_JOB_SLOTS
    if not slots.acquire(blocking=False):
        raise JobProcessingError("当前任务较多，请稍后重试。", 503)

    if SHUTDOWN_EVENT.is_set():
        slots.release()
        raise JobProcessingError(SERVICE_DRAINING_MESSAGE, 503)

    return slots


def claim_processing_job_slot():
    """取得请求入口预留的处理槽位；非请求调用则即时申请。"""
    if has_request_context():
        slots = getattr(g, REQUEST_PROCESSING_SLOT_ATTR, None)
        if slots is not None:
            setattr(g, REQUEST_PROCESSING_SLOT_ATTR, None)
            return slots

    return acquire_processing_job_slot()


def try_acquire_sse_connection_slot():
    """非阻塞预留 SSE 请求线程，并返回可重复调用的安全释放函数。"""
    slots = SSE_CONNECTION_SLOTS
    if not slots.acquire(blocking=False):
        return None

    release_lock = threading.Lock()
    released = False

    def release():
        nonlocal released
        with release_lock:
            if released:
                return
            released = True
        slots.release()

    return release


@contextmanager
def processing_job_slot():
    slots = claim_processing_job_slot()
    try:
        yield
    finally:
        slots.release()


def get_request_upload_reservation_bytes() -> int:
    """按请求体上界预留 spool 与命名副本短时并存所需的临时空间。"""
    content_length = request.content_length
    if content_length is None:
        return MAX_UPLOAD_STORAGE_RESERVATION_BYTES
    return min(
        max(0, int(content_length)) * UPLOAD_STORAGE_RESERVATION_MULTIPLIER,
        MAX_UPLOAD_STORAGE_RESERVATION_BYTES,
    )


def release_request_upload_reservation(*, close_parsed_files: bool = True) -> None:
    """关闭已解析上传并释放本请求的临时空间预留；可重复调用。"""
    if not has_request_context():
        return

    release = getattr(g, REQUEST_UPLOAD_RESERVATION_ATTR, None)
    if release is None:
        return
    setattr(g, REQUEST_UPLOAD_RESERVATION_ATTR, None)

    if close_parsed_files and getattr(g, REQUEST_MULTIPART_PARSED_ATTR, False):
        with suppress(Exception):
            request.close()
    release()


def mark_request_body_consumed() -> None:
    """记录当前请求体已完整解析，可安全复用底层 HTTP 连接。"""
    if has_request_context():
        setattr(g, REQUEST_BODY_CONSUMED_ATTR, True)


def request_may_have_unread_body() -> bool:
    """判断响应结束后是否仍可能留有未读请求体。"""
    # Gunicorn 只有 HTTP/1.x 会在 WSGI 响应后由同一 worker 排空 body；
    # HTTP/2 会先收完整个 stream 再调用应用，且连接可能承载其他并发 stream。
    if not str(request.environ.get("SERVER_PROTOCOL", "")).startswith("HTTP/1."):
        return False

    content_length = request.content_length
    return bool(
        request.method in REQUEST_METHODS_WITH_BODIES
        or (content_length is not None and content_length > 0)
        or request.headers.get("Transfer-Encoding")
    )


def shutdown_server_connection_best_effort(server_socket) -> None:
    """在响应写完后停用连接，跳过 Gunicorn 的慢速 body drain。"""
    with suppress(Exception):
        server_socket.shutdown(socket.SHUT_RDWR)


def create_server_connection_shutdown_callback(server_socket):
    """创建线程安全且可重复调用的单次断连回调。"""
    callback_lock = threading.Lock()
    callback_called = False

    def close_once():
        nonlocal callback_called
        with callback_lock:
            if callback_called:
                return
            callback_called = True
        shutdown_server_connection_best_effort(server_socket)

    return close_once


def force_close_unread_request_connection(response):
    """让未消费请求体的响应在发送完成后立即断连。

    WSGI 不保证应用返回的 hop-by-hop ``Connection`` 头会被服务器采用；
    Gunicorn 会直接丢弃该响应头。因此同时把底层 Gunicorn socket 的关闭动作
    注册到响应 close 回调，在完整响应写出后、worker 排空请求体前执行。
    """
    response.headers["Connection"] = "close"
    # send_file/send_from_directory 默认启用 direct_passthrough，Werkzeug 会直接
    # 返回底层 FileWrapper 而不套 ClosingIterator，call_on_close 因而不会执行。
    # 这里只关闭该旁路，仍保持流式迭代，不会把静态文件或下载内容整体读入内存。
    response.direct_passthrough = False
    server_socket = request.environ.get("gunicorn.socket")
    if server_socket is not None and callable(getattr(server_socket, "shutdown", None)):
        response.call_on_close(
            create_server_connection_shutdown_callback(server_socket)
        )
    return response


def reserve_request_processing_resources() -> None:
    """在请求体解析前取得处理槽与 multipart 临时空间预留。"""
    if getattr(g, REQUEST_PROCESSING_SLOT_ATTR, None) is None:
        setattr(g, REQUEST_PROCESSING_SLOT_ATTR, acquire_processing_job_slot())

    if (
        request.endpoint in MULTIPART_PROCESSING_REQUEST_ENDPOINTS
        and getattr(g, REQUEST_UPLOAD_RESERVATION_ATTR, None) is None
    ):
        # 预留 multipart spool 空间前先回收已过期上传副本；否则旧文件可能
        # 让磁盘余量检查提前返回 503，路由后续的常规清理来不及执行。
        cleanup_expired_files(UPLOAD_FOLDER)
        release = reserve_upload_storage(get_request_upload_reservation_bytes())
        setattr(g, REQUEST_UPLOAD_RESERVATION_ATTR, release)


def get_request_id() -> str | None:
    """返回服务器生成的请求标识，便于通过错误响应定位脱敏日志。"""
    if not has_request_context():
        return None
    request_id = getattr(g, REQUEST_ID_ATTR, None)
    if request_id is None:
        # 不信任客户端提供的标识，避免日志注入或重复值混淆独立请求。
        request_id = uuid.uuid4().hex
        setattr(g, REQUEST_ID_ATTR, request_id)
    return request_id


@app.before_request
def reserve_processing_job_slot_before_body_parsing():
    """在解析可能落盘的请求体前限流，避免繁忙请求放大临时存储占用。"""
    if request.method != "POST" or request.endpoint not in PROCESSING_REQUEST_ENDPOINTS:
        return

    if request.endpoint in ASYNC_PROCESSING_JOB_KINDS:
        client_job_id = request.headers.get(CLIENT_JOB_ID_HEADER)
        if client_job_id is not None:
            if not is_valid_job_id(client_job_id):
                raise JobProcessingError("任务恢复标识无效，请刷新页面后重试。", 400)
            setattr(g, REQUEST_CLIENT_JOB_ID_ATTR, client_job_id)

    if request.endpoint in MULTIPART_PROCESSING_REQUEST_ENDPOINTS:
        content_length = request.content_length
        if content_length is not None and content_length > app.config["MAX_CONTENT_LENGTH"]:
            raise RequestEntityTooLarge()

    # 幂等重试不应因为原任务正占用最后一个处理槽而被限流拒绝。
    # 先锁定当前已存在的同类型任务，路由随后会在解析请求体前直接复用它。
    if request.endpoint in ASYNC_PROCESSING_JOB_KINDS:
        expected_kind = ASYNC_PROCESSING_JOB_KINDS[request.endpoint]
        client_job_id = getattr(g, REQUEST_CLIENT_JOB_ID_ATTR, None)
        if client_job_id is not None:
            with PROGRESS_JOBS_LOCK:
                existing_job = PROGRESS_JOBS.get(client_job_id)
                if existing_job is not None:
                    with existing_job["condition"]:
                        existing_kind = existing_job.get("kind")
                    if existing_kind != expected_kind:
                        raise JobProcessingError(
                            "任务恢复标识已被其他任务使用，请刷新页面后重试。",
                            409,
                        )
                    setattr(g, REQUEST_REUSED_PROGRESS_JOB_ATTR, True)
                    return

    reserve_request_processing_resources()


@app.teardown_request
def release_unclaimed_processing_job_slot(_error=None):
    """校验提前返回或保存失败时，归还尚未移交给处理流程的槽位。"""
    release_request_upload_reservation(close_parsed_files=True)

    slots = getattr(g, REQUEST_PROCESSING_SLOT_ATTR, None)
    if slots is None:
        return

    setattr(g, REQUEST_PROCESSING_SLOT_ATTR, None)
    slots.release()


def sanitize_user_facing_message(value, fallback: str) -> str:
    message = value if isinstance(value, str) else ""
    message = USER_MESSAGE_PATH_TOKEN_RE.sub(lambda match: format_log_path(match.group(0)), message)
    message = "".join(ch if ch.isprintable() else " " for ch in message)
    message = re.sub(r"\s+", " ", message).strip()
    if len(message) > MAX_JOB_ERROR_MESSAGE_LENGTH:
        message = f"{message[:MAX_JOB_ERROR_MESSAGE_LENGTH - 1].rstrip()}…"
    return message or fallback


def normalize_filename_text(value) -> str:
    if not isinstance(value, str):
        return ""

    cleaned = "".join(ch for ch in value.replace("\\", "/") if ch.isprintable()).strip()
    return cleaned[:MAX_FILENAME_TEXT_LENGTH].rstrip()


def allowed_file(filename: str) -> bool:
    return Path(normalize_filename_text(filename)).suffix.lower() in ALLOWED_EXTENSIONS


# Public helper aliases keep the historical app module API while all logic lives
# in the side-effect-free shared validator.
contains_forbidden_xml_declaration = (
    _docx_validation.contains_forbidden_xml_declaration
)
stream_contains_forbidden_xml_declaration = (
    _docx_validation.stream_contains_forbidden_xml_declaration
)
declared_docx_member_content_type = (
    _docx_validation.declared_docx_member_content_type
)
docx_content_type_matches = _docx_validation.docx_content_type_matches
docx_member_content_type = _docx_validation.docx_member_content_type
canonical_docx_part_name = _docx_validation.canonical_docx_part_name
expected_docx_xml_root = _docx_validation.expected_docx_xml_root
is_bounded_xml_structure = _docx_validation.is_bounded_xml_structure
is_safe_docx_member_name = _docx_validation.is_safe_docx_member_name
parse_docx_content_types = _docx_validation.parse_docx_content_types
_relationship_source_part_name = (
    _docx_validation._relationship_source_part_name
)
validate_docx_relationships = _docx_validation.validate_docx_relationships
is_docx_xml_member = _docx_validation.is_docx_xml_member
parse_bounded_grid_span = _docx_validation.parse_bounded_grid_span
is_bounded_docx_document_xml = _docx_validation.is_bounded_docx_document_xml
is_valid_docx_stream = _docx_validation.is_valid_docx_stream
is_valid_docx_path = _docx_validation.is_valid_docx_path


def _current_input_docx_limits() -> DocxValidationLimits:
    """Build the upload profile from compatibility globals patched by callers/tests."""
    raw_archive_limit = app.config.get(
        "MAX_CONTENT_LENGTH",
        MAX_INPUT_DOCX_ARCHIVE_BYTES,
    )
    if (
        not isinstance(raw_archive_limit, int)
        or isinstance(raw_archive_limit, bool)
        or raw_archive_limit < 0
    ):
        raw_archive_limit = MAX_INPUT_DOCX_ARCHIVE_BYTES

    return replace(
        INPUT_DOCX_LIMITS,
        max_archive_bytes=raw_archive_limit,
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


def _current_generated_docx_limits() -> DocxValidationLimits:
    """Build the generated-output profile from compatibility globals."""
    return replace(
        GENERATED_DOCX_LIMITS,
        max_archive_members=MAX_GENERATED_DOCX_ARCHIVE_MEMBERS,
        max_member_uncompressed_bytes=(
            MAX_GENERATED_DOCX_MEMBER_UNCOMPRESSED_BYTES
        ),
        max_total_uncompressed_bytes=MAX_GENERATED_DOCX_TOTAL_UNCOMPRESSED_BYTES,
        max_xml_member_uncompressed_bytes=(
            MAX_GENERATED_DOCX_XML_MEMBER_UNCOMPRESSED_BYTES
        ),
        max_total_xml_uncompressed_bytes=(
            MAX_GENERATED_DOCX_TOTAL_XML_UNCOMPRESSED_BYTES
        ),
        max_content_types_bytes=MAX_GENERATED_DOCX_CONTENT_TYPES_BYTES,
        max_relationship_part_bytes=(
            MAX_GENERATED_DOCX_RELATIONSHIP_PART_BYTES
        ),
        max_total_relationship_bytes=(
            MAX_GENERATED_DOCX_TOTAL_RELATIONSHIP_BYTES
        ),
        max_relationships_per_part=(
            MAX_GENERATED_DOCX_RELATIONSHIPS_PER_PART
        ),
        max_total_relationships=MAX_GENERATED_DOCX_TOTAL_RELATIONSHIPS,
        max_relationship_graph_depth=MAX_DOCX_RELATIONSHIP_GRAPH_DEPTH,
        max_xml_elements=MAX_GENERATED_DOCX_DOCUMENT_XML_ELEMENTS,
        max_xml_depth=MAX_DOCX_XML_DEPTH,
        max_total_xml_elements=MAX_GENERATED_DOCX_DOCUMENT_XML_ELEMENTS,
        max_total_paragraphs=MAX_GENERATED_DOCX_TOTAL_PARAGRAPHS,
        max_total_runs=MAX_GENERATED_DOCX_TOTAL_RUNS,
        max_member_name_length=MAX_DOCX_MEMBER_NAME_LENGTH,
        max_table_grid_columns=MAX_DOCX_TABLE_GRID_COLUMNS,
        max_table_logical_cells=MAX_GENERATED_DOCX_TABLE_LOGICAL_CELLS,
    )


def is_valid_docx_upload(file_storage) -> bool:
    """Validate a Flask upload through the shared stream validator."""
    stream = getattr(file_storage, "stream", None)
    return is_valid_docx_stream(stream, limits=_current_input_docx_limits())


def is_valid_generated_docx(output_path: Path) -> bool:
    """Validate generated output through the shared expanded-budget profile."""
    return _docx_validation.is_valid_generated_docx(
        output_path,
        limits=_current_generated_docx_limits(),
    )


INVALID_DOCX_MESSAGE = "{label}不是有效的 .docx（可能已损坏或被改了后缀），请用 Word 重新导出后再试。"


def is_api_request() -> bool:
    return request.path == "/api" or request.path.startswith("/api/")


def async_creation_failure_response(message: str, status_code: int):
    """为携带恢复标识的异步创建请求附上任务是否已落表的信息。"""
    if not has_request_context() or request.endpoint not in ASYNC_PROCESSING_JOB_KINDS:
        return None
    if not getattr(g, REQUEST_CLIENT_JOB_ID_ATTR, None):
        return None

    job = getattr(g, REQUEST_FAILED_PROGRESS_JOB_ATTR, None)
    return json_async_job_creation_error(message, status_code, job)


def json_error(message, status_code: int):
    normalized_status = normalize_http_status_code(status_code)
    if normalized_status >= 500:
        response = async_creation_failure_response(message, status_code)
        if response is not None:
            return response

    response = jsonify({"success": False, "error": normalize_api_error_message(message)})
    # A draining worker is a transient 503. Expose the same bounded retry
    # contract used by SSE rejection and async result polling.
    if normalized_status == 503 and message == SERVICE_DRAINING_MESSAGE:
        response.headers["Retry-After"] = str(SSE_RETRY_AFTER_SECONDS)
    return response, normalized_status


def json_terminal_job_error(message, status_code: int):
    return (
        jsonify(
            {
                "success": False,
                "error": normalize_api_error_message(message),
                "status": "failed",
                "terminal": True,
            }
        ),
        normalize_http_status_code(status_code),
    )


def json_async_job_creation_error(message, status_code: int, job: dict | None = None):
    """标明失败响应是否已对应到可恢复任务，供前端安全处理 provisional 状态。"""
    payload = {
        "success": False,
        "error": normalize_api_error_message(message),
        "job_created": job is not None,
    }
    if job is not None:
        payload.update(
            {
                "job_id": job.get("id", ""),
                "status": "failed",
                "terminal": True,
                "events_url": job.get("events_url", ""),
                "result_url": job.get("result_url", ""),
            }
        )
    normalized_status = normalize_http_status_code(status_code)
    response = jsonify(payload)
    if normalized_status == 503 and message == SERVICE_DRAINING_MESSAGE:
        response.headers["Retry-After"] = str(SSE_RETRY_AFTER_SECONDS)
    return response, normalized_status


def get_http_exception_message(error: HTTPException) -> str:
    """将 Werkzeug HTTP 异常整理为稳定的 API 错误文案。"""
    status_code = error.code or 500
    if status_code == 400:
        return "请求内容无法解析，请检查提交的数据。"
    if status_code == 413:
        return "文件大小超过 50MB 限制，请压缩后重试。"
    if status_code >= 500:
        return "服务器内部错误，请稍后重试。"
    return error.description or "请求处理失败，请稍后重试。"


def is_real_directory_path(path: Path) -> bool:
    """只接受真实目录，避免存储根目录被符号链接替换。"""
    with suppress(OSError):
        return not path.is_symlink() and path.is_dir()
    return False


def is_writable_dir(path: Path) -> bool:
    """检查目录是否存在且当前进程可实际写入，用于健康诊断。"""
    path = Path(path)
    if not is_real_directory_path(path):
        return False

    probe = path / f".health-write-probe-{secrets.token_hex(8)}.tmp"
    flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL
    if hasattr(os, "O_NOFOLLOW"):
        flags |= os.O_NOFOLLOW

    fd = None
    try:
        fd = os.open(probe, flags, 0o600)
        os.write(fd, b"")
        return True
    except OSError:
        return False
    finally:
        if fd is not None:
            with suppress(OSError):
                os.close(fd)
        with suppress(OSError):
            probe.unlink()


def ensure_storage_ready(*folders: Path) -> None:
    """先回收再创建并校验临时目录，避免 ENOSPC 时写探针阻断清理。"""
    for folder in folders:
        folder = Path(folder)
        if is_real_directory_path(folder):
            cleanup_expired_files(folder)
            if folder == Path(OUTPUT_FOLDER):
                with suppress(JobProcessingError):
                    enforce_output_storage_budget()

        try:
            folder.mkdir(exist_ok=True, parents=True)
        except OSError as exc:
            raise JobProcessingError("服务器临时存储不可用，请稍后重试。", 503) from exc

        if not is_writable_dir(folder):
            raise JobProcessingError("服务器临时存储不可用，请稍后重试。", 503)


def get_progress_job_counts() -> dict:
    """返回不含任务标识或内容的任务表聚合快照，供健康检查使用。"""
    terminal_jobs = 0
    with PROGRESS_JOBS_LOCK:
        tracked_jobs = len(PROGRESS_JOBS)
        for job in PROGRESS_JOBS.values():
            with job["condition"]:
                if job.get("status") in TERMINAL_JOB_STATUSES:
                    terminal_jobs += 1

    return {
        "active": tracked_jobs - terminal_jobs,
        "terminal": terminal_jobs,
        "tracked": tracked_jobs,
        "capacity": MAX_PROGRESS_JOBS,
    }


def save_uploaded_file(file_storage, destination: Path) -> int:
    """独占保存并关闭上传流；返回大小且拒绝覆盖已占用的临时槽位。"""
    destination = Path(destination)
    flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL
    if hasattr(os, "O_NOFOLLOW"):
        flags |= os.O_NOFOLLOW

    # Open the parent directory first and create the file relative to that
    # descriptor.  Checking only the final file with O_NOFOLLOW still leaves
    # a window in which a replaced ``uploads`` directory could redirect the
    # path lookup to an attacker-controlled location.
    directory_fd = None
    fd = None
    created = False
    try:
        # 创建文件与登记活动租约共用存储锁，避免 TTL 清理恰好在两步
        # 之间看到一个尚未受保护的命名副本。
        with OUTPUT_STORAGE_LOCK:
            directory_flags = os.O_RDONLY
            for flag_name in ("O_CLOEXEC", "O_DIRECTORY", "O_NOFOLLOW"):
                directory_flags |= getattr(os, flag_name, 0)
            directory_fd = os.open(destination.parent, directory_flags)
            if not stat.S_ISDIR(os.fstat(directory_fd).st_mode):
                raise OSError("upload destination parent is not a directory")
            if os.supports_dir_fd and os.open in os.supports_dir_fd:
                fd = os.open(destination.name, flags, 0o600, dir_fd=directory_fd)
            else:  # pragma: no cover - legacy platforms without openat support
                fd = os.open(destination, flags, 0o600)
            created = True
            ACTIVE_TEMP_FILES.add(destination)
        if directory_fd is not None:
            os.close(directory_fd)
            directory_fd = None
        with os.fdopen(fd, "wb") as target:
            fd = None
            file_storage.save(target)

        if not is_regular_file_path(destination):
            raise OSError("uploaded file was not saved as a regular file")

        return destination.lstat().st_size
    except OSError as exc:
        if created:
            cleanup_path(destination)
        raise JobProcessingError("服务器临时存储不可用，请稍后重试。", 503) from exc
    except BaseException:
        # 上传流实现也可能抛出非 OSError（如请求中断或
        # 自定义存储后端异常）。保留原异常语义，但不留下
        # 部分文件或永久占用的清理租约。
        if created:
            cleanup_path(destination)
        raise
    finally:
        if fd is not None:
            with suppress(OSError):
                os.close(fd)
        if directory_fd is not None:
            with suppress(OSError):
                os.close(directory_fd)
        # 命名副本已经独立，尽早释放 Werkzeug 的 multipart 临时文件。
        with suppress(Exception):
            file_storage.close()


def get_render_git_commit() -> str:
    """返回可公开用于部署探活的 Render commit；拒绝异常环境值。"""
    value = os.environ.get("RENDER_GIT_COMMIT", "")
    if not isinstance(value, str):
        return ""
    normalized = value.strip().lower()
    return normalized if RENDER_GIT_COMMIT_RE.fullmatch(normalized) else ""


def get_health_payload() -> dict:
    for folder in (UPLOAD_FOLDER, OUTPUT_FOLDER):
        with suppress(OSError):
            folder.mkdir(exist_ok=True, parents=True)

    # 保活流量也承担轻量 TTL 维护，避免长期无业务请求时崩溃遗留文件
    # 或任务状态持续占用资源；活动产物、worker 与下载租约仍由
    # 清理函数内部的锁和租约保护。
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_jobs()

    storage = {
        "upload_dir": format_log_path(UPLOAD_FOLDER),
        "output_dir": format_log_path(OUTPUT_FOLDER),
        "upload_ready": is_real_directory_path(UPLOAD_FOLDER),
        "output_ready": is_real_directory_path(OUTPUT_FOLDER),
        "upload_writable": is_writable_dir(UPLOAD_FOLDER),
        "output_writable": is_writable_dir(OUTPUT_FOLDER),
    }
    try:
        output_storage = get_output_storage_snapshot()
        # “当前未超限”不等于“仍能接单”；部署探活必须确认至少还能
        # 为一份标准产物预留空间，避免所有新任务都会 503 时仍报 ok。
        storage["output_budget_ready"] = bool(
            output_storage.get("within_budget") is True
            and output_storage.get("can_reserve_output") is True
        )
    except JobProcessingError:
        output_storage = {
            "files": 0,
            "bytes": 0,
            "effective_files": 0,
            "effective_bytes": 0,
            "active_reservations": 0,
            "active_upload_reservations": 0,
            "active_downloads": 0,
            "upload_reserved_bytes": 0,
            "within_budget": False,
            "can_reserve_output": False,
        }
        storage["output_budget_ready"] = False
    storage["storage_ready"] = all(
        storage[key]
        for key in (
            "upload_ready",
            "output_ready",
            "upload_writable",
            "output_writable",
            "output_budget_ready",
        )
    )

    draining = SHUTDOWN_EVENT.is_set()
    service_ready = storage["storage_ready"] and not draining

    docxcompose_state = get_docxcompose_support_state()

    return {
        "success": service_ready,
        "status": "draining" if draining else ("ok" if service_ready else "degraded"),
        "timestamp": int(time.time()),
        "draining": draining,
        "deployment": {
            "commit": get_render_git_commit(),
        },
        "storage": storage,
        "output_storage": output_storage,
        "jobs": get_progress_job_counts(),
            "features": {
                "cover_merge": docxcompose_state["supported"],
                "concat": docxcompose_state["supported"],
                "docxcompose_version": docxcompose_state["version"],
                "docxcompose_support_reason": docxcompose_state["reason"],
                "docxcompose_min_version": docxcompose_state["min_version"],
                "docxcompose_min_version_source": docxcompose_state[
                    "min_version_source"
                ],
                "docxcompose_min_version_source_label": docxcompose_state[
                    "min_version_source_label"
                ],
            },
        "limits": {
            "max_concurrent_processing_jobs": MAX_CONCURRENT_PROCESSING_JOBS,
            "max_output_storage_bytes": MAX_OUTPUT_STORAGE_BYTES,
            "max_output_storage_files": MAX_OUTPUT_STORAGE_FILES,
            "max_output_file_bytes": MAX_OUTPUT_FILE_BYTES,
            "max_upload_storage_reservation_bytes": MAX_UPLOAD_STORAGE_RESERVATION_BYTES,
            "min_output_retention_seconds": MIN_OUTPUT_RETENTION_SECONDS,
        },
    }


def get_display_name(filename: str) -> str:
    """提取用于展示/下载的原始文件名，保留中文等 Unicode 字符。"""
    normalized = normalize_filename_text(filename)
    basename = Path(normalized).name
    stem = Path(basename).stem.strip()
    stem = "".join(ch for ch in stem if ch.isprintable()).strip()
    stem = stem[:MAX_DOWNLOAD_STEM_LENGTH].rstrip()
    return stem or "document"


def get_log_filename(filename) -> str:
    """整理日志里的上传文件名，避免路径片段或控制字符污染日志输出。"""
    normalized = normalize_filename_text(filename)
    basename = Path(normalized).name.strip()
    basename = basename[:MAX_LOG_FILENAME_LENGTH].rstrip()
    return basename or "unnamed"


def cleanup_expired_files(*folders: Path, max_age_seconds: int = TEMP_FILE_TTL_SECONDS) -> None:
    """清理过期的临时 docx 文件与空目录，避免 /tmp 持续膨胀。"""
    now = time.time()
    cutoff = now - max_age_seconds
    formatter_temp_paths = get_active_temp_paths()

    for folder in folders:
        folder = Path(folder)
        # 核心格式化器用同目录隐藏暂存文件实现原子发布；进程被强杀时
        # 可能留下该固定前缀的临时文件，TTL 到期后可安全回收。
        with suppress(OSError):
            if is_real_directory_path(folder):
                for staging_path in folder.glob(".docx-output-*.tmp"):
                    with suppress(OSError):
                        if staging_path not in formatter_temp_paths:
                            # 输出存储锁保护 Web 层租约；核心格式化
                            # registry 再在内层把“复核 + 删除”原子化。
                            with OUTPUT_STORAGE_LOCK:
                                if staging_path not in ACTIVE_TEMP_FILES:
                                    remove_expired_inactive_temp_path(
                                        staging_path,
                                        cutoff,
                                    )
        if folder == Path(OUTPUT_FOLDER):
            cleanup_expired_output_files(folder, cutoff=cutoff, now=now)
            continue

        with suppress(OSError):
            if not is_real_directory_path(folder):
                continue

            with OUTPUT_STORAGE_LOCK:
                active_temp_files = {
                    path
                    for path in ACTIVE_TEMP_FILES
                    if path.parent == folder
                }
            active_temp_files.update(
                path
                for path in formatter_temp_paths
                if path.parent == folder
            )
            for path in folder.glob("*.docx"):
                with suppress(OSError):
                    if path in active_temp_files:
                        continue
                    # save_uploaded_file() 也在同一把锁内创建文件并
                    # 登记租约，因此此处持锁复核后可安全删除。
                    with OUTPUT_STORAGE_LOCK:
                        if path not in ACTIVE_TEMP_FILES:
                            remove_expired_inactive_temp_path(path, cutoff)


def get_output_job_states() -> dict:
    """快照输出文件对应的异步任务状态；调用方不得同时持有输出存储锁。"""
    states = {}
    with PROGRESS_JOBS_LOCK:
        for job_id, job in PROGRESS_JOBS.items():
            with job["condition"]:
                states[f"{job_id}_output.docx"] = {
                    "status": job.get("status"),
                    "result_requested_at": job.get("result_requested_at"),
                }
    return states


@contextmanager
def lock_output_eviction_state():
    """原子衔接任务状态快照与输出淘汰，避免结果领取插入两者之间。"""
    with OUTPUT_RESULT_GUARD_LOCK:
        job_states = get_output_job_states()
        with OUTPUT_STORAGE_LOCK:
            yield job_states


def get_output_eviction_priority(job_state, now: float, recent_requested_at=None):
    """返回安全淘汰优先级；None 表示任务仍活动或结果刚被领取。"""
    if job_state and job_state.get("status") not in TERMINAL_JOB_STATUSES:
        return None

    requested_at = job_state.get("result_requested_at") if job_state else None
    if isinstance(recent_requested_at, (int, float)) and math.isfinite(recent_requested_at):
        if (
            not isinstance(requested_at, (int, float))
            or not math.isfinite(requested_at)
            or recent_requested_at > requested_at
        ):
            requested_at = recent_requested_at
    if isinstance(requested_at, (int, float)) and math.isfinite(requested_at):
        if requested_at > now - MIN_REQUESTED_TERMINAL_RETENTION_SECONDS:
            return None
        return 0
    if job_state is None:
        return 1
    return 2


def prune_recent_output_requests_locked(now: float) -> None:
    """清理过期下载保护元数据；调用方必须持有输出存储锁。"""
    cutoff = now - MIN_REQUESTED_TERMINAL_RETENTION_SECONDS
    for path, timestamp in list(RECENT_OUTPUT_REQUESTS.items()):
        if (
            not isinstance(timestamp, (int, float))
            or isinstance(timestamp, bool)
            or not math.isfinite(timestamp)
            or timestamp <= cutoff
        ):
            RECENT_OUTPUT_REQUESTS.pop(path, None)


def get_recent_output_request_locked(path: Path, now: float):
    """读取仍在保护窗内的最近结果请求；调用方必须持有输出存储锁。"""
    requested_at = RECENT_OUTPUT_REQUESTS.get(path)
    if (
        isinstance(requested_at, (int, float))
        and not isinstance(requested_at, bool)
        and math.isfinite(requested_at)
        and requested_at > now - MIN_REQUESTED_TERMINAL_RETENTION_SECONDS
    ):
        return requested_at
    RECENT_OUTPUT_REQUESTS.pop(path, None)
    return None


def record_recent_output_request_locked(path: Path, requested_at: float) -> None:
    """刷新下载重试保护窗；调用方必须持有输出存储锁。"""
    prune_recent_output_requests_locked(requested_at)
    RECENT_OUTPUT_REQUESTS[path] = requested_at


def cleanup_expired_output_files(folder: Path, *, cutoff: float, now: float) -> None:
    """按输出锁与任务状态安全清理，保留活动、刚领取及未知来源文件。"""
    with lock_output_eviction_state() as job_states:
        if not is_real_directory_path(folder):
            return
        prune_recent_output_requests_locked(now)
        try:
            records = scan_output_files_locked(folder)
        except JobProcessingError:
            return

        active_paths = {
            path
            for path in ACTIVE_OUTPUT_RESERVATIONS
            if path.parent == folder
        }
        active_paths.update(
            path
            for path, count in ACTIVE_OUTPUT_DOWNLOADS.items()
            if path.parent == folder and count > 0
        )
        for path, record in records.items():
            if (
                path in active_paths
                or record["mtime"] >= cutoff
                or not is_generated_output(path.name)
                or record.get("nlink", 1) != 1
            ):
                continue
            eviction_priority = get_output_eviction_priority(
                job_states.get(path.name),
                now,
                get_recent_output_request_locked(path, now),
            )
            if eviction_priority is None:
                continue
            with suppress(OSError):
                path.unlink()


def scan_output_files_locked(folder: Path) -> dict:
    """在输出存储锁内读取普通 .docx 文件元数据，不跟随符号链接。"""
    try:
        paths = list(folder.glob("*.docx"))
    except OSError as exc:
        raise JobProcessingError(OUTPUT_STORAGE_BUSY_MESSAGE, 503) from exc

    records = {}
    for path in paths:
        try:
            metadata = path.lstat()
        except OSError:
            continue
        if not stat.S_ISREG(metadata.st_mode):
            continue
        records[path] = {
            "size": max(0, metadata.st_size),
            "mtime": metadata.st_mtime,
            # Do not evict hard-linked inodes: another directory entry may
            # reference the same content and unlinking this name would still
            # mutate ownership outside the output directory.
            "nlink": getattr(metadata, "st_nlink", 1),
        }
    return records


def get_output_storage_snapshot() -> dict:
    """返回不含文件名的输出目录聚合用量，不执行淘汰。"""
    folder = Path(OUTPUT_FOLDER)
    with OUTPUT_STORAGE_LOCK:
        if not is_real_directory_path(folder):
            raise JobProcessingError(OUTPUT_STORAGE_BUSY_MESSAGE, 503)
        prune_recent_output_requests_locked(time.time())
        records = scan_output_files_locked(folder)
        reservations = {
            path: max(0, int(reserved_bytes))
            for path, reserved_bytes in ACTIVE_OUTPUT_RESERVATIONS.items()
            if path.parent == folder
        }
        upload_reserved_bytes = sum(
            max(0, int(reserved_bytes))
            for reserved_bytes in ACTIVE_UPLOAD_RESERVATIONS.values()
        )
        total_bytes = sum(record["size"] for record in records.values())
        reservation_remaining = sum(
            max(0, reserved_bytes - records.get(path, {}).get("size", 0))
            for path, reserved_bytes in reservations.items()
        )
        effective_files = len(set(records) | set(reservations))
        effective_bytes = total_bytes + reservation_remaining
        try:
            free_bytes = shutil.disk_usage(folder).free
        except OSError as exc:
            raise JobProcessingError(OUTPUT_STORAGE_BUSY_MESSAGE, 503) from exc
        return {
            "files": len(records),
            "bytes": total_bytes,
            "effective_files": effective_files,
            "effective_bytes": effective_bytes,
            "active_reservations": len(reservations),
            "active_upload_reservations": len(ACTIVE_UPLOAD_RESERVATIONS),
            "active_downloads": sum(
                max(0, int(count))
                for path, count in ACTIVE_OUTPUT_DOWNLOADS.items()
                if path.parent == folder
            ),
            "upload_reserved_bytes": upload_reserved_bytes,
            "within_budget": (
                effective_files <= MAX_OUTPUT_STORAGE_FILES
                and effective_bytes <= MAX_OUTPUT_STORAGE_BYTES
            ),
            "can_reserve_output": (
                effective_files + 1 <= MAX_OUTPUT_STORAGE_FILES
                and effective_bytes + OUTPUT_STORAGE_RESERVATION_BYTES <= MAX_OUTPUT_STORAGE_BYTES
                and free_bytes
                >= (
                    MIN_OUTPUT_FREE_BYTES
                    + reservation_remaining
                    + upload_reserved_bytes
                    + OUTPUT_STORAGE_RESERVATION_BYTES
                )
            ),
        }


def enforce_output_storage_budget() -> dict:
    """回收超过保留期的旧产物，并强制输出目录满足数量、字节与剩余空间预算。"""
    folder = Path(OUTPUT_FOLDER)
    now = time.time()

    with lock_output_eviction_state() as job_states:
        if not is_real_directory_path(folder):
            raise JobProcessingError(OUTPUT_STORAGE_BUSY_MESSAGE, 503)

        prune_recent_output_requests_locked(now)
        records = scan_output_files_locked(folder)
        reservations = {
            path: max(0, int(reserved_bytes))
            for path, reserved_bytes in ACTIVE_OUTPUT_RESERVATIONS.items()
            if path.parent == folder
        }
        upload_reserved_bytes = sum(
            max(0, int(reserved_bytes))
            for reserved_bytes in ACTIVE_UPLOAD_RESERVATIONS.values()
        )

        def calculate_usage():
            total_bytes = sum(record["size"] for record in records.values())
            reservation_remaining = sum(
                max(0, reserved_bytes - records.get(path, {}).get("size", 0))
                for path, reserved_bytes in reservations.items()
            )
            effective_files = len(set(records) | set(reservations))
            try:
                free_bytes = shutil.disk_usage(folder).free
            except OSError as exc:
                raise JobProcessingError(OUTPUT_STORAGE_BUSY_MESSAGE, 503) from exc
            return {
                "files": len(records),
                "bytes": total_bytes,
                "reserved_bytes": reservation_remaining,
                "upload_reserved_bytes": upload_reserved_bytes,
                "effective_files": effective_files,
                "effective_bytes": total_bytes + reservation_remaining,
                "free_bytes": free_bytes,
            }

        def exceeds_budget(usage: dict) -> bool:
            return (
                usage["effective_files"] > MAX_OUTPUT_STORAGE_FILES
                or usage["effective_bytes"] > MAX_OUTPUT_STORAGE_BYTES
                or usage["free_bytes"]
                < (
                    MIN_OUTPUT_FREE_BYTES
                    + usage["reserved_bytes"]
                    + usage["upload_reserved_bytes"]
                )
            )

        usage = calculate_usage()
        if not exceeds_budget(usage):
            return usage

        active_paths = set(reservations)
        active_paths.update(
            path
            for path, count in ACTIVE_OUTPUT_DOWNLOADS.items()
            if path.parent == folder and count > 0
        )
        retention_cutoff = now - MIN_OUTPUT_RETENTION_SECONDS
        candidates = []
        for path, record in records.items():
            if (
                path in active_paths
                or record["mtime"] > retention_cutoff
                or record.get("nlink", 1) != 1
            ):
                continue
            if not is_generated_output(path.name):
                continue

            priority = get_output_eviction_priority(
                job_states.get(path.name),
                now,
                get_recent_output_request_locked(path, now),
            )
            if priority is None:
                continue
            candidates.append((priority, record["mtime"], path.name, path))

        for _priority, _mtime, _name, path in sorted(candidates):
            try:
                path.unlink()
            except FileNotFoundError:
                pass
            except OSError:
                continue
            records.pop(path, None)
            usage = calculate_usage()
            if not exceeds_budget(usage):
                return usage

        usage = calculate_usage()
        if exceeds_budget(usage):
            raise JobProcessingError(OUTPUT_STORAGE_BUSY_MESSAGE, 503)
        return usage


def reserve_output_storage(output_path: Path, reserved_bytes: int | None = None):
    """为一个正在生成的输出预留预算，返回可重复调用的释放函数。"""
    output_path = Path(output_path)
    if reserved_bytes is None:
        reserved_bytes = OUTPUT_STORAGE_RESERVATION_BYTES
    if output_path.parent != Path(OUTPUT_FOLDER) or not is_generated_output(output_path.name):
        raise JobProcessingError(OUTPUT_STORAGE_BUSY_MESSAGE, 503)

    with OUTPUT_STORAGE_LOCK:
        if output_path in ACTIVE_OUTPUT_RESERVATIONS:
            raise JobProcessingError(OUTPUT_STORAGE_BUSY_MESSAGE, 503)
        ACTIVE_OUTPUT_RESERVATIONS[output_path] = max(0, int(reserved_bytes))

    try:
        enforce_output_storage_budget()
    except Exception:
        with OUTPUT_STORAGE_LOCK:
            ACTIVE_OUTPUT_RESERVATIONS.pop(output_path, None)
        raise

    release_lock = threading.Lock()
    released = False

    def release():
        nonlocal released
        with release_lock:
            if released:
                return
            released = True
        with OUTPUT_STORAGE_LOCK:
            ACTIVE_OUTPUT_RESERVATIONS.pop(output_path, None)

    return release


def reserve_upload_storage(reserved_bytes: int):
    """预留 multipart spool 与命名副本的增长空间，返回幂等释放函数。"""
    reservation_id = object()
    with OUTPUT_STORAGE_LOCK:
        ACTIVE_UPLOAD_RESERVATIONS[reservation_id] = max(0, int(reserved_bytes))

    try:
        enforce_output_storage_budget()
    except Exception:
        with OUTPUT_STORAGE_LOCK:
            ACTIVE_UPLOAD_RESERVATIONS.pop(reservation_id, None)
        raise

    release_lock = threading.Lock()
    released = False

    def release():
        nonlocal released
        with release_lock:
            if released:
                return
            released = True
        with OUTPUT_STORAGE_LOCK:
            ACTIVE_UPLOAD_RESERVATIONS.pop(reservation_id, None)

    return release


def complete_output_storage_reservation(output_path: Path, actual_size: int) -> None:
    """产物写完后把预留量收敛为真实大小，随后可执行最终容量检查。"""
    output_path = Path(output_path)
    with OUTPUT_STORAGE_LOCK:
        if output_path in ACTIVE_OUTPUT_RESERVATIONS:
            ACTIVE_OUTPUT_RESERVATIONS[output_path] = max(0, int(actual_size))


def acquire_output_download_lease(output_path: Path):
    """安全打开并租用一个输出，返回（幂等释放函数、文件流、文件大小）。"""
    output_path = Path(output_path)
    folder = Path(OUTPUT_FOLDER)
    now = time.time()
    download_file = None
    with OUTPUT_STORAGE_LOCK:
        if (
            output_path.parent != folder
            or not is_generated_output(output_path.name)
            or output_path in ACTIVE_OUTPUT_RESERVATIONS
        ):
            return None

        directory_fd = None
        fd = None
        try:
            directory_flags = os.O_RDONLY
            for flag_name in ("O_CLOEXEC", "O_DIRECTORY", "O_NOFOLLOW"):
                directory_flags |= getattr(os, flag_name, 0)
            directory_fd = os.open(folder, directory_flags)
            if not stat.S_ISDIR(os.fstat(directory_fd).st_mode):
                return None

            path_metadata = os.stat(
                output_path.name,
                dir_fd=directory_fd,
                follow_symlinks=False,
            )
            if not stat.S_ISREG(path_metadata.st_mode):
                return None

            flags = os.O_RDONLY
            for flag_name in ("O_CLOEXEC", "O_NOFOLLOW", "O_NONBLOCK", "O_BINARY"):
                flags |= getattr(os, flag_name, 0)
            fd = os.open(output_path.name, flags, dir_fd=directory_fd)
            opened_metadata = os.fstat(fd)
            if (
                not stat.S_ISREG(opened_metadata.st_mode)
                or (opened_metadata.st_dev, opened_metadata.st_ino)
                != (path_metadata.st_dev, path_metadata.st_ino)
                # Generated outputs must own their inode. A hard link could
                # otherwise make cleanup of an output affect an unrelated
                # file that happens to be linkable in this folder.
                or getattr(opened_metadata, "st_nlink", 1) != 1
                or opened_metadata.st_size <= 0
                or opened_metadata.st_size > MAX_OUTPUT_FILE_BYTES
            ):
                return None

            download_file = os.fdopen(fd, "rb")
            fd = None
        except OSError as exc:
            if exc.errno in {errno.ENOENT, errno.ENOTDIR, errno.ELOOP, errno.EISDIR}:
                return None
            raise JobProcessingError(OUTPUT_STORAGE_BUSY_MESSAGE, 503) from exc
        except (ValueError, TypeError, NotImplementedError) as exc:
            raise JobProcessingError(OUTPUT_STORAGE_BUSY_MESSAGE, 503) from exc
        finally:
            if fd is not None:
                with suppress(OSError):
                    os.close(fd)
            if directory_fd is not None:
                with suppress(OSError):
                    os.close(directory_fd)

        ACTIVE_OUTPUT_DOWNLOADS[output_path] = ACTIVE_OUTPUT_DOWNLOADS.get(output_path, 0) + 1
        record_recent_output_request_locked(output_path, now)

    release_lock = threading.Lock()
    released = False

    def release():
        nonlocal released
        with release_lock:
            if released:
                return
            released = True

        with suppress(Exception):
            download_file.close()
        with OUTPUT_STORAGE_LOCK:
            remaining = ACTIVE_OUTPUT_DOWNLOADS.get(output_path, 0) - 1
            if remaining > 0:
                ACTIVE_OUTPUT_DOWNLOADS[output_path] = remaining
            else:
                ACTIVE_OUTPUT_DOWNLOADS.pop(output_path, None)
            record_recent_output_request_locked(output_path, time.time())

    return release, download_file, opened_metadata.st_size


def close_output_download_response(response, release_download) -> None:
    """异常路径关闭响应且无条件释放租约，避免 close 异常掩盖原错误。"""
    if response is not None:
        with suppress(Exception):
            response.close()
    release_download()


class OutputDownloadLeaseIterable:
    """关闭底层下载迭代器后必定释放租约，包括 HEAD 与中途断开。"""

    def __init__(self, iterable, release_download):
        self._iterable = iterable
        self._iterator = iter(iterable)
        self._release_download = release_download
        self._close_lock = threading.Lock()
        self._closed = False

    def __iter__(self):
        return self

    def __next__(self):
        try:
            return next(self._iterator)
        except BaseException:
            self.close()
            raise

    def close(self) -> None:
        with self._close_lock:
            if self._closed:
                return
            self._closed = True

        try:
            close = getattr(self._iterable, "close", None)
            if close is not None:
                close()
        except Exception:
            pass
        finally:
            self._release_download()


def get_progress_job_activity_time_locked(job: dict) -> float:
    """返回任务最近有效活动时间；调用方必须持有任务 condition。"""
    created_at = job.get("created_at", 0)
    updated_at = job.get("updated_at", created_at)
    if not isinstance(updated_at, (int, float)) or isinstance(updated_at, bool) or not math.isfinite(updated_at):
        updated_at = created_at if isinstance(created_at, (int, float)) else 0

    if job.get("status") not in TERMINAL_JOB_STATUSES:
        return float(updated_at)

    requested_at = job.get("result_requested_at")
    if (
        isinstance(requested_at, (int, float))
        and not isinstance(requested_at, bool)
        and math.isfinite(requested_at)
    ):
        return max(float(updated_at), float(requested_at))
    return float(updated_at)


def _expire_progress_job_if_stale_locked(job: dict, cutoff: float) -> bool:
    if get_progress_job_activity_time_locked(job) >= cutoff:
        return False

    if job.get("status") in TERMINAL_JOB_STATUSES:
        return True

    job["status"] = "error"
    job["error"] = {
        "message": JOB_EXPIRED_MESSAGE,
        "status_code": 410,
    }
    _append_job_event_locked(
        job,
        "failed",
        {
            "job_id": job.get("id", ""),
            "message": JOB_EXPIRED_MESSAGE,
        },
    )
    return True


def expire_progress_job_if_stale(job: dict, cutoff: float) -> bool:
    """在任务锁内复核活动时间；确实过期时通知 SSE。"""
    with job["condition"]:
        return _expire_progress_job_if_stale_locked(job, cutoff)


def cleanup_expired_jobs(max_age_seconds: int = JOB_TTL_SECONDS) -> None:
    """清理长时间无更新的异步任务状态，避免卡死任务永久占用内存。"""
    cutoff = time.time() - max_age_seconds

    with PROGRESS_JOBS_LOCK:
        expiration_candidates = []
        for job_id, job in PROGRESS_JOBS.items():
            with job["condition"]:
                if get_progress_job_activity_time_locked(job) < cutoff:
                    expiration_candidates.append((job_id, job))

    for job_id, job in expiration_candidates:
        with OUTPUT_RESULT_GUARD_LOCK:
            with PROGRESS_JOBS_LOCK:
                if PROGRESS_JOBS.get(job_id) is not job:
                    continue
                with job["condition"]:
                    if not _expire_progress_job_if_stale_locked(job, cutoff):
                        continue
                    # 可以立即向客户端暴露超时终态，但实际工作线程退出前仍需
                    # 保留任务标识，防止同一客户端 ID 被第二个任务复用。
                    if job.get("worker_active") is True:
                        continue
                    if not remove_terminal_job_output_for_eviction_locked(
                        job_id,
                        time.time(),
                    ):
                        continue
                    PROGRESS_JOBS.pop(job_id, None)


def is_generated_output(filename: str) -> bool:
    """仅允许下载系统生成的输出文件。"""
    normalized = normalize_filename_text(filename)
    if "/" in normalized:
        return False
    return bool(GENERATED_OUTPUT_RE.fullmatch(normalized))


def is_regular_file_path(path: Path) -> bool:
    """只接受真实普通文件，避免跟随符号链接。"""
    with suppress(OSError):
        return not path.is_symlink() and path.is_file()
    return False


def is_valid_job_id(job_id: str) -> bool:
    return bool(JOB_ID_RE.fullmatch(job_id or ""))


def iter_job_file_paths(job_id: str):
    """列出某个任务编号可能占用的临时/输出文件路径。"""
    yield OUTPUT_FOLDER / f"{job_id}_output.docx"
    for suffix in JOB_TEMP_SUFFIXES:
        yield UPLOAD_FOLDER / f"{job_id}_{suffix}.docx"


def is_occupied_temp_path(path: Path) -> bool:
    """判断临时槽位是否已被实体文件或当前进程资源租约占用。"""
    path = Path(path)
    if path.exists() or path.is_symlink():
        return True
    with OUTPUT_STORAGE_LOCK:
        return (
            path in ACTIVE_TEMP_FILES
            or path in ACTIVE_OUTPUT_RESERVATIONS
            or ACTIVE_OUTPUT_DOWNLOADS.get(path, 0) > 0
        )


def create_unique_job_id(existing_job_ids=None) -> str:
    """生成当前进程与临时文件目录中未被占用的 128 位随机任务编号。"""
    for _ in range(JOB_ID_GENERATION_ATTEMPTS):
        job_id = secrets.token_hex(JOB_ID_BYTES)
        if not is_valid_job_id(job_id):
            continue
        if existing_job_ids is not None:
            if job_id in existing_job_ids:
                continue
        else:
            with PROGRESS_JOBS_LOCK:
                if job_id in PROGRESS_JOBS:
                    continue

        if any(is_occupied_temp_path(path) for path in iter_job_file_paths(job_id)):
            continue

        return job_id

    raise JobProcessingError("当前任务较多，请稍后重试。", 503)


def get_download_name(filename: str, fallback: str) -> str:
    """规范下载文件名，去掉路径片段、控制字符并强制使用 .docx 后缀。"""
    normalized = normalize_filename_text(filename)
    fallback_name = normalize_filename_text(fallback) or "document"
    basename = "".join(ch for ch in Path(normalized).name.strip() if ch.isprintable())
    candidate = basename or fallback_name
    fallback_display = get_display_name(fallback_name)
    stem = Path(candidate).stem.strip() or fallback_display
    stem = "".join(ch for ch in stem if ch.isprintable()).strip() or fallback_display
    stem = stem[:MAX_DOWNLOAD_STEM_LENGTH].rstrip() or fallback_display
    return f"{stem}.docx"


def is_truthy_flag(value) -> bool:
    if isinstance(value, bool):
        return value

    try:
        return str(value).strip().lower() in {"1", "true", "yes", "on"}
    except Exception:
        return False


def parse_bool_flag(payload, key: str, default: bool = True) -> bool:
    """从表单/JSON 载荷中解析布尔开关；字段缺省时返回 default。"""
    raw = payload.get(key) if payload else None
    if raw is None:
        return default
    return is_truthy_flag(raw)


def get_request_payload(*, max_content_length: int | None = None, too_large_message: str | None = None):
    """读取 JSON 对象或表单载荷；可在解析前对当前路由施加更严格的请求体上限。"""
    try:
        if max_content_length is not None:
            request.max_content_length = max_content_length
            if request.content_length is not None and request.content_length > max_content_length:
                raise RequestEntityTooLarge()

        if request.is_json:
            json_payload = request.get_json(silent=False)
            mark_request_body_consumed()
            return json_payload if hasattr(json_payload, "get") else {}

        form_payload = request.form or {}
        # Accessing ``request.form`` consumes urlencoded/multipart bodies, but
        # Werkzeug intentionally leaves bodies with an unknown content type
        # untouched.  Do not claim those requests were consumed: the response
        # security hook must close the HTTP/1 connection instead of allowing
        # the server to reuse it with unread attacker-controlled bytes.
        if request.mimetype in {"application/x-www-form-urlencoded", "multipart/form-data"}:
            mark_request_body_consumed()
        return form_payload
    except RequestEntityTooLarge as exc:
        message = too_large_message or get_http_exception_message(exc)
        raise JobProcessingError(message, 413) from exc
    except HTTPException as exc:
        raise JobProcessingError(get_http_exception_message(exc), exc.code or 400) from exc


def get_payload_text(payload, key: str = "text") -> str:
    """从请求载荷中提取文本字段；非字符串字段按空值处理。"""
    raw_text = payload.get(key, "") if payload else ""
    if not isinstance(raw_text, str):
        return ""

    text = raw_text.strip()
    # python-docx ultimately emits XML 1.0.  Reject characters outside the
    # XML 1.0 Char production up front so malformed input gets a stable 400
    # instead of a UnicodeEncodeError/ValueError from a later formatting step.
    for character in text:
        codepoint = ord(character)
        if not (
            codepoint in {0x9, 0xA, 0xD}
            or 0x20 <= codepoint <= 0xD7FF
            or 0xE000 <= codepoint <= 0xFFFD
            or 0x10000 <= codepoint <= 0x10FFFF
        ):
            raise JobProcessingError(TEXT_INVALID_CHARACTERS_MESSAGE, 400)

    return text


def is_text_payload_too_large(text: str) -> bool:
    return len(text.encode("utf-8")) > MAX_TEXT_INPUT_BYTES


def is_text_paragraph_count_too_large(text: str) -> bool:
    return text_exceeds_paragraph_limit(
        text,
        max_paragraphs=MAX_TEXT_PARAGRAPHS,
    )


def normalize_cover_field(value) -> str:
    """整理单行封面字段，避免异常 JSON 类型或控制字符进入 DOCX 生成链路。"""
    if not isinstance(value, str):
        return ""

    cleaned = "".join(ch for ch in value.strip() if ch.isprintable()).strip()
    if len(cleaned) <= MAX_COVER_FIELD_LENGTH:
        return cleaned

    return f"{cleaned[:MAX_COVER_FIELD_LENGTH - 1].rstrip()}…"


def extract_cover_info(payload) -> dict | None:
    """从表单或 JSON 载荷中提取自动封面信息。"""
    if not payload:
        return None

    enabled = is_truthy_flag(payload.get("generate_cover"))
    if not enabled:
        return None

    cover_info = {}
    for field in COVER_FORM_FIELDS:
        value = normalize_cover_field(payload.get(field))
        if value:
            cover_info[field] = value

    return cover_info


def coerce_mapping(value) -> Mapping:
    return value if isinstance(value, Mapping) else {}


def coerce_list(value) -> list:
    if isinstance(value, list):
        return value
    if isinstance(value, tuple):
        return list(value)
    return []


def coerce_text(value) -> str:
    return value if isinstance(value, str) else ""


def coerce_bool(value, default: bool = False) -> bool:
    return value if isinstance(value, bool) else default


def truncate_text(text: str, max_length: int) -> str:
    stripped = (text or "").strip()
    if len(stripped) <= max_length:
        return stripped

    return f"{stripped[:max_length - 1].rstrip()}…"


def truncate_json_string(text: str) -> str:
    if len(text) <= MAX_JSON_STRING_LENGTH:
        return text

    return f"{text[:MAX_JSON_STRING_LENGTH - 1].rstrip()}…"


def normalize_json_key(value) -> str:
    try:
        raw_key = str(value)
    except Exception:
        raw_key = "invalid_key"

    buffer = []
    pending_spaces = []
    pending_overflow = False
    started = False
    truncated = False
    probe_limit = MAX_JSON_KEY_LENGTH + 1

    for ch in raw_key:
        if not ch.isprintable():
            continue
        if not started:
            if ch.isspace():
                continue
            started = True
        if ch.isspace():
            if len(buffer) + len(pending_spaces) < probe_limit:
                pending_spaces.append(ch)
            else:
                pending_overflow = True
            continue

        if pending_spaces:
            for space in pending_spaces:
                if len(buffer) < probe_limit:
                    buffer.append(space)
                else:
                    truncated = True
                    break
            pending_spaces = []
        if pending_overflow:
            truncated = True
            pending_overflow = False

        if len(buffer) < probe_limit:
            buffer.append(ch)
        else:
            truncated = True

        if truncated and len(buffer) >= probe_limit:
            break

    key = "".join(buffer).rstrip()
    if not key:
        return "invalid_key"
    if truncated or len(key) > MAX_JSON_KEY_LENGTH:
        return f"{key[:MAX_JSON_KEY_LENGTH - 1].rstrip()}…"
    return key


def unique_json_key(value, used_keys: set[str]) -> str:
    """生成归一化且不覆盖已有字段的 JSON key。"""
    base_key = normalize_json_key(value)
    candidate = base_key
    suffix_index = 2
    while candidate in used_keys:
        suffix = f"_{suffix_index}"
        stem_length = max(1, MAX_JSON_KEY_LENGTH - len(suffix))
        stem = base_key[:stem_length].rstrip() or "key"
        candidate = f"{stem}{suffix}"
        suffix_index += 1

    used_keys.add(candidate)
    return candidate


def coerce_int(value, default: int = 0) -> int:
    if isinstance(value, bool):
        return default

    try:
        coerced = int(value)
    except (TypeError, ValueError, OverflowError):
        return default

    return coerced if coerced >= 0 else default


def coerce_float(value, default: float) -> float:
    if isinstance(value, bool):
        return default

    try:
        coerced = float(value)
    except (TypeError, ValueError, OverflowError):
        return default

    return coerced if math.isfinite(coerced) and coerced > 0 else default


def normalize_stats(value) -> dict:
    stats = coerce_mapping(value)
    used_keys = set()
    return {
        unique_json_key(key, used_keys): coerce_int(count)
        for key, count in stats.items()
    }


def normalize_outline_level(value) -> str:
    level = truncate_text(coerce_text(value), 24)
    return level if level in OUTLINE_LEVELS else "section"


def normalize_outline(value, max_items: int = MAX_SUMMARY_OUTLINE_ITEMS) -> list:
    outline = []
    for item in coerce_list(value):
        item = coerce_mapping(item)
        if not item:
            continue

        text = truncate_text(coerce_text(item.get("text")), MAX_OUTLINE_TEXT_LENGTH)
        if not text:
            continue

        level = normalize_outline_level(item.get("level"))
        outline.append({"level": level, "text": text})
        if len(outline) >= max_items:
            break

    return outline


def normalize_page_setup(value) -> dict:
    page_setup = coerce_mapping(value)
    if not page_setup:
        return {}

    normalized = {}
    for field in PAGE_SETUP_TEXT_FIELDS:
        if field in page_setup:
            normalized[field] = truncate_text(
                coerce_text(page_setup.get(field)),
                MAX_PAGE_SETUP_TEXT_LENGTH,
            )

    for field in PAGE_SETUP_FLOAT_FIELDS:
        if field in page_setup:
            numeric_value = coerce_float(page_setup.get(field), None)
            if numeric_value is not None:
                normalized[field] = numeric_value

    margins = coerce_mapping(page_setup.get("margins_cm"))
    if margins or "margins_cm" in page_setup:
        normalized_margins = {}
        for field in PAGE_MARGIN_FIELDS:
            if field in margins:
                numeric_value = coerce_float(margins.get(field), None)
                if numeric_value is not None:
                    normalized_margins[field] = numeric_value
        normalized["margins_cm"] = normalized_margins

    return normalized


def normalize_progress_step(value) -> int:
    step = coerce_int(value, 1)
    return min(max(step, 1), MAX_PROCESSING_STEP)


def normalize_progress_message(value) -> str:
    return truncate_text(coerce_text(value), MAX_PROGRESS_MESSAGE_LENGTH) or "正在处理文档"


def normalize_progress_detail(value) -> str:
    return truncate_text(coerce_text(value), MAX_PROGRESS_DETAIL_LENGTH)


def normalize_job_error_message(value) -> str:
    return sanitize_user_facing_message(value, "处理失败")


def normalize_api_error_message(value) -> str:
    return sanitize_user_facing_message(value, "请求处理失败，请稍后重试。")


def normalize_http_status_code(value, default: int = 500) -> int:
    status_code = coerce_int(value, default)
    return status_code if 400 <= status_code <= 599 else default


def normalize_sse_event_name(value) -> str:
    if isinstance(value, str) and SSE_EVENT_NAME_RE.fullmatch(value):
        return value
    return "message"


def normalize_sse_event_id(value, fallback_id: int) -> int:
    fallback = max(coerce_int(fallback_id, 1), 1)
    event_id = coerce_int(value, fallback)
    return event_id if event_id >= 1 else fallback


def parse_last_event_id(value: str, event_count: int) -> int:
    """解析 SSE Last-Event-ID，限制长度以避开异常大整数解析。"""
    raw_value = (value or "").strip()
    if not raw_value or len(raw_value) > MAX_LAST_EVENT_ID_LENGTH:
        return 0
    if not raw_value.isascii() or not raw_value.isdigit():
        return 0

    event_id = int(raw_value)
    return min(max(0, event_id), event_count)


def normalize_format_summary(result) -> dict:
    """兼容旧布尔返回值，统一整理为结构化摘要。"""
    if isinstance(result, Mapping):
        return {
            "stats": normalize_stats(result.get("stats")),
            "page_setup": normalize_page_setup(result.get("page_setup")),
            "outline": normalize_outline(result.get("outline")),
            "title_text": truncate_text(coerce_text(result.get("title_text")), MAX_SUMMARY_TITLE_LENGTH),
            "table_paragraphs": coerce_int(result.get("table_paragraphs")),
            "equation_paragraphs": coerce_int(result.get("equation_paragraphs")),
            "formatted_footnotes": coerce_int(result.get("formatted_footnotes")),
            "resized_images": coerce_int(result.get("resized_images")),
            "cover_generated": coerce_bool(result.get("cover_generated")),
        }

    return {
        "stats": {},
        "page_setup": {},
        "outline": [],
        "title_text": "",
        "table_paragraphs": 0,
        "equation_paragraphs": 0,
        "formatted_footnotes": 0,
        "resized_images": 0,
        "cover_generated": False,
    }


def truncate_preview_text(text: str, max_length: int = 24) -> str:
    return truncate_text(text, max_length)


def describe_structure_count(count: int, label: str, unit: str = "个") -> str:
    return f"{label} {count} {unit}"


def build_preview(summary: dict) -> dict:
    """将排版摘要转换为前端可直接展示的结果预览。"""
    summary = coerce_mapping(summary)
    stats = coerce_mapping(summary.get("stats"))
    page_setup = coerce_mapping(summary.get("page_setup"))
    outline = normalize_outline(summary.get("outline"), MAX_PREVIEW_OUTLINE_ITEMS)
    title_text = coerce_text(summary.get("title_text")).strip()
    cover_generated = coerce_bool(summary.get("cover_generated"))
    equation_paragraphs = coerce_int(summary.get("equation_paragraphs"))
    formatted_footnotes = coerce_int(summary.get("formatted_footnotes"))

    margins = coerce_mapping(page_setup.get("margins_cm"))
    top_margin = coerce_float(margins.get("top"), 2.54)
    left_margin = coerce_float(margins.get("left"), 3.18)
    raw_header_text = coerce_text(page_setup.get("header_text"))
    header_text = truncate_preview_text(raw_header_text)

    structure_bits = []
    for key, label, unit in (
        (ParagraphType.TITLE, "论文标题", "个"),
        (ParagraphType.ENGLISH_ABSTRACT_HEADING, "英文摘要", "个"),
        (ParagraphType.HEADING_L1, "一级标题", "个"),
        (ParagraphType.HEADING_L2, "二级标题", "个"),
        (ParagraphType.HEADING_L3, "三级标题", "个"),
        (ParagraphType.FIGURE_CAPTION, "图标题", "个"),
        (ParagraphType.TABLE_CAPTION, "表标题", "个"),
        (ParagraphType.SECTION_HEADING, "非编号章节", "个"),
        (ParagraphType.REFERENCES_HEADING, "参考文献标题", "个"),
    ):
        count = coerce_int(stats.get(key))
        if count:
            structure_bits.append(describe_structure_count(count, label, unit))
    if equation_paragraphs:
        structure_bits.append(describe_structure_count(equation_paragraphs, "公式段落", "个"))
    if formatted_footnotes:
        structure_bits.append(describe_structure_count(formatted_footnotes, "脚注", "条"))

    structure_description = "、".join(structure_bits) or "正文段落已统一为小四、首行缩进和 1.5 倍行距"
    page_description = f"纸张为 A4，上下 {top_margin:.2f} cm，左右 {left_margin:.2f} cm"
    if cover_generated:
        page_description += "，并在首页插入了自动生成的课程论文封面"
    if not header_text:
        header_description = "正文未设置固定默认页眉，页脚仍会插入可更新的自动页码字段"
    elif title_text and raw_header_text != title_text:
        header_description = f"页眉显示“{header_text}”，超长标题会自动缩成更适合打印的运行页眉"
    else:
        header_description = f"页眉显示“{header_text}”，页脚插入可更新的自动页码字段"

    reference_count = coerce_int(stats.get(ParagraphType.REFERENCE_ENTRY))
    if reference_count:
        reference_description = f"识别到 {reference_count} 条参考文献，已统一为左对齐、五号字号和更紧凑的文献列表样式"
    else:
        reference_description = "暂未识别到“参考文献”标题，本次仍已完成正文和标题层级排版"

    return {
        "highlights": [
            {
                "eyebrow": "页面设置",
                "title": "A4 页面与规范页边距已应用",
                "description": page_description,
            },
            {
                "eyebrow": "页眉页码",
                "title": "页眉与居中页码已自动生成",
                "description": header_description,
            },
            {
                "eyebrow": "结构识别",
                "title": "标题层级已完成识别和套用",
                "description": structure_description,
            },
            {
                "eyebrow": "参考文献",
                "title": "尾部参考文献已单独整理",
                "description": reference_description,
            },
        ],
        "outline": outline,
    }


def build_concat_preview(summary: dict | None = None) -> dict:
    """为纯拼接结果生成展示用预览（不做排版分析，只说明拼接动作）。"""
    summary = coerce_mapping(summary)
    page_number_restarted = coerce_bool(summary.get("page_number_restarted"))

    if page_number_restarted:
        page_highlight = {
            "eyebrow": "页码衔接",
            "title": "封面不计页码，正文从第 1 页重新编号",
            "description": "正文另起一页，并把正文页码从第 1 页重新计数，封面不显示页码——正文原有的页边距、页眉页脚与页码字段保持不变。",
        }
    else:
        page_highlight = {
            "eyebrow": "分页衔接",
            "title": "正文从新的一页开始",
            "description": "在两个文档衔接处分节换页，正文另起一页。页码沿用各自原有设置，未做重新编号。",
        }

    return {
        "highlights": [
            {
                "eyebrow": "原样保留",
                "title": "正文的排版形态完整保留",
                "description": "正文的字体、字号、行距、页边距、页眉页脚、图片与表格全部保持原有样式，只是把封面接到了它前面。",
            },
            {
                "eyebrow": "智能合并",
                "title": "文档级合并，规避手动拼接乱码",
                "description": "采用 docxcompose 合并文档主体、样式、编号与图片关系，避免直接复制 XML 造成的样式错乱与乱码。",
            },
            page_highlight,
        ],
        "outline": [],
    }


def get_or_create_progress_job(kind: str, requested_job_id: str | None = None) -> tuple[dict, bool]:
    """创建异步任务；客户端重试相同标识时返回原任务而不重复执行。"""
    # 新任务可能触发终态 job/输出淘汰，必须与已经开始的
    # 结果领取串行化，再遵循任务表 → condition → 输出锁的顺序。
    with OUTPUT_RESULT_GUARD_LOCK:
        return _get_or_create_progress_job_with_result_guard(
            kind,
            requested_job_id,
        )


def _get_or_create_progress_job_with_result_guard(
    kind: str,
    requested_job_id: str | None = None,
) -> tuple[dict, bool]:
    """在已持有结果领取保护锁时创建或复用任务。"""
    with PROGRESS_JOBS_LOCK:
        if requested_job_id is not None:
            if not is_valid_job_id(requested_job_id):
                raise JobProcessingError("任务恢复标识无效，请刷新页面后重试。", 400)

            existing_job = PROGRESS_JOBS.get(requested_job_id)
            if existing_job is not None:
                with existing_job["condition"]:
                    if existing_job.get("kind") != kind:
                        raise JobProcessingError("任务恢复标识已被其他任务使用，请刷新页面后重试。", 409)
                return existing_job, False

        prune_progress_jobs_for_capacity_locked()
        if len(PROGRESS_JOBS) >= MAX_PROGRESS_JOBS:
            raise JobProcessingError("当前任务较多，请稍后重试。", 503)

        if requested_job_id is None:
            job_id = create_unique_job_id(PROGRESS_JOBS)
        else:
            job_id = requested_job_id
            if any(is_occupied_temp_path(path) for path in iter_job_file_paths(job_id)):
                raise JobProcessingError("任务恢复标识对应的文件仍被占用，请刷新页面后重试。", 409)

        job = {
            "id": job_id,
            "kind": kind,
            "status": "queued",
            "created_at": time.time(),
            "updated_at": time.time(),
            "events": [],
            "next_event_id": 0,
            "progress_event_count": 0,
            "condition": threading.Condition(),
            "result": None,
            "error": None,
            "result_requested_at": None,
            "worker_active": False,
            "events_url": f"/api/jobs/{job_id}/events",
            "result_url": f"/api/jobs/{job_id}/result",
        }
        PROGRESS_JOBS[job_id] = job

    return job, True


def create_progress_job(kind: str) -> dict:
    """创建一个由服务端分配标识的异步处理任务。"""
    job, _created = get_or_create_progress_job(kind)
    return job


def get_or_create_request_progress_job(kind: str) -> tuple[dict, bool]:
    """使用当前请求携带的恢复标识创建或复用异步任务。"""
    requested_job_id = getattr(g, REQUEST_CLIENT_JOB_ID_ATTR, None)
    return get_or_create_progress_job(kind, requested_job_id)


def get_reused_request_progress_job(kind: str) -> dict | None:
    """返回 before_request 已确认的同类型任务；不存在时让路由走正常创建流程。"""
    if not getattr(g, REQUEST_REUSED_PROGRESS_JOB_ATTR, False):
        return None

    requested_job_id = getattr(g, REQUEST_CLIENT_JOB_ID_ATTR, None)
    if requested_job_id is None:
        return None

    with PROGRESS_JOBS_LOCK:
        job = PROGRESS_JOBS.get(requested_job_id)
        if job is not None:
            with job["condition"]:
                if job.get("kind") != kind:
                    raise JobProcessingError(
                        "任务恢复标识已被其他任务使用，请刷新页面后重试。",
                        409,
                    )
            return job

    # before_request 与路由之间可能恰好发生任务过期/容量淘汰；不能因此
    # 让后续请求跳过处理槽与上传预留，再解析一个新的大请求体。
    if any(is_occupied_temp_path(path) for path in iter_job_file_paths(requested_job_id)):
        raise JobProcessingError(
            "任务恢复标识对应的文件仍被占用，请刷新页面后重试。",
            409,
        )
    reserve_request_processing_resources()
    return None


def get_progress_job(job_id: str):
    with PROGRESS_JOBS_LOCK:
        return PROGRESS_JOBS.get(job_id)


def get_terminal_job_eviction_priority_locked(job: dict, now: float):
    """返回终态任务容量淘汰优先级；调用方必须持有任务 condition。"""
    if job.get("status") not in TERMINAL_JOB_STATUSES or job.get("worker_active") is True:
        return None

    requested_at = job.get("result_requested_at")
    if (
        isinstance(requested_at, (int, float))
        and not isinstance(requested_at, bool)
        and math.isfinite(requested_at)
    ):
        if requested_at > now - MIN_REQUESTED_TERMINAL_RETENTION_SECONDS:
            return None
        return (0, requested_at)

    updated_at = job.get("updated_at", job.get("created_at", 0))
    if (
        isinstance(updated_at, (int, float))
        and not isinstance(updated_at, bool)
        and math.isfinite(updated_at)
        and updated_at <= now - MIN_UNREAD_TERMINAL_RETENTION_SECONDS
    ):
        return (1, updated_at)
    return None


def remove_terminal_job_output_for_eviction_locked(job_id: str, now: float) -> bool:
    """淘汰 job 前同步移除其输出；调用方必须已持有结果保护、任务表与单任务锁。

    返回 False 表示输出仍受生成、下载或最近请求保护，此时必须保留 job，
    避免产生“任务不存在但文件仍占用恢复 ID”的孤儿状态。锁序保持为
    结果保护锁 → 任务表锁 → 单任务 condition → 输出存储锁。
    """
    if not is_valid_job_id(job_id):
        return False

    output_path = Path(OUTPUT_FOLDER) / f"{job_id}_output.docx"
    with OUTPUT_STORAGE_LOCK:
        if (
            output_path in ACTIVE_TEMP_FILES
            or output_path in ACTIVE_OUTPUT_RESERVATIONS
            or ACTIVE_OUTPUT_DOWNLOADS.get(output_path, 0) > 0
        ):
            return False

        try:
            metadata = output_path.lstat()
        except FileNotFoundError:
            RECENT_OUTPUT_REQUESTS.pop(output_path, None)
            return True
        except OSError:
            return False

        # 最近一次领取可能由直接下载刷新，时间比 job 快照更新；仍需保护。
        if get_recent_output_request_locked(output_path, now) is not None:
            return False

        if not (stat.S_ISREG(metadata.st_mode) or stat.S_ISLNK(metadata.st_mode)):
            return False

        # Keep the terminal job when its output is a hard link.  Removing the
        # generated name would alter link ownership outside the output store
        # and would also make the job's output disappear from the cleanup
        # protection model while another directory entry still references the
        # same inode.  Normal generated outputs are single-link files; a
        # hard-linked file is treated as externally managed and left intact.
        if stat.S_ISREG(metadata.st_mode) and getattr(metadata, "st_nlink", 1) != 1:
            return False

        try:
            output_path.unlink()
        except FileNotFoundError:
            pass
        except OSError:
            return False
        RECENT_OUTPUT_REQUESTS.pop(output_path, None)
        return True


def prune_progress_jobs_for_capacity_locked() -> None:
    """在已持有结果保护锁与任务表锁时执行容量淘汰。"""
    while len(PROGRESS_JOBS) >= MAX_PROGRESS_JOBS:
        now = time.time()
        candidates = []
        for job_id, job in PROGRESS_JOBS.items():
            with job["condition"]:
                priority = get_terminal_job_eviction_priority_locked(job, now)

            if priority is None:
                continue
            candidates.append((priority, job_id, job))

        if not candidates:
            break

        evicted = False
        for _priority, job_id, job in sorted(candidates, key=lambda item: item[:2]):
            with job["condition"]:
                recheck_now = time.time()
                if (
                    PROGRESS_JOBS.get(job_id) is not job
                    or get_terminal_job_eviction_priority_locked(job, recheck_now) is None
                    or not remove_terminal_job_output_for_eviction_locked(
                        job_id,
                        recheck_now,
                    )
                ):
                    continue
                PROGRESS_JOBS.pop(job_id, None)
                evicted = True
                break

        if not evicted:
            break


def _mark_progress_job_result_requested_locked(job: dict, now: float):
    if job.get("status") not in TERMINAL_JOB_STATUSES:
        return False, None
    job["result_requested_at"] = now
    return job.get("status") == "done", job.get("id")


def record_requested_output(job_id, protect_output: bool, now: float) -> None:
    if not protect_output or not is_valid_job_id(job_id):
        return

    output_path = Path(OUTPUT_FOLDER) / f"{job_id}_output.docx"
    with OUTPUT_STORAGE_LOCK:
        record_recent_output_request_locked(output_path, now)


def mark_progress_job_result_requested(job: dict) -> None:
    """刷新终态结果领取时间，并保护对应输出的短期重试窗口。"""
    now = time.time()
    with OUTPUT_RESULT_GUARD_LOCK:
        with job["condition"]:
            protect_output, job_id = _mark_progress_job_result_requested_locked(job, now)

        record_requested_output(job_id, protect_output, now)


def get_progress_job_result_snapshot(job_id: str, *, mark_requested: bool = True):
    """原子读取结果快照并刷新终态领取时间，避免容量回收竞态。"""
    now = time.time()
    protect_output = False
    output_job_id = None
    with OUTPUT_RESULT_GUARD_LOCK:
        with PROGRESS_JOBS_LOCK:
            job = PROGRESS_JOBS.get(job_id)
            if job is None:
                return None
            with job["condition"]:
                status = job.get("status")
                error_info = job.get("error")
                result_payload = job.get("result")
                if mark_requested and status in TERMINAL_JOB_STATUSES:
                    protect_output, output_job_id = _mark_progress_job_result_requested_locked(job, now)

        record_requested_output(output_job_id, protect_output, now)
    return status, error_info, result_payload


def _append_job_event_locked(job: dict, event: str, data: dict) -> None:
    job["next_event_id"] += 1
    job["events"].append(
        {
            "id": job["next_event_id"],
            "event": event,
            "data": data,
        }
    )
    job["updated_at"] = time.time()
    job["condition"].notify_all()


def emit_job_progress(job: dict, step: int, message: str, detail: str | None = None) -> None:
    step = normalize_progress_step(step)
    message = normalize_progress_message(message)
    detail = normalize_progress_detail(detail)

    with job["condition"]:
        if job["status"] in TERMINAL_JOB_STATUSES:
            return

        if job["status"] == "queued":
            job["status"] = "running"

        progress_event_count = max(0, coerce_int(job.get("progress_event_count")))
        if progress_event_count >= MAX_PROGRESS_EVENTS_PER_JOB:
            job["updated_at"] = time.time()
            return

        payload = {
            "job_id": job["id"],
            "step": step,
            "step_label": PROCESSING_STEP_LABELS.get(step, "处理中"),
            "message": message,
        }
        if detail:
            payload["detail"] = detail

        _append_job_event_locked(job, "progress", payload)
        job["progress_event_count"] = progress_event_count + 1


def complete_progress_job(job: dict, result_payload: dict) -> bool:
    """发布异步任务结果；任务已进入终态时返回 False。"""
    result_snapshot = snapshot_progress_job_result(result_payload)
    with job["condition"]:
        if job["status"] in TERMINAL_JOB_STATUSES:
            return False

        job["status"] = "done"
        job["result"] = result_snapshot
        job["error"] = None
        _append_job_event_locked(
            job,
            "complete",
            {
                "job_id": job["id"],
                "message": "排版完成",
                "result_url": job["result_url"],
            },
        )
        return True


def fail_progress_job(job: dict, message: str, status_code: int = 500) -> None:
    message = normalize_job_error_message(message)
    status_code = normalize_http_status_code(status_code)

    with job["condition"]:
        if job["status"] in TERMINAL_JOB_STATUSES:
            return

        job["status"] = "error"
        job["error"] = {
            "message": message,
            "status_code": status_code,
        }
        _append_job_event_locked(
            job,
            "failed",
            {
                "job_id": job["id"],
                "message": message,
            },
        )


def handle_progress_job_creation_failure(job: dict, error: Exception) -> bool:
    """客户端已知任务标识时保留终态错误；旧式随机任务仍直接回收。"""
    client_job_id = (
        getattr(g, REQUEST_CLIENT_JOB_ID_ATTR, None)
        if has_request_context()
        else None
    )
    if client_job_id == job.get("id"):
        if isinstance(error, JobProcessingError):
            fail_progress_job(job, error.message, error.status_code)
        else:
            fail_progress_job(job, "服务器内部错误，请稍后重试。", 500)
        setattr(g, REQUEST_FAILED_PROGRESS_JOB_ATTR, job)
        return True

    with PROGRESS_JOBS_LOCK:
        if PROGRESS_JOBS.get(job.get("id")) is job:
            PROGRESS_JOBS.pop(job["id"], None)
    return False


def build_job_progress_callback(job: dict):
    def callback(payload: dict):
        payload = coerce_mapping(payload)
        emit_job_progress(
            job,
            step=payload.get("step", 1),
            message=payload.get("message", "正在处理文档"),
            detail=payload.get("detail"),
        )

    return callback


def cleanup_path(path: Path | str) -> None:
    """清理单个临时路径：普通文件/符号链接直接删除，空目录仅移除目录本身。"""
    target = Path(path)
    try:
        with suppress(OSError):
            if target.is_dir() and not target.is_symlink():
                target.rmdir()
            else:
                target.unlink()
    finally:
        with OUTPUT_STORAGE_LOCK:
            ACTIVE_TEMP_FILES.discard(target)


def log_unexpected_exception(context: str, exc: Exception, *paths) -> None:
    """记录清洗后的异常摘要；DEBUG 级别才附带完整堆栈。"""
    request_id = get_request_id()
    if request_id is not None:
        context = f"[request_id={request_id}] {context}"
    logger.error(
        "%s: %s",
        context,
        format_log_exception(exc, *paths),
        exc_info=logger.isEnabledFor(logging.DEBUG),
    )


def cleanup_path_best_effort(path: Path | str, context: str) -> None:
    """尽力清理后台临时路径，清理故障不得阻断任务终态发布。"""
    try:
        cleanup_path(path)
    except Exception as exc:
        # 日志后端或异常格式化本身也不应破坏 worker 的收尾契约。
        with suppress(Exception):
            log_unexpected_exception(context, exc, path)


def cleanup_unpublished_background_job_output(job: dict) -> None:
    """后台任务未成功发布终态时，清理其标准输出路径。"""
    job_id = job.get("id", "")
    if not is_valid_job_id(job_id):
        return

    with job["condition"]:
        # complete_progress_job() 可能在写入 done 终态后才由自定义
        # condition/通知逻辑抛错；此时结果已发布，不能反向删除产物。
        if job.get("status") == "done":
            return

    cleanup_path_best_effort(
        Path(OUTPUT_FOLDER) / f"{job_id}_output.docx",
        "后台任务输出清理失败",
    )


def launch_background_job(
    job: dict,
    task_name: str,
    work_fn,
    cleanup_paths=(),
    log_paths=(),
) -> threading.Thread:
    slots = claim_processing_job_slot()

    # The request may have reserved its processing slot before graceful
    # shutdown began.  Re-check here as well so a request already past the
    # ``before_request`` gate cannot start new background work during drain.
    if SHUTDOWN_EVENT.is_set():
        slots.release()
        raise JobProcessingError(SERVICE_DRAINING_MESSAGE, 503)

    with job["condition"]:
        if job.get("status") in TERMINAL_JOB_STATUSES or job.get("worker_active") is True:
            slots.release()
            raise JobProcessingError("当前任务状态异常，请重新提交任务。", 409)
        if SHUTDOWN_EVENT.is_set():
            slots.release()
            raise JobProcessingError(SERVICE_DRAINING_MESSAGE, 503)
        job["worker_active"] = True

    def runner():
        progress_callback = build_job_progress_callback(job)

        def cleanup_inputs():
            for path in cleanup_paths:
                cleanup_path_best_effort(path, f"{task_name}临时输入清理失败")

        try:
            result_payload = work_fn(progress_callback)
            # 终态事件一旦发布，客户端就可能立即读取结果并检查临时
            # 文件是否已清理。因此输入清理必须先于 complete/failed
            # 通知；finally 中仍保留幂等兜底，覆盖通知阶段的异常。
            cleanup_inputs()
            if not complete_progress_job(job, result_payload):
                with job["condition"]:
                    discard_output = job.get("status") == "error"
                if discard_output:
                    # 过期清理可能在 worker 收尾前先把任务标记
                    # 为失败。此时不能发布刚生成的产物，也不应
                    # 让它变成占用预算与任务 ID 的孤儿文件。
                    cleanup_unpublished_background_job_output(job)
        except JobProcessingError as exc:
            logger.error(f"{task_name}失败: {exc.message}")
            cleanup_unpublished_background_job_output(job)
            cleanup_inputs()
            fail_progress_job(job, exc.message, exc.status_code)
        except Exception as exc:  # pragma: no cover - 作为最终兜底
            log_unexpected_exception(f"{task_name}出现异常", exc, *cleanup_paths, *log_paths)
            cleanup_unpublished_background_job_output(job)
            cleanup_inputs()
            fail_progress_job(job, "服务器内部错误，请稍后重试。", 500)
        finally:
            try:
                cleanup_inputs()
            finally:
                try:
                    slots.release()
                finally:
                    with job["condition"]:
                        job["worker_active"] = False
                        job["condition"].notify_all()

    try:
        worker = threading.Thread(
            target=runner,
            name="paper-background-job",
            daemon=False,
        )
        worker.start()
    except Exception as exc:
        try:
            slots.release()
        finally:
            with job["condition"]:
                job["worker_active"] = False
                job["condition"].notify_all()
        raise JobProcessingError("当前任务较多，请稍后重试。", 503) from exc

    return worker


def build_async_job_response(job: dict) -> dict:
    return {
        "success": True,
        "job_id": job["id"],
        "status": job["status"],
        "events_url": job["events_url"],
        "result_url": job["result_url"],
    }


def build_async_job_http_response(job: dict):
    """将异步任务当前状态编码为创建/复用响应，避免终态失败被误报为成功。"""
    with job["condition"]:
        status = job.get("status")
        error_info = coerce_mapping(job.get("error"))
        # 与分支判断读取同一状态，避免 worker 在解锁后失败时编码出
        # HTTP 202 / success=True / status=error 的矛盾响应。
        response_payload = build_async_job_response(job) if status != "error" else None

    if status == "error":
        return json_async_job_creation_error(
            normalize_job_error_message(error_info.get("message", "处理失败")),
            normalize_http_status_code(error_info.get("status_code", 500)),
            job,
        )

    response = jsonify(response_payload)
    if status in {"queued", "running"}:
        # A queued/running snapshot is inherently short-lived.  Prevent an
        # intermediary from replaying a stale 202 after the worker has
        # already published a terminal result.
        response.headers["Cache-Control"] = "no-store"
        response.headers["Retry-After"] = str(JOB_RESULT_RETRY_AFTER_SECONDS)
    return response, 202


def build_success_response(
    *,
    job_id: str,
    original_name: str,
    output_filename: str,
    download_name: str,
    input_size_text: str,
    output_size_text: str,
    elapsed: float,
    format_result,
    include_format_result_alias: bool = False,
    preview: dict | None = None,
) -> dict:
    format_summary = normalize_format_summary(format_result)
    response_data = {
        "success": True,
        "job_id": job_id,
        "original_name": original_name,
        "download_url": f"/api/download/{output_filename}",
        "download_name": download_name,
        "format_summary": format_summary,
        "preview": preview if preview is not None else build_preview(format_summary),
        "stats": {
            "input_size": input_size_text,
            "output_size": output_size_text,
            "elapsed": f"{elapsed:.2f}s",
        },
    }

    if include_format_result_alias:
        response_data["format_result"] = format_summary

    return response_data


def get_result_output_filename(result_payload: dict | None) -> str | None:
    """从异步任务结果中提取系统生成的输出文件名。"""
    if not isinstance(result_payload, Mapping):
        return None

    download_url = coerce_text(result_payload.get("download_url")).split("?", 1)[0].strip()
    download_prefix = "/api/download/"
    if not download_url.startswith(download_prefix):
        return None

    output_filename = download_url.removeprefix(download_prefix)
    if "/" in output_filename or "\\" in output_filename:
        return None

    if not is_generated_output(output_filename):
        return None

    return output_filename


def get_generated_output_file_state(output_path: Path) -> tuple[str, int]:
    """按下载准入规则检查产物类型与大小。"""
    try:
        metadata = os.stat(output_path, follow_symlinks=False)
    except OSError:
        return "missing", 0

    if not stat.S_ISREG(metadata.st_mode):
        return "missing", 0
    # Generated outputs must own their inode.  A hard-linked name can point
    # at an unrelated file outside the output store; accepting it here would
    # advertise a successful result that the download lease correctly refuses
    # later, leaving the client with a misleading result response.
    if getattr(metadata, "st_nlink", 1) != 1:
        return "missing", 0
    if metadata.st_size <= 0:
        return "empty", metadata.st_size
    if metadata.st_size > MAX_OUTPUT_FILE_BYTES:
        return "too_large", metadata.st_size
    return "ok", metadata.st_size


def validate_result_output(result_payload: dict | None, job_id: str | None = None) -> tuple[bool, str, int]:
    """确认异步任务结果结构合法、归属当前任务且下载文件仍存在。"""
    if not isinstance(result_payload, Mapping) or result_payload.get("success") is not True:
        return False, "任务结果异常，请重新提交任务。", 500

    output_filename = get_result_output_filename(result_payload)
    if output_filename is None:
        return False, "任务结果异常，请重新提交任务。", 500

    if job_id is not None and output_filename != f"{job_id}_output.docx":
        return False, "任务结果异常，请重新提交任务。", 500

    output_path = OUTPUT_FOLDER / output_filename
    output_state, _ = get_generated_output_file_state(output_path)
    if output_state == "missing":
        return False, "生成文件已过期，请重新提交任务。", 404
    if output_state == "empty":
        cleanup_output_on_failure(output_path)
        return False, "任务结果异常，请重新提交任务。", 500
    if output_state == "too_large":
        cleanup_output_on_failure(output_path)
        return False, OUTPUT_TOO_LARGE_MESSAGE, 413

    return True, "", 200


def sanitize_result_payload(result_payload: dict, output_filename: str) -> dict:
    """返回给前端前规范化异步结果中的下载字段。"""
    payload = make_json_safe(result_payload)
    if not isinstance(payload, dict):
        payload = {}

    payload["success"] = True
    payload["download_url"] = f"/api/download/{output_filename}"
    payload["download_name"] = get_download_name(payload.get("download_name", ""), output_filename)
    return payload


def make_json_safe(value, depth: int = 0):
    """递归整理为 Flask JSON provider 可稳定序列化的值。"""
    if depth > 8:
        return None
    if value is None or isinstance(value, (bool, int)):
        return value
    if isinstance(value, str):
        return truncate_json_string(value)
    if isinstance(value, float):
        return value if math.isfinite(value) else None
    if isinstance(value, Mapping):
        used_keys = set()
        return {
            unique_json_key(key, used_keys): make_json_safe(child, depth + 1)
            for key, child in islice(value.items(), MAX_JSON_COLLECTION_ITEMS)
        }
    if isinstance(value, (list, tuple)):
        return [make_json_safe(child, depth + 1) for child in value[:MAX_JSON_COLLECTION_ITEMS]]

    return None


def snapshot_progress_job_result(result_payload) -> dict:
    """生成有界独立快照，避免终态结果被外部修改或长期持有任意对象。"""
    if not isinstance(result_payload, Mapping):
        return {}

    try:
        payload = {}
        used_keys = set()
        for key in RESULT_PAYLOAD_PRIORITY_KEYS:
            if key in result_payload:
                payload[key] = make_json_safe(result_payload[key], 1)
                used_keys.add(key)

        for key, child in result_payload.items():
            if len(payload) >= MAX_JSON_COLLECTION_ITEMS:
                break
            if isinstance(key, str) and key in RESULT_PAYLOAD_PRIORITY_KEYS:
                continue
            payload[unique_json_key(key, used_keys)] = make_json_safe(child, 1)
        return payload
    except Exception:
        return {}


def format_sse_event(event, fallback_id: int) -> str:
    event = coerce_mapping(event)
    event_id = normalize_sse_event_id(event.get("id"), fallback_id)
    event_name = normalize_sse_event_name(event.get("event"))
    payload = json.dumps(make_json_safe(event.get("data")), ensure_ascii=False)
    return f"id: {event_id}\nevent: {event_name}\ndata: {payload}\n\n"


def cleanup_output_on_failure(output_path: Path) -> None:
    """处理失败时清理当前任务的输出文件，避免遗留半成品。"""
    cleanup_path(output_path)


def get_generated_output_size(output_path: Path, failure_message: str) -> int:
    """确认格式化器真实生成了可下载文件，并返回文件大小。"""
    try:
        output_state, output_size = get_generated_output_file_state(output_path)
        if output_state in {"missing", "empty"}:
            raise FileNotFoundError(output_path)
        if output_state == "too_large":
            raise OutputSizeLimitExceeded(MAX_OUTPUT_FILE_BYTES)
        if not is_valid_generated_docx(output_path):
            raise OSError("generated DOCX failed integrity validation")

        # ZIP 校验会在整个产物上进行分块读取；重新读取元数据后再记账，
        # 避免校验期间文件被替换/增长时把旧的 stat 大小写入容量账本。
        final_state, final_size = get_generated_output_file_state(output_path)
        if final_state in {"missing", "empty"}:
            raise FileNotFoundError(output_path)
        if final_state == "too_large":
            raise OutputSizeLimitExceeded(MAX_OUTPUT_FILE_BYTES)
        if final_size != output_size:
            raise OSError("generated DOCX changed during integrity validation")

        complete_output_storage_reservation(output_path, final_size)
        enforce_output_storage_budget()
        return final_size
    except OutputSizeLimitExceeded:
        cleanup_output_on_failure(output_path)
        raise
    except OSError as exc:
        cleanup_output_on_failure(output_path)
        raise JobProcessingError(failure_message, 500) from exc


def process_uploaded_document(input_path: Path, original_name: str, job_id: str, progress_callback=None, cover_info=None) -> dict:
    output_filename = f"{job_id}_output.docx"
    output_path = OUTPUT_FOLDER / output_filename
    release_output_reservation = reserve_output_storage(output_path)

    try:
        input_size_bytes = input_path.stat().st_size
        start_time = time.time()
        format_result = format_academic_paper(
            str(input_path),
            str(output_path),
            progress_callback=progress_callback,
            cover_info=cover_info,
            max_output_bytes=MAX_OUTPUT_FILE_BYTES,
        )
        elapsed = time.time() - start_time

        if not format_result:
            cleanup_output_on_failure(output_path)
            raise JobProcessingError("排版处理失败，请检查文档格式是否正确", 500)

        output_size_bytes = get_generated_output_size(
            output_path,
            "排版处理失败，未生成输出文件，请重试。",
        )
        return build_success_response(
            job_id=job_id,
            original_name=original_name,
            output_filename=output_filename,
            download_name=f"{original_name}_排版后.docx",
            input_size_text=f"{input_size_bytes / 1024:.1f} KB",
            output_size_text=f"{output_size_bytes / 1024:.1f} KB",
            elapsed=elapsed,
            format_result=format_result,
        )
    except FileNotFoundError as exc:
        cleanup_output_on_failure(output_path)
        raise JobProcessingError("上传文件已失效，请重新提交。", 500) from exc
    except OutputSizeLimitExceeded as exc:
        cleanup_output_on_failure(output_path)
        raise JobProcessingError(OUTPUT_TOO_LARGE_MESSAGE, 413) from exc
    except Exception:
        cleanup_output_on_failure(output_path)
        raise
    finally:
        release_output_reservation()


def process_text_document(text: str, job_id: str, progress_callback=None, cover_info=None) -> dict:
    output_filename = f"{job_id}_output.docx"
    output_path = OUTPUT_FOLDER / output_filename
    original_name = "黏贴文本排版"
    release_output_reservation = reserve_output_storage(output_path)

    try:
        start_time = time.time()
        format_result = format_academic_paper_from_text(
            text,
            str(output_path),
            progress_callback=progress_callback,
            cover_info=cover_info,
            max_output_bytes=MAX_OUTPUT_FILE_BYTES,
        )
        elapsed = time.time() - start_time

        if not format_result:
            cleanup_output_on_failure(output_path)
            raise JobProcessingError("排版处理失败，请检查文本内容", 500)

        output_size_bytes = get_generated_output_size(
            output_path,
            "排版处理失败，未生成输出文件，请重试。",
        )
        return build_success_response(
            job_id=job_id,
            original_name=original_name,
            output_filename=output_filename,
            download_name=f"{original_name}_结果.docx",
            input_size_text=f"{len(text.encode('utf-8')) / 1024:.1f} KB",
            output_size_text=f"{output_size_bytes / 1024:.1f} KB",
            elapsed=elapsed,
            format_result=format_result,
        )
    except OutputSizeLimitExceeded as exc:
        cleanup_output_on_failure(output_path)
        raise JobProcessingError(OUTPUT_TOO_LARGE_MESSAGE, 413) from exc
    except Exception:
        cleanup_output_on_failure(output_path)
        raise
    finally:
        release_output_reservation()


def process_merged_document(cover_path: Path, body_path: Path, body_name: str, job_id: str, progress_callback=None) -> dict:
    output_filename = f"{job_id}_output.docx"
    output_path = OUTPUT_FOLDER / output_filename
    release_output_reservation = reserve_output_storage(
        output_path,
        reserved_bytes=MAX_OUTPUT_FILE_BYTES * 2,
    )

    try:
        cover_size_bytes = cover_path.stat().st_size
        body_size_bytes = body_path.stat().st_size
        start_time = time.time()
        format_result = merge_cover_and_body(
            str(cover_path),
            str(body_path),
            str(output_path),
            progress_callback=progress_callback,
            max_output_bytes=MAX_OUTPUT_FILE_BYTES,
        )
        elapsed = time.time() - start_time

        if not format_result:
            cleanup_output_on_failure(output_path)
            raise JobProcessingError("合并排版失败，请检查文档格式是否正确", 500)

        output_size_bytes = get_generated_output_size(
            output_path,
            "合并排版失败，未生成输出文件，请重试。",
        )
        return build_success_response(
            job_id=job_id,
            original_name=body_name,
            output_filename=output_filename,
            download_name=f"{body_name}_合并排版后.docx",
            input_size_text=f"{(cover_size_bytes + body_size_bytes) / 1024:.1f} KB",
            output_size_text=f"{output_size_bytes / 1024:.1f} KB",
            elapsed=elapsed,
            format_result=format_result,
            include_format_result_alias=True,
        )
    except OutputSizeLimitExceeded as exc:
        cleanup_output_on_failure(output_path)
        raise JobProcessingError(OUTPUT_TOO_LARGE_MESSAGE, 413) from exc
    except Exception:
        cleanup_output_on_failure(output_path)
        raise
    finally:
        release_output_reservation()


def process_concatenated_document(first_path: Path, second_path: Path, output_name: str, job_id: str, progress_callback=None, restart_body_page_number: bool = True) -> dict:
    output_filename = f"{job_id}_output.docx"
    output_path = OUTPUT_FOLDER / output_filename
    release_output_reservation = reserve_output_storage(output_path)

    try:
        first_size_bytes = first_path.stat().st_size
        second_size_bytes = second_path.stat().st_size
        start_time = time.time()
        concat_result = concatenate_documents(
            str(first_path),
            str(second_path),
            str(output_path),
            progress_callback=progress_callback,
            restart_body_page_number=restart_body_page_number,
            max_output_bytes=MAX_OUTPUT_FILE_BYTES,
        )
        elapsed = time.time() - start_time

        if not concat_result:
            cleanup_output_on_failure(output_path)
            raise JobProcessingError("文档拼接失败，请检查两个文档是否均为有效的 .docx 文件", 500)

        output_size_bytes = get_generated_output_size(
            output_path,
            "文档拼接失败，未生成输出文件，请重试。",
        )
        return build_success_response(
            job_id=job_id,
            original_name=output_name,
            output_filename=output_filename,
            download_name=f"{output_name}_拼接后.docx",
            input_size_text=f"{(first_size_bytes + second_size_bytes) / 1024:.1f} KB",
            output_size_text=f"{output_size_bytes / 1024:.1f} KB",
            elapsed=elapsed,
            format_result=concat_result,
            preview=build_concat_preview(concat_result),
        )
    except DocumentConcatError as exc:
        cleanup_output_on_failure(output_path)
        raise JobProcessingError(str(exc), 400) from exc
    except OutputSizeLimitExceeded as exc:
        cleanup_output_on_failure(output_path)
        raise JobProcessingError(OUTPUT_TOO_LARGE_MESSAGE, 413) from exc
    except Exception:
        cleanup_output_on_failure(output_path)
        raise
    finally:
        release_output_reservation()


# ============================================================
# 路由
# ============================================================
@app.after_request
def add_response_security_headers(response):
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault("X-Frame-Options", "DENY")
    # Outputs and API responses are intended for this origin only.  Explicitly
    # opt into CORP so a cross-origin page cannot embed a generated document or
    # response as a readable resource through a permissive browser context.
    response.headers.setdefault("Cross-Origin-Resource-Policy", "same-origin")
    response.headers.setdefault("Permissions-Policy", PERMISSIONS_POLICY)
    response.headers.setdefault("Content-Security-Policy", CONTENT_SECURITY_POLICY)

    if is_api_request():
        response.headers["X-Request-ID"] = get_request_id()
        response.headers["Cache-Control"] = "no-store"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"

    if (
        request_may_have_unread_body()
        and not getattr(g, REQUEST_BODY_CONSUMED_ATTR, False)
    ):
        force_close_unread_request_connection(response)

    return response


@app.route("/")
def index():
    """提供主页"""
    return send_from_directory(app.static_folder, "index.html")


@app.route("/api/health", methods=["GET"])
def api_health():
    """健康检查接口，方便部署后探活与诊断。"""
    payload = get_health_payload()
    response = jsonify(payload)
    # A draining worker is intentionally unavailable only until the platform
    # routes traffic to a replacement.  Keep health checks and API clients on
    # the same bounded retry contract as other drain rejections.
    if payload.get("status") == "draining":
        response.headers["Retry-After"] = str(SSE_RETRY_AFTER_SECONDS)
    return response, 200 if payload["success"] else 503


def get_single_uploaded_file(field_name: str, missing_message: str, duplicate_message: str):
    """返回指定 multipart 字段的唯一上传文件，拒绝缺失或重复字段。"""
    # 触发 request.files 解析后，即使解析器中途失败也可能
    # 已创建 multipart spool。先登记关闭责任，teardown 才不会
    # 因异常发生在 getlist() 返回之前而漏掉这些文件句柄。
    setattr(g, REQUEST_MULTIPART_PARSED_ATTR, True)
    try:
        files = request.files.getlist(field_name)
    except OSError as exc:
        raise JobProcessingError("服务器临时存储不可用，请稍后重试。", 503) from exc
    mark_request_body_consumed()
    if not files:
        return None, json_error(missing_message, 400)
    if len(files) > 1:
        return None, json_error(duplicate_message, 400)

    return files[0], None


@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(_error):
    """统一返回 JSON，避免前端把 413 误判成网络错误。"""
    if is_api_request():
        return json_error("文件大小超过 50MB 限制，请压缩后重试。", 413)

    return _error


@app.errorhandler(404)
def handle_not_found(_error):
    if is_api_request():
        return json_error("接口不存在，请确认请求地址是否正确。", 404)

    return _error


@app.errorhandler(405)
def handle_method_not_allowed(_error):
    if is_api_request():
        return json_error("请求方法不受支持，请检查接口调用方式。", 405)

    return _error


@app.errorhandler(Exception)
def handle_unexpected_error(error):
    if isinstance(error, JobProcessingError) and is_api_request():
        if error.status_code >= 500:
            response = async_creation_failure_response(error.message, error.status_code)
            if response is not None:
                return response
        return json_error(error.message, error.status_code)

    if isinstance(error, HTTPException):
        if is_api_request():
            message = get_http_exception_message(error)
            status_code = error.code or 500
            if status_code >= 500:
                response = async_creation_failure_response(message, status_code)
                if response is not None:
                    return response
            return json_error(message, status_code)
        return error

    log_unexpected_exception("未处理异常", error)

    if is_api_request():
        response = async_creation_failure_response("服务器内部错误，请稍后重试。", 500)
        if response is not None:
            return response
        return json_error("服务器内部错误，请稍后重试。", 500)

    raise error


@app.route("/api/format", methods=["POST"])
def api_format():
    """
    接收上传的 .docx 文件，进行排版处理，返回处理结果。
    """
    ensure_storage_ready(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_jobs()

    file, error = get_single_uploaded_file("file", "未检测到上传文件", "一次只能上传一个待排版文件")
    if error is not None:
        return error

    if file.filename == "":
        return json_error("未选择文件", 400)

    if not allowed_file(file.filename):
        return json_error("仅支持 .docx 格式的文件", 400)

    if not is_valid_docx_upload(file):
        return json_error(INVALID_DOCX_MESSAGE.format(label="上传的文件"), 400)

    input_path = None
    try:
        job_id = create_unique_job_id()
        original_name = get_display_name(file.filename)
        input_path = UPLOAD_FOLDER / f"{job_id}_input.docx"
        cover_info = extract_cover_info(request.form)

        file_size = save_uploaded_file(file, input_path)
        release_request_upload_reservation()
        logger.info(f"收到文件: {get_log_filename(file.filename)} ({file_size / 1024:.1f} KB)")

        with processing_job_slot():
            response_data = process_uploaded_document(
                input_path,
                original_name,
                job_id,
                cover_info=cover_info,
            )
        return jsonify(response_data)

    except JobProcessingError as exc:
        return json_error(exc.message, exc.status_code)
    except Exception as e:
        log_unexpected_exception("处理过程中出现异常", e, input_path)
        return json_error("服务器内部错误，请稍后重试。", 500)
    finally:
        if input_path is not None:
            cleanup_path(input_path)


@app.route("/api/format_text", methods=["POST"])
def api_format_text():
    """
    接收纯文本内容，进行排版处理并生成 .docx，返回下载链接。
    """
    ensure_storage_ready(OUTPUT_FOLDER)
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_jobs()

    try:
        data = get_request_payload(
            max_content_length=MAX_TEXT_REQUEST_BYTES,
            too_large_message=TEXT_REQUEST_TOO_LARGE_MESSAGE,
        )
        text = get_payload_text(data)
        cover_info = extract_cover_info(data)

        if not text:
            return json_error("请输入有效的文字内容", 400)
        if is_text_payload_too_large(text):
            return json_error(TEXT_INPUT_TOO_LARGE_MESSAGE, 413)
        if is_text_paragraph_count_too_large(text):
            return json_error(TEXT_PARAGRAPH_LIMIT_MESSAGE, 413)

        job_id = create_unique_job_id()
        with processing_job_slot():
            response_data = process_text_document(text, job_id, cover_info=cover_info)
        return jsonify(response_data)
    except JobProcessingError as exc:
        return json_error(exc.message, exc.status_code)
    except Exception as e:
        log_unexpected_exception("文本处理异常", e)
        return json_error("服务器内部错误，请稍后重试。", 500)


@app.route("/api/format_merge", methods=["POST"])
def api_format_merge():
    """接收封面文档和正文文档，排版正文后合并为一个文档。"""
    ensure_storage_ready(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_jobs()

    if not is_docxcompose_supported():
        return json_error("当前服务未安装封面合并组件，请改用“自动生成模板封面”或补装 docxcompose。", 503)

    cover_file, error = get_single_uploaded_file("cover", "未检测到封面文档", "封面文档只能上传一个")
    if error is not None:
        return error
    body_file, error = get_single_uploaded_file("body", "未检测到正文文档", "正文文档只能上传一个")
    if error is not None:
        return error

    if cover_file.filename == "":
        return json_error("未选择封面文档", 400)
    if body_file.filename == "":
        return json_error("未选择正文文档", 400)

    if not allowed_file(cover_file.filename):
        return json_error("封面文档仅支持 .docx 格式", 400)
    if not allowed_file(body_file.filename):
        return json_error("正文文档仅支持 .docx 格式", 400)

    if not is_valid_docx_upload(cover_file):
        return json_error(INVALID_DOCX_MESSAGE.format(label="封面文档"), 400)
    if not is_valid_docx_upload(body_file):
        return json_error(INVALID_DOCX_MESSAGE.format(label="正文文档"), 400)

    cover_path = None
    body_path = None
    try:
        job_id = create_unique_job_id()
        body_name = get_display_name(body_file.filename)

        cover_path = UPLOAD_FOLDER / f"{job_id}_cover.docx"
        body_path = UPLOAD_FOLDER / f"{job_id}_body.docx"

        cover_size = save_uploaded_file(cover_file, cover_path)
        body_size = save_uploaded_file(body_file, body_path)
        release_request_upload_reservation()
        logger.info(
            f"收到合并请求: 封面={get_log_filename(cover_file.filename)} ({cover_size / 1024:.1f} KB), "
            f"正文={get_log_filename(body_file.filename)} ({body_size / 1024:.1f} KB)"
        )

        with processing_job_slot():
            response_data = process_merged_document(cover_path, body_path, body_name, job_id)
        return jsonify(response_data)

    except JobProcessingError as exc:
        return json_error(exc.message, exc.status_code)
    except Exception as e:
        log_unexpected_exception("合并处理过程中出现异常", e, cover_path, body_path)
        return json_error("服务器内部错误，请稍后重试。", 500)
    finally:
        for path in (cover_path, body_path):
            if path is not None:
                cleanup_path(path)


@app.route("/api/format_async", methods=["POST"])
def api_format_async():
    """创建正文排版异步任务，并通过 SSE 推送真实进度。"""
    ensure_storage_ready(UPLOAD_FOLDER, OUTPUT_FOLDER)
    reused_job = get_reused_request_progress_job("format")
    if reused_job is not None:
        return build_async_job_http_response(reused_job)
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_jobs()

    file, error = get_single_uploaded_file("file", "未检测到上传文件", "一次只能上传一个待排版文件")
    if error is not None:
        return error

    if file.filename == "":
        return json_error("未选择文件", 400)

    if not allowed_file(file.filename):
        return json_error("仅支持 .docx 格式的文件", 400)

    if not is_valid_docx_upload(file):
        return json_error(INVALID_DOCX_MESSAGE.format(label="上传的文件"), 400)

    job, created = get_or_create_request_progress_job("format")
    if not created:
        return build_async_job_http_response(job)
    input_path = UPLOAD_FOLDER / f"{job['id']}_input.docx"
    output_path = OUTPUT_FOLDER / f"{job['id']}_output.docx"
    original_name = get_display_name(file.filename)
    cover_info = extract_cover_info(request.form)

    try:
        file_size = save_uploaded_file(file, input_path)
        release_request_upload_reservation()
        logger.info(f"收到异步排版请求: {get_log_filename(file.filename)} ({file_size / 1024:.1f} KB)")

        emit_job_progress(job, 1, "文件上传完成，正在准备排版", f"{file_size / 1024:.1f} KB")
        response = build_async_job_http_response(job)
        launch_background_job(
            job,
            "异步正文排版",
            lambda progress_callback: process_uploaded_document(
                input_path,
                original_name,
                job["id"],
                progress_callback=progress_callback,
                cover_info=cover_info,
            ),
            cleanup_paths=(input_path,),
            log_paths=(output_path,),
        )
        return response
    except Exception as exc:
        handle_progress_job_creation_failure(job, exc)
        cleanup_path(input_path)
        raise


@app.route("/api/format_text_async", methods=["POST"])
def api_format_text_async():
    """创建黏贴文本排版异步任务。"""
    ensure_storage_ready(OUTPUT_FOLDER)
    reused_job = get_reused_request_progress_job("format_text")
    if reused_job is not None:
        return build_async_job_http_response(reused_job)
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_jobs()

    data = get_request_payload(
        max_content_length=MAX_TEXT_REQUEST_BYTES,
        too_large_message=TEXT_REQUEST_TOO_LARGE_MESSAGE,
    )
    text = get_payload_text(data)
    cover_info = extract_cover_info(data)

    if not text:
        return json_error("请输入有效的文字内容", 400)
    if is_text_payload_too_large(text):
        return json_error(TEXT_INPUT_TOO_LARGE_MESSAGE, 413)
    if is_text_paragraph_count_too_large(text):
        return json_error(TEXT_PARAGRAPH_LIMIT_MESSAGE, 413)

    job, created = get_or_create_request_progress_job("format_text")
    if not created:
        return build_async_job_http_response(job)
    output_path = OUTPUT_FOLDER / f"{job['id']}_output.docx"
    try:
        emit_job_progress(job, 1, "文本内容已接收，正在创建排版任务")
        response = build_async_job_http_response(job)
        launch_background_job(
            job,
            "异步文本排版",
            lambda progress_callback: process_text_document(
                text,
                job["id"],
                progress_callback=progress_callback,
                cover_info=cover_info,
            ),
            log_paths=(output_path,),
        )
        return response
    except Exception as exc:
        handle_progress_job_creation_failure(job, exc)
        cleanup_path(output_path)
        raise


@app.route("/api/format_merge_async", methods=["POST"])
def api_format_merge_async():
    """创建封面 + 正文合并排版异步任务。"""
    ensure_storage_ready(UPLOAD_FOLDER, OUTPUT_FOLDER)
    reused_job = get_reused_request_progress_job("format_merge")
    if reused_job is not None:
        return build_async_job_http_response(reused_job)
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_jobs()

    if not is_docxcompose_supported():
        return json_error("当前服务未安装封面合并组件，请改用“自动生成模板封面”或补装 docxcompose。", 503)

    cover_file, error = get_single_uploaded_file("cover", "未检测到封面文档", "封面文档只能上传一个")
    if error is not None:
        return error
    body_file, error = get_single_uploaded_file("body", "未检测到正文文档", "正文文档只能上传一个")
    if error is not None:
        return error

    if cover_file.filename == "":
        return json_error("未选择封面文档", 400)
    if body_file.filename == "":
        return json_error("未选择正文文档", 400)

    if not allowed_file(cover_file.filename):
        return json_error("封面文档仅支持 .docx 格式", 400)
    if not allowed_file(body_file.filename):
        return json_error("正文文档仅支持 .docx 格式", 400)

    if not is_valid_docx_upload(cover_file):
        return json_error(INVALID_DOCX_MESSAGE.format(label="封面文档"), 400)
    if not is_valid_docx_upload(body_file):
        return json_error(INVALID_DOCX_MESSAGE.format(label="正文文档"), 400)

    job, created = get_or_create_request_progress_job("format_merge")
    if not created:
        return build_async_job_http_response(job)
    cover_path = UPLOAD_FOLDER / f"{job['id']}_cover.docx"
    body_path = UPLOAD_FOLDER / f"{job['id']}_body.docx"
    output_path = OUTPUT_FOLDER / f"{job['id']}_output.docx"
    body_name = get_display_name(body_file.filename)

    try:
        cover_size = save_uploaded_file(cover_file, cover_path)
        body_size = save_uploaded_file(body_file, body_path)
        release_request_upload_reservation()
        logger.info(
            f"收到异步合并请求: 封面={get_log_filename(cover_file.filename)} ({cover_size / 1024:.1f} KB), "
            f"正文={get_log_filename(body_file.filename)} ({body_size / 1024:.1f} KB)"
        )

        emit_job_progress(
            job,
            1,
            "封面与正文上传完成，正在准备合并排版",
            f"总大小 {(cover_size + body_size) / 1024:.1f} KB",
        )
        response = build_async_job_http_response(job)
        launch_background_job(
            job,
            "异步合并排版",
            lambda progress_callback: process_merged_document(
                cover_path,
                body_path,
                body_name,
                job["id"],
                progress_callback=progress_callback,
            ),
            cleanup_paths=(cover_path, body_path),
            log_paths=(output_path,),
        )
        return response
    except Exception as exc:
        handle_progress_job_creation_failure(job, exc)
        for path in (cover_path, body_path):
            cleanup_path(path)
        raise


def _validate_concat_files():
    """校验拼接接口的两个上传文件，通过则返回 (first_file, second_file)，否则返回错误响应。"""
    if not is_docxcompose_supported():
        return None, json_error("当前服务未安装文档拼接组件，请补装 docxcompose 后重试。", 503)

    first_file, error = get_single_uploaded_file("first", "未检测到第一个文档", "第一个文档只能上传一个")
    if error is not None:
        return None, error
    second_file, error = get_single_uploaded_file("second", "未检测到第二个文档", "第二个文档只能上传一个")
    if error is not None:
        return None, error

    if first_file.filename == "":
        return None, json_error("未选择第一个文档", 400)
    if second_file.filename == "":
        return None, json_error("未选择第二个文档", 400)

    if not allowed_file(first_file.filename):
        return None, json_error("第一个文档仅支持 .docx 格式", 400)
    if not allowed_file(second_file.filename):
        return None, json_error("第二个文档仅支持 .docx 格式", 400)

    if not is_valid_docx_upload(first_file):
        return None, json_error(INVALID_DOCX_MESSAGE.format(label="第一个文档"), 400)
    if not is_valid_docx_upload(second_file):
        return None, json_error(INVALID_DOCX_MESSAGE.format(label="第二个文档"), 400)

    return (first_file, second_file), None


@app.route("/api/concat", methods=["POST"])
def api_concat():
    """拼接两个文档：保持各自排版形态不变，合并为一个 .docx。"""
    ensure_storage_ready(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_jobs()

    files, error = _validate_concat_files()
    if error is not None:
        return error
    first_file, second_file = files

    first_path = None
    second_path = None
    try:
        job_id = create_unique_job_id()
        output_name = get_display_name(second_file.filename)
        restart_page_number = parse_bool_flag(request.form, "restart_page_number")

        first_path = UPLOAD_FOLDER / f"{job_id}_first.docx"
        second_path = UPLOAD_FOLDER / f"{job_id}_second.docx"

        first_size = save_uploaded_file(first_file, first_path)
        second_size = save_uploaded_file(second_file, second_path)
        release_request_upload_reservation()
        logger.info(
            f"收到拼接请求: 文档一={get_log_filename(first_file.filename)} ({first_size / 1024:.1f} KB), "
            f"文档二={get_log_filename(second_file.filename)} ({second_size / 1024:.1f} KB)"
        )

        with processing_job_slot():
            response_data = process_concatenated_document(
                first_path,
                second_path,
                output_name,
                job_id,
                restart_body_page_number=restart_page_number,
            )
        return jsonify(response_data)

    except JobProcessingError as exc:
        return json_error(exc.message, exc.status_code)
    except Exception as e:
        log_unexpected_exception("拼接处理过程中出现异常", e, first_path, second_path)
        return json_error("服务器内部错误，请稍后重试。", 500)
    finally:
        for path in (first_path, second_path):
            if path is not None:
                cleanup_path(path)


@app.route("/api/concat_async", methods=["POST"])
def api_concat_async():
    """创建文档拼接异步任务，并通过 SSE 推送真实进度。"""
    ensure_storage_ready(UPLOAD_FOLDER, OUTPUT_FOLDER)
    reused_job = get_reused_request_progress_job("concat")
    if reused_job is not None:
        return build_async_job_http_response(reused_job)
    cleanup_expired_files(UPLOAD_FOLDER, OUTPUT_FOLDER)
    cleanup_expired_jobs()

    files, error = _validate_concat_files()
    if error is not None:
        return error
    first_file, second_file = files

    job, created = get_or_create_request_progress_job("concat")
    if not created:
        return build_async_job_http_response(job)
    first_path = UPLOAD_FOLDER / f"{job['id']}_first.docx"
    second_path = UPLOAD_FOLDER / f"{job['id']}_second.docx"
    output_path = OUTPUT_FOLDER / f"{job['id']}_output.docx"
    output_name = get_display_name(second_file.filename)
    restart_page_number = parse_bool_flag(request.form, "restart_page_number")

    try:
        first_size = save_uploaded_file(first_file, first_path)
        second_size = save_uploaded_file(second_file, second_path)
        release_request_upload_reservation()
        logger.info(
            f"收到异步拼接请求: 文档一={get_log_filename(first_file.filename)} ({first_size / 1024:.1f} KB), "
            f"文档二={get_log_filename(second_file.filename)} ({second_size / 1024:.1f} KB)"
        )

        emit_job_progress(
            job,
            1,
            "两个文档上传完成，正在准备拼接",
            f"总大小 {(first_size + second_size) / 1024:.1f} KB",
        )
        response = build_async_job_http_response(job)
        launch_background_job(
            job,
            "异步文档拼接",
            lambda progress_callback: process_concatenated_document(
                first_path,
                second_path,
                output_name,
                job["id"],
                progress_callback=progress_callback,
                restart_body_page_number=restart_page_number,
            ),
            cleanup_paths=(first_path, second_path),
            log_paths=(output_path,),
        )
        return response
    except Exception as exc:
        handle_progress_job_creation_failure(job, exc)
        for path in (first_path, second_path):
            cleanup_path(path)
        raise


@app.route("/api/jobs/<job_id>/events", methods=["GET"])
def api_job_events(job_id):
    """输出指定任务的 Server-Sent Events 进度流。"""
    if not is_valid_job_id(job_id):
        return json_error("任务不存在或已过期", 404)
    if SHUTDOWN_EVENT.is_set():
        response, status_code = json_error(SERVICE_DRAINING_MESSAGE, 503)
        # A draining worker rejects new streams during deploy.  Tell browser
        # and generic API clients when it is reasonable to retry instead of
        # making them spin or surface a permanent-looking error.
        response.headers["Retry-After"] = str(SSE_RETRY_AFTER_SECONDS)
        return response, status_code

    cleanup_expired_jobs()
    job = get_progress_job(job_id)
    if job is None:
        return json_error("任务不存在或已过期", 404)

    last_event_id = request.headers.get("Last-Event-ID", "")
    with job["condition"]:
        sent_count = parse_last_event_id(last_event_id, len(job["events"]))

    release_sse_slot = try_acquire_sse_connection_slot()
    if release_sse_slot is None:
        response, status_code = json_error(SSE_CAPACITY_MESSAGE, 429)
        response.headers["Retry-After"] = str(SSE_RETRY_AFTER_SECONDS)
        return response, status_code

    # 排空可能恰好发生在入口检查与槽位取得之间。再次复核，避免在
    # shutdown 已开始后仍建立新的长连接并占用 worker 线程；释放动作幂等，
    # 因而不会与响应 close 回调产生重复归还。
    if SHUTDOWN_EVENT.is_set():
        release_sse_slot()
        response, status_code = json_error(SERVICE_DRAINING_MESSAGE, 503)
        response.headers["Retry-After"] = str(SSE_RETRY_AFTER_SECONDS)
        return response, status_code

    def generate():
        try:
            delivered_count = sent_count

            while True:
                keepalive = False
                with job["condition"]:
                    notified = job["condition"].wait_for(
                        lambda: len(job["events"]) > delivered_count
                        or job["status"] in TERMINAL_JOB_STATUSES
                        or SHUTDOWN_EVENT.is_set(),
                        timeout=15,
                    )
                    pending_events = job["events"][delivered_count:]
                    finished = job["status"] in TERMINAL_JOB_STATUSES
                    draining = SHUTDOWN_EVENT.is_set()
                    if not notified and not pending_events and not finished:
                        keepalive = True

                if keepalive:
                    yield ": keep-alive\n\n"
                    continue

                for event in pending_events:
                    delivered_count += 1
                    yield format_sse_event(event, delivered_count)

                if (finished or draining) and delivered_count >= len(job["events"]):
                    break
        finally:
            release_sse_slot()

    try:
        response = Response(
            generate(),
            mimetype="text/event-stream",
            headers={
                "Cache-Control": "no-store",
                "X-Accel-Buffering": "no",
            },
        )
        response.call_on_close(release_sse_slot)
        return response
    except Exception:
        release_sse_slot()
        raise


@app.route("/api/jobs/<job_id>/result", methods=["GET"])
def api_job_result(job_id):
    """读取异步任务的最终结果。"""
    if not is_valid_job_id(job_id):
        response, status_code = json_error("任务不存在或已过期", 404)
        response.headers["Cache-Control"] = "no-store"
        return response, status_code

    cleanup_expired_jobs()
    result_snapshot = get_progress_job_result_snapshot(
        job_id,
        mark_requested=request.method == "GET",
    )
    if result_snapshot is None:
        response, status_code = json_error("任务不存在或已过期", 404)
        response.headers["Cache-Control"] = "no-store"
        return response, status_code
    status, error_info, result_payload = result_snapshot

    if status in {"queued", "running"}:
        response = jsonify({"success": False, "status": "processing"})
        # 202 responses are intentionally cache-free and should tell generic
        # clients when to retry instead of forcing them to guess a polling
        # interval.  The browser client already uses bounded polling, while
        # this header keeps API consumers interoperable with the same contract.
        response.headers["Cache-Control"] = "no-store"
        response.headers["Retry-After"] = str(JOB_RESULT_RETRY_AFTER_SECONDS)
        return response, 202

    if status == "error":
        error_info = coerce_mapping(error_info)
        response = json_terminal_job_error(
            normalize_job_error_message(error_info.get("message", "处理失败")),
            normalize_http_status_code(error_info.get("status_code", 500)),
        )
        response[0].headers["Cache-Control"] = "no-store"
        return response

    result_valid, result_error, result_status = validate_result_output(result_payload, job_id=job_id)
    if not result_valid:
        response = json_terminal_job_error(result_error, result_status)
        response[0].headers["Cache-Control"] = "no-store"
        return response

    output_filename = get_result_output_filename(result_payload)
    response = jsonify(sanitize_result_payload(result_payload, output_filename))
    response.headers["Cache-Control"] = "no-store"
    return response


@app.route("/api/download/<filename>")
def api_download(filename):
    """提供排版后文件的下载，并安全支持单段字节范围请求。"""
    cleanup_expired_files(OUTPUT_FOLDER)

    if not is_generated_output(filename):
        return json_error("文件不存在或已过期", 404)

    file_path = OUTPUT_FOLDER / filename
    lease = acquire_output_download_lease(file_path)
    if lease is None:
        return json_error("文件不存在或已过期", 404)
    release_download, download_file, file_size = lease

    response = None
    try:
        # Keep every operation after acquiring the lease inside the cleanup
        # guard.  Query-string parsing and filename normalization are request
        # controlled and may raise; leaking the lease here would keep the
        # descriptor open and prevent TTL/capacity cleanup from reclaiming the
        # output.
        download_name = get_download_name(request.args.get("name", ""), filename)
        response = send_file(
            download_file,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            conditional=False,
            etag=False,
        )
        response.content_length = file_size
        # Gunicorn 的 wsgi.file_wrapper 不保证 seekable；统一替换为 Werkzeug
        # 的可 seek 包装，避免大偏移 Range 为跳过前缀而顺序读取整份文件。
        response.response = FileWrapper(download_file)
        # send_file 默认的 direct_passthrough 会绕过 Response.close 回调；关闭
        # 直通后仍按迭代器流式发送，并由 ClosingIterator 可靠释放下载租约。
        response.direct_passthrough = False
        response.call_on_close(release_download)
        # 不暴露可复用验证器：输出位于临时存储，下载始终以当前文件内容为准。
        response.headers.pop("Last-Modified", None)
        # RFC 9110 §14.2：Range 只适用于 GET；HEAD 必须返回完整表示的
        # 元数据，不能因客户端携带无效范围而变成 416。
        response.headers["Accept-Ranges"] = "bytes"
        response.make_conditional(
            request.environ,
            accept_ranges=request.method == "GET",
            complete_length=file_size,
        )
        response.response = OutputDownloadLeaseIterable(
            response.response,
            release_download,
        )
        return response
    except RequestedRangeNotSatisfiable:
        close_output_download_response(response, release_download)
        error_response, status_code = json_error("请求的下载范围无效。", 416)
        error_response.headers["Accept-Ranges"] = "bytes"
        error_response.headers["Content-Range"] = f"bytes */{file_size}"
        return error_response, status_code
    except Exception:
        close_output_download_response(response, release_download)
        raise


# ============================================================
# 启动
# ============================================================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001)
