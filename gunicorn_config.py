"""Gunicorn request-boundary and lifecycle hardening."""

import os
import signal
import socket
import time
from contextlib import contextmanager, nullcontext, suppress

from gunicorn import http as gunicorn_http
from gunicorn.http.errors import NoMoreData
from gunicorn.http.parser import RequestParser
from gunicorn.http.unreader import SocketUnreader


DEFAULT_REQUEST_HEADER_TIMEOUT_SECONDS = 10
MIN_REQUEST_HEADER_TIMEOUT_SECONDS = 1
MAX_REQUEST_HEADER_TIMEOUT_SECONDS = 60
DEFAULT_REQUEST_BODY_TIMEOUT_SECONDS = 300
MIN_REQUEST_BODY_TIMEOUT_SECONDS = 1
MAX_REQUEST_BODY_TIMEOUT_SECONDS = 600


def parse_request_header_timeout(value) -> int:
    """解析请求头总时限；无效或越界配置回退到安全默认值。"""
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        return DEFAULT_REQUEST_HEADER_TIMEOUT_SECONDS
    if not MIN_REQUEST_HEADER_TIMEOUT_SECONDS <= parsed <= MAX_REQUEST_HEADER_TIMEOUT_SECONDS:
        return DEFAULT_REQUEST_HEADER_TIMEOUT_SECONDS
    return parsed


REQUEST_HEADER_TIMEOUT_SECONDS = parse_request_header_timeout(
    os.environ.get("REQUEST_HEADER_TIMEOUT_SECONDS"),
)


def parse_request_body_timeout(value) -> int:
    """解析请求正文总时限；无效或越界配置回退到安全默认值。"""
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        return DEFAULT_REQUEST_BODY_TIMEOUT_SECONDS
    if not MIN_REQUEST_BODY_TIMEOUT_SECONDS <= parsed <= MAX_REQUEST_BODY_TIMEOUT_SECONDS:
        return DEFAULT_REQUEST_BODY_TIMEOUT_SECONDS
    return parsed


REQUEST_BODY_TIMEOUT_SECONDS = parse_request_body_timeout(
    os.environ.get("REQUEST_BODY_TIMEOUT_SECONDS"),
)
_ORIGINAL_GET_PARSER = gunicorn_http.get_parser


class HeaderDeadlineSocketUnreader(SocketUnreader):
    """仅在 HTTP/1 请求头解析期间强制不可被 trickle 重置的总时限。"""

    def __init__(
        self,
        sock,
        timeout_seconds: float,
        body_timeout_seconds: float = REQUEST_BODY_TIMEOUT_SECONDS,
    ):
        super().__init__(sock)
        self.timeout_seconds = max(0.001, float(timeout_seconds))
        self.body_timeout_seconds = max(0.001, float(body_timeout_seconds))
        self._header_deadline = None
        self._body_deadline = None

    def begin_body_deadline(self):
        """从请求头解析完成起，给本请求正文设置不可被 trickle 重置的总时限。"""
        self._body_deadline = time.monotonic() + self.body_timeout_seconds

    def clear_body_deadline(self):
        """正文完全消费后撤销 deadline，避免影响后续业务处理时间。"""
        self._body_deadline = None

    @contextmanager
    def enforce_header_deadline(self):
        prior_timeout = self.sock.gettimeout()
        self._header_deadline = time.monotonic() + self.timeout_seconds
        try:
            yield
        finally:
            self._header_deadline = None
            with suppress(OSError):
                self.sock.settimeout(prior_timeout)

    def chunk(self):
        deadline = self._header_deadline
        expiry = self.expire_incomplete_header
        if deadline is None:
            deadline = self._body_deadline
            expiry = self.expire_incomplete_body
        if deadline is None:
            return super().chunk()

        remaining = deadline - time.monotonic()
        if remaining <= 0:
            expiry()
        prior_timeout = self.sock.gettimeout()
        # Gunicorn may be draining an unread body with its own shorter
        # timeout.  Preserve that bound and restore the socket after every
        # read, so an upload deadline cannot leak into response writes.
        read_timeout = remaining if prior_timeout is None else min(remaining, prior_timeout)
        self.sock.settimeout(read_timeout)
        try:
            return super().chunk()
        except (socket.timeout, TimeoutError) as exc:
            if read_timeout < remaining:
                raise
            expiry(exc)
        finally:
            with suppress(OSError):
                self.sock.settimeout(prior_timeout)

    def expire_incomplete_header(self, cause=None):
        # Gunicorn 的主线程默认会为“优雅关闭”逐个排空未读数据，
        # 每个连接最多 2 秒。不完整请求头尚未产生任何响应，可直接双向
        # shutdown，避免攻击者把 N 个超时连接变成 N × 2 秒的串行阻塞。
        with suppress(OSError):
            self.sock.shutdown(socket.SHUT_RDWR)
        error = NoMoreData("request header deadline exceeded")
        if cause is None:
            raise error
        raise error from cause

    def expire_incomplete_body(self, cause=None):
        """正文总时限到期时断开连接，防止处理槽被慢速上传永久占用。"""
        with suppress(OSError):
            self.sock.shutdown(socket.SHUT_RDWR)
        # Gunicorn may still run its connection cleanup after the parser
        # raises.  Do not leak the temporary deadline into that path (or a
        # later keep-alive request); restore the normal blocking socket mode.
        with suppress(OSError):
            self.sock.settimeout(None)
        error = NoMoreData("request body deadline exceeded")
        if cause is None:
            raise error
        raise error from cause


class BodyDeadlineProxy:
    """在 Gunicorn Body 完整消费后清除 socket 正文 deadline。"""

    def __init__(self, body, clear_deadline):
        self._body = body
        self._clear_deadline = clear_deadline

    def _clear_if_finished(self):
        reader = getattr(self._body, "reader", None)
        if (
            getattr(reader, "length", None) == 0
            or getattr(reader, "finished", False)
            or getattr(reader, "parser", object()) is None
        ):
            self._clear_deadline()

    def read(self, size=None):
        value = self._body.read(size)
        self._clear_if_finished()
        return value

    def readline(self, size=None):
        value = self._body.readline(size)
        self._clear_if_finished()
        return value

    def readlines(self, size=None):
        value = self._body.readlines(size)
        self._clear_if_finished()
        return value

    def __iter__(self):
        return self

    def __next__(self):
        value = next(self._body)
        self._clear_if_finished()
        return value

    next = __next__

    def __getattr__(self, name):
        return getattr(self._body, name)


class HeaderDeadlineRequestParser(RequestParser):
    """Gunicorn HTTP/1 parser with a total deadline around each header block."""

    def __init__(
        self,
        cfg,
        source,
        source_addr,
        *,
        header_timeout_seconds: float = REQUEST_HEADER_TIMEOUT_SECONDS,
        body_timeout_seconds: float = REQUEST_BODY_TIMEOUT_SECONDS,
    ):
        super().__init__(cfg, source, source_addr)
        if hasattr(source, "recv"):
            self.unreader = HeaderDeadlineSocketUnreader(
                source,
                header_timeout_seconds,
                body_timeout_seconds,
            )

    def __next__(self):
        # 与锁定的 Gunicorn 26 Parser.__next__ 保持同样的请求边界，
        # 但只把 Message 构造（请求行 + headers）放进 deadline。
        # 前一请求的 body 排空明确留在时限之外。
        if self.mesg and self.mesg.should_close():
            raise StopIteration()
        # Never perform Gunicorn's default unbounded keep-alive drain.  A
        # client can leave a declared body unread and trickle bytes forever;
        # draining it before parsing the next request would pin this worker
        # thread indefinitely.  Reuse the same absolute body budget as an
        # upload and close the connection when the drain cannot complete.
        if self.mesg:
            body_timeout = getattr(
                self.unreader,
                "body_timeout_seconds",
                REQUEST_BODY_TIMEOUT_SECONDS,
            )
            drained = self.finish_body(
                deadline=time.monotonic() + body_timeout,
            )
            if not drained:
                expire_body = getattr(self.unreader, "expire_incomplete_body", None)
                if callable(expire_body):
                    expire_body()
                raise StopIteration()
        self.req_count += 1

        deadline_context = getattr(
            self.unreader,
            "enforce_header_deadline",
            nullcontext,
        )
        with deadline_context():
            message = self.mesg_class(
                self.cfg,
                self.unreader,
                self.source_addr,
                self.req_count,
            )
        if not message:
            raise StopIteration()
        begin_body_deadline = getattr(self.unreader, "begin_body_deadline", None)
        clear_body_deadline = getattr(self.unreader, "clear_body_deadline", None)
        if callable(begin_body_deadline):
            begin_body_deadline()
        if not callable(clear_body_deadline):
            clear_body_deadline = lambda: None
        message.body = BodyDeadlineProxy(message.body, clear_body_deadline)
        self.mesg = message
        return self.mesg

    def finish_body(self, deadline=None, max_bytes=None):
        drained = super().finish_body(deadline=deadline, max_bytes=max_bytes)
        if drained:
            clear_body_deadline = getattr(self.unreader, "clear_body_deadline", None)
            if callable(clear_body_deadline):
                clear_body_deadline()
        return drained

    next = __next__


def get_parser_with_header_deadline(
    cfg,
    source,
    source_addr,
    http2_connection=False,
):
    """只替换 HTTP/1 请求解析器；HTTP/2 与非 HTTP 协议继续由 Gunicorn 处理。"""
    if http2_connection or getattr(cfg, "protocol", "http") != "http":
        return _ORIGINAL_GET_PARSER(
            cfg,
            source,
            source_addr,
            http2_connection=http2_connection,
        )
    return HeaderDeadlineRequestParser(cfg, source, source_addr)


def post_worker_init(worker) -> None:
    """Install request-header deadline and drain before Gunicorn handles SIGTERM."""
    gunicorn_http.get_parser = get_parser_with_header_deadline

    from app import begin_graceful_shutdown

    def handle_term(signum, frame):
        begin_graceful_shutdown()
        worker.handle_exit(signum, frame)

    signal.signal(signal.SIGTERM, handle_term)
