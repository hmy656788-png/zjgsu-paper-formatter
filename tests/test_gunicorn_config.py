import socket
import time
import unittest

from gunicorn.config import Config

from gunicorn_config import HeaderDeadlineRequestParser, NoMoreData


class GunicornBodySocketTimeoutTests(unittest.TestCase):
    def test_keepalive_drain_is_bounded_when_previous_body_is_unread(self):
        server_socket, client_socket = socket.socketpair()
        try:
            parser = HeaderDeadlineRequestParser(
                Config(),
                server_socket,
                ("127.0.0.1", 12345),
                body_timeout_seconds=0.05,
            )
            client_socket.sendall(
                b"POST / HTTP/1.1\r\nHost: localhost\r\n"
                b"Content-Length: 4\r\n\r\n"
            )
            next(parser)

            started = time.monotonic()
            with self.assertRaises(NoMoreData):
                next(parser)
            elapsed = time.monotonic() - started
            self.assertLess(elapsed, 0.5)
            self.assertIsNone(server_socket.gettimeout())
        finally:
            client_socket.close()
            server_socket.close()

    def test_completed_upload_restores_response_socket_timeout(self):
        for original_timeout in (None, 2.0):
            with self.subTest(original_timeout=original_timeout):
                server_socket, client_socket = socket.socketpair()
                try:
                    parser = HeaderDeadlineRequestParser(
                        Config(),
                        server_socket,
                        ("127.0.0.1", 12345),
                        body_timeout_seconds=1,
                    )
                    client_socket.sendall(
                        b"POST / HTTP/1.1\r\nHost: localhost\r\n"
                        b"Content-Length: 4\r\n\r\n"
                    )
                    request = next(parser)
                    server_socket.settimeout(original_timeout)
                    # Send the body after parsing the headers so the body
                    # actually reads the socket instead of a buffered chunk.
                    client_socket.sendall(b"body")

                    self.assertEqual(request.body.read(), b"body")
                    self.assertIsNone(parser.unreader._body_deadline)
                    self.assertEqual(server_socket.gettimeout(), original_timeout)
                finally:
                    client_socket.close()
                    server_socket.close()

    def test_unread_body_honors_shorter_gunicorn_drain_deadline(self):
        server_socket, client_socket = socket.socketpair()
        try:
            parser = HeaderDeadlineRequestParser(
                Config(),
                server_socket,
                ("127.0.0.1", 12345),
                body_timeout_seconds=2,
            )
            client_socket.sendall(
                b"POST / HTTP/1.1\r\nHost: localhost\r\n"
                b"Content-Length: 4\r\n\r\n"
            )
            next(parser)

            started = time.monotonic()
            self.assertFalse(parser.finish_body(deadline=started + 0.05))
            self.assertLess(time.monotonic() - started, 0.5)
            self.assertIsNone(server_socket.gettimeout())
        finally:
            client_socket.close()
            server_socket.close()


if __name__ == "__main__":
    unittest.main()
