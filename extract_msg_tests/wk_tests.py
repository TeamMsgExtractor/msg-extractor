from __future__ import annotations

__all__ = [
    'WkTests'
]

import http.server
import io
import socketserver
import threading
import unittest

import pypdf

#from pathlib import Path

from .constants import WK_PATH
from extract_msg.msg_classes.message_base import MessageBase


class LoopbackHandler(http.server.BaseHTTPRequestHandler):
    requests_seen: list[str] = []

    def do_GET(self):
        self.__class__.requests_seen.append(self.path)
        body = b"LOCAL_TEST_SECRET"
        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *_args, **_kwargs):
        pass



class ProbeMessage(MessageBase):
    def getSaveHtmlBody(self, *_args, **_kwargs):
        return self.html



class WkTests(unittest.TestCase):
    @unittest.skipIf(WK_PATH is None, 'Wkhtmltopdf not found.')
    def testWkSecurity(self):
        """
        Tests the security around calling wkhtmltopdf.

        Test for Advisory GHSA-cxpg-prmv-g4pp.
        """
        # Thank you to Kadhir (https://github.com/JOKER-53) for nearly all of
        # this test code.

        # Setup the loopback server first.
        with socketserver.TCPServer(("127.0.0.1", 0), LoopbackHandler) as server:
            threading.Thread(target=server.serve_forever, daemon=True).start()
            port = server.server_address[1]

            vectors = [
                (
                    "iframe",
                    (
                        f'<html><body>'
                        f'<iframe src="http://127.0.0.1:{port}/ssrf-iframe"'
                        f' width="600" height="200"></iframe>'
                        f'</body></html>'
                    ).encode(),
                ),
                (
                    "img",
                    (
                        f'<html><body>'
                        f'<img src="http://127.0.0.1:{port}/ssrf-img">'
                        f'</body></html>'
                    ).encode(),
                ),
                (
                    "link-css",
                    (
                        f'<html><head>'
                        f'<link rel="stylesheet" href="http://127.0.0.1:{port}/ssrf-css">'
                        f'</head><body>test</body></html>'
                    ).encode(),
                ),
                (
                    "css-style",
                    (
                        f'<html><head>'
                        f'<style>body {{background-image: url(http://127.0.0.1:{port}/ssrf-css-style)}}</style>'
                        f'</head><body>test</body></html>'
                    ).encode(),
                ),
                (
                    "css-tag",
                    (
                        f'<html><head></head>'
                        f'<body style="background-image: url(http://127.0.0.1:{port}/ssrf-css-tag)">test</body></html>'
                    ).encode(),
                ),
            ]

            # We want to run all of the tests really quickly and then evaluate
            # them later.
            results = []
            for label, html in vectors:
                r = run_vector(label, html, port, LoopbackHandler)
                results.append(r)

            server.shutdown()

        for result in results:
            with self.subTest(f'{result["label"]} request seen'):
                self.assertListEqual(result['requests_seen'], [])
            with self.subTest(f'{result["label"]} secret injected'):
                self.assertFalse(result['secret_in_pdf'])


# For testWkSecurity
def run_vector(label: str, html: bytes, port: int, handler_cls) -> dict:
    """Run a single vector, return result dict."""
    handler_cls.requests_seen = []
    probe = object.__new__(ProbeMessage)
    probe.html = html
    try:
        pdf = probe.getSavePdfBody(WK_PATH)
    except Exception as e:
        return {
            "label": label,
            "requests_seen": [],
            "secret_in_pdf": False,
            "pdf_bytes": 0,
            "error": str(e),
            "result": "ERROR",
        }

    # No reason to output it to a file, we should just pass it directly to pypdf
    # inside of an io.
    #output = Path(f"ssrf-proof-{label.lower().replace(' ', '-')}.pdf")
    #output.write_bytes(pdf)

    try:
        #pdfFile = pypdf.PdfReader(output)
        pdfFile = pypdf.PdfReader(io.BytesIO(pdf))
        extracted = ''.join(x.extract_text() for x in pdfFile.pages)
        secret_in_pdf = b"LOCAL_TEST_SECRET" in extracted.encode()
    except Exception:
        secret_in_pdf = False

    requests_hit = list(handler_cls.requests_seen)
    vuln_confirmed = bool(requests_hit) or secret_in_pdf

    return {
        "label": label,
        "requests_seen": requests_hit,
        "secret_in_pdf": secret_in_pdf,
        "pdf_bytes": len(pdf),
        #"pdf_path": str(output.resolve()),
        "error": None,
        # From DEFENDER perspective: PASS = blocked, FAIL = vulnerable
        "result": "FAIL (VULNERABLE)" if vuln_confirmed else "PASS (BLOCKED)",
    }


