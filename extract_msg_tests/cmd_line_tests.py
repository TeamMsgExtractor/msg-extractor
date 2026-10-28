from __future__ import annotations

__all__ = [
    'CommandLineTests',
]


import io
import subprocess
import sys
import unittest

from typing import Tuple

from .constants import TEST_FILE_DIR, USER_TEST_DIR
from extract_msg import __main__


class _StdRedirect:
    def __init__(self, data: bytes = None):
        if data:
            self.inBuffer = io.BytesIO(data)
        else:
            self.inBuffer = io.BytesIO()
        self.outBuffer = io.BytesIO()
        self.errBuffer = io.BytesIO()
        self.inReplacer = io.TextIOWrapper(self.inBuffer, sys.stdin.encoding, newline = '')
        self.outReplacer = io.TextIOWrapper(self.outBuffer, sys.stdout.encoding, newline = '', write_through = True)
        self.errReplacer = io.TextIOWrapper(self.errBuffer, sys.stderr.encoding, newline = '', write_through = True)
        self.stdin = None
        self.stdout = None
        self.stderr = None
        self.entered = False

    def bufferInput(self, data: bytes) -> _StdRedirect:
        """
        Set the stdin data to :param data:.
        """
        self.inBuffer.truncate(0)
        self.inBuffer.seek(0)
        self.inBuffer.write(data)
        self.inBuffer.seek(0)
        return self

    def __enter__(self, *args, **kwargs):
        if not self.entered:
            self.entered = True
            self.stdin = sys.stdin
            self.stdout = sys.stdout
            self.stderr = sys.stderr
            # I feel like this should be simple, but without doing both it tends
            # to just leave the data null in the buffer.
            self.inBuffer.seek(0)
            self.outBuffer.truncate(0)
            self.outBuffer.seek(0)
            self.errBuffer.truncate(0)
            self.errBuffer.seek(0)
            self.outReplacer.truncate(0)
            self.outReplacer.seek(0)
            self.errReplacer.truncate(0)
            self.errReplacer.seek(0)
            sys.stdin = self.inReplacer
            sys.stdout = self.outReplacer
            sys.stderr = self.errReplacer

    def __exit__(self, *args, **kwargs):
        if self.entered:
            self.entered = False
            sys.stdin = self.stdin
            self.stdin = None
            sys.stdout = self.stdout
            self.stdout = None
            sys.stderr = self.stderr
            self.stderr = None

    def dumpRead(self) -> Tuple[bytes, bytes]:
        """
        Dumps the output to stdout and stderr.
        """
        self.outBuffer.seek(0)
        self.errBuffer.seek(0)
        return self.outBuffer.read(), self.errBuffer.read()



class CommandLineTests(unittest.TestCase):
    def testStdin(self, testFileDir = TEST_FILE_DIR):
        for path in testFileDir.glob('*.msg'):
            # First, let's do the file on the disk.
            process = subprocess.Popen([sys.executable, '-m', 'extract_msg', '--dump-stdout', str(path)], stdin = subprocess.PIPE, stdout = subprocess.PIPE, stderr = subprocess.PIPE)
            # Wait for the process to return data.
            stdout1, stderr1 = process.communicate()

            # Now, do the same thing with stdin.
            process = subprocess.Popen([sys.executable, '-m', 'extract_msg', '-s', '--dump-stdout'], stdin = subprocess.PIPE, stdout = subprocess.PIPE, stderr = subprocess.PIPE)
            with open(path, 'rb') as f:
                stdout2, stderr2 = process.communicate(f.read())

            # Now, compare the two.
            with self.subTest(path):
                self.assertEqual(stdout1, stdout2)
                self.assertEqual(stderr1, stderr2)

    def testDirectRunStdin(self, testFileDir = TEST_FILE_DIR):
        """
        Attempt the stdin tests without calling a subprocess.
        """
        ioRedirect = _StdRedirect()
        for path in testFileDir.glob('*.msg'):
            with ioRedirect.bufferInput(b''):
                # First, let's do the file on the disk.
                __main__.main([__main__.__name__, '--dump-stdout', str(path)])

            stdout1, stderr1 = ioRedirect.dumpRead()

            with open(path, 'rb') as f:
                fileRead = f.read()

            with ioRedirect.bufferInput(fileRead):
                # Now, do the same thing with stdin.
                __main__.main([__main__.__name__, '-s', '--dump-stdout'])

            stdout2, stderr2 = ioRedirect.dumpRead()

            # Now, compare the two.
            with self.subTest(path):
                self.assertEqual(stdout1, stdout2)
                self.assertEqual(stderr1, stderr2)

    @unittest.skipIf(USER_TEST_DIR is None, 'User test files not defined.')
    def testUserStdin(self):
        self.testStdin(USER_TEST_DIR)