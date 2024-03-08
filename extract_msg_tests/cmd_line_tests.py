__all__ = [
    'CommandLineTests',
]


import pathlib
import subprocess
import sys
import unittest

from .constants import TEST_FILE_DIR, USER_TEST_DIR


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

    @unittest.skipIf(USER_TEST_DIR is None, 'User test files not defined.')
    def testUserStdin(self):
        self.testStdin(USER_TEST_DIR)