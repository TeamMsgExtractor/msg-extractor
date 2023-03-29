"""
The tests for various parts of extract_msg. Tests can be run on user files by
defining the EXTRACT_MSG_TEST_DIR environment variable and setting it to the
folder with your own tests files. It must match the folder structure for the
specific test for that test to run.
"""

import shutil
import os
import unittest

import extract_msg

from pathlib import Path


TEST_FILE_DIR = Path(__file__).parent / 'example-msg-files'
USER_TEST_DIR = None
if bool(userTestDir := os.environ.get('EXTRACT_MSG_TEST_DIR')):
    userTestDir = Path(userTestDir)
    if userTestDir.exists():
        USER_TEST_DIR = userTestDir


class OleWriterTests(unittest.TestCase):
    def testExportExamples(self, testFileDir = TEST_FILE_DIR):
        """
        Tests exporting the example files.

        Uses the files in export-results to determine which test files to use.
        """
        for exportResultFile in (testFileDir / 'export-results').glob('*.msg'):
            # If we have an export result, find the original file, open it,
            # and export it as bytes to check against the known result.
            with extract_msg.openMsg(testFileDir / exportResultFile.name) as msg:
                exportedBytes = msg.exportBytes()

            with open(exportResultFile, 'rb') as f:
                exportResult = f.read()

            l1 = len(exportedBytes)
            l2 = len(exportResult)
            # We use three assertions to give better error messages.
            self.assertFalse(l1 > l2, 'Exported data is too large (expected {12}, got {l1}).')
            self.assertFalse(l1 < l2, 'Exported data is too small (expected {12}, got {l1}).')
            self.assertEqual(exportedBytes, exportResult, 'Exported data is incorrect.')

    @unittest.skipIf(USER_TEST_DIR is None, 'User test files not defined.')
    @unittest.skipIf(not (USER_TEST_DIR / 'export-results').exists(), 'User export tests not defined.')
    def testExtraExportExamples(self):
        """
        Tests exporting the user's example files.

        Uses the files in export-results to determine which test files to use.
        """
        self.testExportExamples(USER_TEST_DIR)


unittest.main(verbosity=2)
