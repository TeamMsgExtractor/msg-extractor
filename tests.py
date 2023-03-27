import shutil
import os
import unittest

import extract_msg

from pathlib import Path


TEST_FILE_DIR = Path(__file__).parent / 'example-msg-files'
TEST_EXPORT_DIR = TEST_FILE_DIR / 'export-results'


class OleWriterTests(unittest.TestCase):
    def testExportExamples(self):
        """
        Tests exporting the example files.
        """
        for exportResultFile in TEST_EXPORT_DIR.glob('*.msg'):
            # If we have an export result, find the original file, open it,
            # and export it as bytes to check against the known result.
            with extract_msg.openMsg(TEST_FILE_DIR / exportResultFile.name) as msg:
                exportedBytes = msg.exportBytes()

            with open(exportResultFile, 'rb') as f:
                exportResult = f.read()

            l1 = len(exportedBytes)
            l2 = len(exportResult)
            # We use three assertions to give better error messages.
            self.assertFalse(l1 > l2, 'Exported data is too large (expected {12}, got {l1}).')
            self.assertFalse(l1 < l2, 'Exported data is too small (expected {12}, got {l1}).')
            self.assertEqual(exportedBytes, exportResult, 'Exported data is incorrect.')


unittest.main(verbosity=2)
