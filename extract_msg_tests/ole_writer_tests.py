import unittest

import extract_msg

from .constants import TEST_FILE_DIR, USER_TEST_DIR


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
