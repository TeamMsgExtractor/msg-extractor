import unittest

import extract_msg

from .constants import TEST_FILE_DIR, USER_TEST_DIR
from extract_msg.ole_writer import OleWriter


class OleWriterEditingTests(unittest.TestCase):
    def testAddEntryManual(self):
        writer = OleWriter()
        writer.addEntry('storage_1', storage = True)
        writer.addEntry('stream_1', b'Hello World')
        with self.assertRaises(OSError, msg = 'Cannot add an entry that already exists.'):
            # Try to add an entry that already exists.
            writer.addEntry('storage_1', b'')
        with self.assertRaises(OSError, msg = 'Attempted to access children of a stream.'):
            # Try to add an entry as a child of a stream.
            writer.addEntry('stream_1/child', b'')
        with self.assertRaises(ValueError, msg = 'Path segments must not be greater than 31 characters.'):
            # Try to use a path with a name that is too long.
            writer.addEntry('storage_1/12345678901234567890123456789012', b'')
        with self.assertRaises(ValueError, msg = 'Illegal character ("!" or ":") found in MSG path.'):
            # Attempt to use : in path.
            writer.addEntry('::InternalName', b'')

        with self.assertRaises(ValueError, msg = 'Illegal character ("!" or ":") found in MSG path.'):
            # Attempt to use ! in path.
            writer.addEntry('!bang', b'')



class OleWriterExportTests(unittest.TestCase):
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

            # We use two assertions to give better error messages.
            self.assertCountEqual(exportResult, exportedBytes, 'Exported data is wrong size.')
            self.assertEqual(exportedBytes, exportResult, 'Exported data is incorrect.')

    @unittest.skipIf(USER_TEST_DIR is None, 'User test files not defined.')
    @unittest.skipIf(not (USER_TEST_DIR / 'export-results').exists(), 'User export tests not defined.')
    def testExtraExportExamples(self):
        """
        Tests exporting the user's example files.

        Uses the files in export-results to determine which test files to use.
        """
        self.testExportExamples(USER_TEST_DIR)
