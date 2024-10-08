__all__ = [
    'OleWriterEditingTests',
    'OleWriterExportTests',
]


import unittest

import extract_msg

from .constants import TEST_FILE_DIR, USER_TEST_DIR
from extract_msg.ole_writer import OleWriter


class OleWriterEditingTests(unittest.TestCase):
    def _setupWriter(self) -> OleWriter:
        """
        Sets up the olewriter in a way that shouldn't throw any errors. Used for
        every test.
        """
        writer = OleWriter()
        writer.addEntry('storage_1', storage = True)
        writer.addEntry('stream_1', b'Hello World')
        writer.editEntry('storage_1',
                         clsid = b'0123456789012345',
                         creationTime = 1,
                         modifiedTime = 2,
                         stateBits = 3)

        return writer

    def testSetup(self):
        """
        Initial test to check that the setup does not throw exceptions.
        """
        self._setupWriter()
        raise Exception()
    
    def testAddEntryManual(self):
        """
        Tests the `addEntry` method.
        """
        writer = self._setupWriter()

        # Check that addEntry function in setup actually worked.
        self.assertEqual(writer.getEntry('stream_1').data, b'Hello World')
        writer.getEntry('storage_1')

        # Try to add an entry that already exists.
        with self.assertRaises(OSError, msg = 'Cannot add an entry that already exists.'):
            writer.addEntry('storage_1', b'')

        # Try to add an entry as a child of a stream.
        with self.assertRaises(OSError, msg = 'Attempted to access children of a stream.'):
            writer.addEntry('stream_1/child', b'')

        # Try to use a path with a name that is too long.
        with self.assertRaises(ValueError, msg = 'Path segments must not be greater than 31 characters.'):
            writer.addEntry('storage_1/12345678901234567890123456789012', b'')

        # Attempt to use ":" in path.
        with self.assertRaises(ValueError, msg = 'Illegal character ("!" or ":") found in MSG path.'):
            writer.addEntry('::InternalName', b'')

        # Attempt to use "!" in path.
        with self.assertRaises(ValueError, msg = 'Illegal character ("!" or ":") found in MSG path.'):
            writer.addEntry('!bang', b'')

    def testEditEntryManual(self):
        """
        Tests the `editEntry` method.
        """
        writer = self._setupWriter()

        # Check the edits made in self._setupWriter() This is the basic test to
        # make sure that the editEntry method works at all.
        testEntry = writer.getEntry('storage_1')
        self.assertEqual(testEntry.clsid, b'0123456789012345')
        self.assertEqual(testEntry.creationTime, 1)
        self.assertEqual(testEntry.modifiedTime, 2)
        self.assertEqual(testEntry.stateBits, 3)

        # If all of the previous ones worked, test editing the data of an entry.
        writer.editEntry('stream_1', data = b'New data')
        self.assertEqual(writer.getEntry('stream_1').data, b'New data')

        # Try editing the data of a storage.
        with self.assertRaises(TypeError, msg = 'Cannot set the data of a storage object.'):
            writer.editEntry('storage_1', data = b'h')

        # Try editing the data of a bad path.
        with self.assertRaises(ValueError, msg = 'Illegal character ("!" or ":") found in MSG path.'):
            writer.editEntry('::invalid', data = b'1')

        # Try editing the data of a missing file.
        with self.assertRaises(OSError, msg = f'Entry not found: not_found'):
            writer.editEntry('not_found', data = b'')

        # Try to use bad data for each of the parameters.
        with self.assertRaises(ValueError, msg = 'Data must be a bytes instance if set.'):
            writer.editEntry('stream_1', data = 'h')
        with self.assertRaises(ValueError, msg = 'CLSID must be bytes.'):
            writer.editEntry('storage_1', clsid = 'h')
        with self.assertRaises(ValueError, msg = 'CLSID must be 16 bytes.'):
            writer.editEntry('storage_1', clsid = b'h')
        with self.assertRaises(ValueError, msg = 'Modification of creation time cannot be done on a stream.'):
            writer.editEntry('stream_1', creationTime = -1)
        with self.assertRaises(ValueError, msg = 'Creation time must be a positive 8 byte int.'):
            writer.editEntry('storage_1',creationTime = -1)
        with self.assertRaises(ValueError, msg = 'Modification of modified time cannot be done on a stream.'):
            writer.editEntry('stream_1', modifiedTime = -1)
        with self.assertRaises(ValueError, msg = 'Modified time must be a positive 8 byte int.'):
            writer.editEntry('storage_1', modifiedTime = -1)
        with self.assertRaises(ValueError, msg = 'State bits must be a positive 4 byte int.'):
            writer.editEntry('storage_1', stateBits = -1)
        with self.assertRaises(ValueError, msg = 'State bits must be a positive 4 byte int.'):
            writer.editEntry('storage_1', stateBits = 0x100000000)

    def testRemoveEntryManual(self):
        writer = self._setupWriter()

        # Delete the stream. No exceptions should raise.
        writer.deleteEntry('stream_1')

        # Try deleting the data of a missing file.
        with self.assertRaises(OSError, msg = f'Entry not found: not_found'):
            writer.deleteEntry('not_found')

        # Try deleting the data of a bad path.
        with self.assertRaises(ValueError, msg = 'Illegal character ("!" or ":") found in MSG path.'):
            writer.deleteEntry('::invalid')



class OleWriterExportTests(unittest.TestCase):
    def testExportExamples(self, testFileDir = TEST_FILE_DIR):
        """
        Tests exporting the example files.

        Uses the files in export-results to determine which test files to use.
        """
        for exportResultFile in (testFileDir / 'export-results').glob('*.msg'):
            # If we have an export result, find the original file, open it,
            # and export it as bytes to check against the known result.
            with extract_msg.openMsg(testFileDir / exportResultFile.name, delayAttachments = True) as msg:
                exportedBytes = msg.exportBytes()

            with open(exportResultFile, 'rb') as f:
                exportResult = f.read()

            # Use a subtest to print the file name.
            with self.subTest(str(testFileDir / exportResultFile.name)):
                # We use two assertions to give better error messages.
                self.assertCountEqual(exportResult, exportedBytes, 'Exported data is wrong size.')
                self.assertEqual(exportedBytes, exportResult, 'Exported data is incorrect.')

    @unittest.skipIf(USER_TEST_DIR is None, 'User test files not defined.')
    @unittest.skipIf(USER_TEST_DIR is not None and not (USER_TEST_DIR / 'export-results').exists(), 'User export tests not defined.')
    def testExtraExportExamples(self):
        """
        Tests exporting the user's example files.

        Uses the files in export-results to determine which test files to use.
        """
        self.testExportExamples(USER_TEST_DIR)
