import shutil
import os
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import extract_msg

TEST_FILE = "example-msg-files/unicode.msg"
base = extract_msg

class TestCase(unittest.TestCase):

    def setUp(self):
        shutil.rmtree("raw", ignore_errors=True)
        for name in os.listdir('.'):
            if os.path.isdir(name) and name.startswith('2013-11-18_1026'):
                shutil.rmtree(name, ignore_errors=True)

    tearDown = setUp

    def test_message(self):
        msg = base.Message(TEST_FILE)
        self.assertEqual(msg.subject, u'Test for TIF files')
        self.assertEqual(
            msg.body,
            u'This is a test email to experiment with the MS Outlook MSG '
            u'Extractor\r\n\r\n\r\n-- \r\n\r\n\r\nKind regards'
            u'\r\n\r\n\r\n\r\n\r\nBrian Zhou\r\n\r\n')
        #
        # Note, this will fail in all system time zones outside of +0200.  We need to parse the returned date and compare it with the current time zone.
        # Expected :Mon, 18 Nov 2013 03:26:24 -0500
        # Actual   :Mon, 18 Nov 2013 10:26:24 +0200
        #
        if '+0200' in msg.date:
            self.assertEqual(msg.date, 'Mon, 18 Nov 2013 10:26:24 +0200')
            self.assertEqual(msg.parsedDate, (2013, 11, 18, 10, 26, 24, 0, 1, -1))
        self.assertEqual(msg.sender, 'Brian Zhou <brizhou@gmail.com>')
        self.assertEqual(msg.to, 'brianzhou@me.com')
        self.assertEqual(msg.cc, 'Brian Zhou <brizhou@gmail.com>')
        self.assertEqual(len(msg.attachments), 2)
        msg.dump()
        msg.debug()

    def test_save(self):
        msg = base.Message(TEST_FILE)
        msg.save()
        self.assertEqual(
            sorted(os.listdir('2013-11-18_1026 Test for TIF files')),
            sorted(['message.text', 'import OleFileIO.tif',
                    'raised value error.tif']))
        msg.saveRaw()

    def test_saveRaw(self):
        msg = base.Message(TEST_FILE)
        msg.saveRaw()
        assert os.listdir('raw')

    def test_saveInCustomLocation(self):
        """
        Tests saving an attachment into a custom location.
        """
        msg = base.Message(TEST_FILE)

        tempDir = TemporaryDirectory(prefix='msg-extractor--test-')
        tmpDirPath = Path(tempDir.name)

        fileNamePaths: [Path] = []

        for att in msg.attachments:
            att: extract_msg.attachment.Attachment

            longFileName = att.longFilename
            assert longFileName, 'Attachment should have a long file name'

            shortFilename = att.shortFilename
            assert shortFilename, 'Attachment should have a short file name'

            savedAttFile = att.save(customPath=tmpDirPath, customFilename=longFileName)
            assert savedAttFile, 'Attachment have returned the exact path'

            assert Path(savedAttFile).exists(), 'returned path should actually exist'
            fileNamePaths.append(savedAttFile)


        #
        # Assert the proper output of the files by ensuring their hash.  We could check the size, but that would
        # just change the hash anyway, so we just check the file name and hash.
        #
        expectedResults = [
            {'hash': '6f36cb718943751db9dc4c9df624c4390d8a13674127b7e919f419061e856dfa', 'filename': 'import OleFileIO.tif'},
            {'hash': '3b43fdeca80da38c80e918e8f21cbd6fc925dac994f3922c7893fc6c1326fb92', 'filename': 'raised value error.tif'}
        ]

        for expectedResult in expectedResults:
            expectedHash = expectedResult['hash']
            expectedFileName = expectedResult['filename']
            expectedPath = Path(tempDir.name).joinpath(expectedFileName)
            hashOnDisk = self.hashFile(filename=expectedPath)
            assert expectedHash == hashOnDisk, f'File \'{expectedFileName}\' should have had hash {expectedHash}, but had a hash of {hashOnDisk}'

        tempDir.cleanup()


    def hashFile(self, filename: Path):
        """
        Helper function that returns a sha hash for a particular file, for verification.
        """
        import hashlib
        hash = hashlib.sha256()

        with open(filename, 'rb') as file:
            while True:
                # Reading is buffered, so we can read smaller chunks.
                chunk = file.read(hash.block_size)
                if not chunk:
                    break
                hash.update(chunk)

        return str(hash.hexdigest())

if __name__ == '__main__':
    """
    Main entry point for unit tests.
    """
    unittest.main(verbosity=2)

