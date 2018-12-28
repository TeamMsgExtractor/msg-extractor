import shutil
import os
import unittest

import extract_msg

TEST_FILE = "example-msg-files/unicode.msg"


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


unittest.main(verbosity=2)
