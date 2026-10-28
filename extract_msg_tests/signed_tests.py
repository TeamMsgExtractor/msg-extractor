__all__ = [
    'SignedTests',
]


import pathlib
import unittest

from .constants import TEST_FILE_DIR, USER_TEST_DIR
from extract_msg import openMsg
from extract_msg.msg_classes import MessageSignedBase


class SignedTests(unittest.TestCase):
    def testUnwrap(self, testFileDir: pathlib.Path = TEST_FILE_DIR):
        """
        Tests the results of extract_msg.utils.unwrapMultipart().

        The initial test is ensuring the attachments generate at all. After
        that, test to make sure the types in the signedBody and signedHtmlBody
        are correct.
        """
        for path in TEST_FILE_DIR.glob('*.msg'):
            with openMsg(path, strict = False, delayAttachments = True) as msg:
                if isinstance(msg, MessageSignedBase):
                    with self.subTest(str(path)):
                        msg.attachments
                        if msg.signedBody:
                            self.assertIsInstance(msg.signedBody, str)
                        if msg.signedHtmlBody:
                            self.assertIsInstance(msg.signedHtmlBody, bytes)

    @unittest.skipIf(USER_TEST_DIR is None, 'User test files not defined.')
    def testUserStdin(self):
        self.testUnwrap(USER_TEST_DIR)
