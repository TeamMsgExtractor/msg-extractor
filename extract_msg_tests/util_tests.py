__all__ = [
    'UtilTests',
]


import unittest

import RTFDE

from .constants import TEST_FILE_DIR, USER_TEST_DIR
from extract_msg import enums, openMsg, utils
from extract_msg.msg_classes import MessageBase


class UtilTests(unittest.TestCase):
    def test_dictGetCasedKey(self):
        caseDict = {'hello': 1, 'HeUtQjWkW': 2}

        self.assertEqual(utils.dictGetCasedKey(caseDict, 'Hello'), 'hello')
        self.assertEqual(utils.dictGetCasedKey(caseDict, 'heutqjwkw'), 'HeUtQjWkW')
        with self.assertRaises(KeyError):
            utils.dictGetCasedKey(caseDict, 'jjjjj')

    def test_divide(self):
        inputString = '12345678901234567890'
        expectedOutputs = {
            1: ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0'],
            2: ['12', '34', '56', '78', '90', '12', '34', '56', '78', '90'],
            3: ['123', '456', '789', '012', '345', '678', '90'],
            4: ['1234', '5678', '9012', '3456', '7890'],
            5: ['12345', '67890', '12345', '67890'],
            6: ['123456', '789012', '345678', '90'],
            7: ['1234567', '8901234', '567890'],
            8: ['12345678', '90123456', '7890'],
            9: ['123456789', '012345678', '90'],
            10: ['1234567890', '1234567890'],
            11: ['12345678901', '234567890'],
        }

        for divideBy, expectedResult in expectedOutputs.items():
            self.assertListEqual(utils.divide(inputString, divideBy), expectedResult)

    def test_makeWeakRef(self):
        self.assertIsNone(utils.makeWeakRef(None))
        class TestClass:
            pass
        self.assertIsNotNone(utils.makeWeakRef(TestClass()))

    def test_minutesToDurationStr(self):
        self.assertEqual(utils.minutesToDurationStr(0), '0 hours')
        self.assertEqual(utils.minutesToDurationStr(1), '1 minute')
        self.assertEqual(utils.minutesToDurationStr(2), '2 minutes')
        self.assertEqual(utils.minutesToDurationStr(59), '59 minutes')
        self.assertEqual(utils.minutesToDurationStr(60), '1 hour')
        self.assertEqual(utils.minutesToDurationStr(61), '1 hour 1 minute')
        self.assertEqual(utils.minutesToDurationStr(62), '1 hour 2 minutes')
        self.assertEqual(utils.minutesToDurationStr(120), '2 hours')
        self.assertEqual(utils.minutesToDurationStr(121), '2 hours 1 minute')
        self.assertEqual(utils.minutesToDurationStr(122), '2 hours 2 minutes')

    def test_msgPathToStr(self):
        self.assertEqual(utils.msgPathToString('hello/world/one'), 'hello/world/one')
        self.assertEqual(utils.msgPathToString('hello/world\\one'), 'hello/world/one')
        self.assertEqual(utils.msgPathToString(['hello', 'world', 'one']), 'hello/world/one')
        self.assertEqual(utils.msgPathToString(['hello\\world', 'one']), 'hello/world/one')

    def test_stripRtf(self, testFileDir = TEST_FILE_DIR):
        """
        Tests the stripRtf() function.

        Confirms that the output of pure deencapsulation is identical after
        using stripRtf(). By default this works on the provided tests, however
        an alternate tests calls this method for user-defined tests.
        """
        for testFile in testFileDir.glob('*.msg'):
            # If we have an export result, find the original file, open it,
            # and export it as bytes to check against the known result.
            with openMsg(testFile, delayAttachments = True) as msg:
                if not isinstance(msg, MessageBase) or not msg.rtfBody:
                    continue

                rtfBody = msg.rtfBody

                try:
                    htmlBody = msg.deencapsulateBody(rtfBody, enums.DeencapType.HTML)
                except Exception: 
                    # If it's throwing an exception, it's not our problem here.
                    htmlBody = None

            if rtfBody and htmlBody:
                de = RTFDE.DeEncapsulator(rtfBody)
                de.deencapsulate()

                with self.subTest(str(testFile)):
                    self.assertEqual(htmlBody, de.html, 'Stripped output does not match original.')

        for testFile in testFileDir.glob('*.rtf'):
            # If we have an export result, find the original file, open it,
            # and export it as bytes to check against the known result.
            with open(testFile, 'rb') as f:
                rtfBody = f.read()

            try:
                de = RTFDE.DeEncapsulator(utils.stripRtf(rtfBody))
                de.deencapsulate()
                htmlBody = de.html
            except Exception: 
                # If it's throwing an exception, it's not our problem here.
                htmlBody = None

            if rtfBody and htmlBody:
                de = RTFDE.DeEncapsulator(rtfBody)
                de.deencapsulate()

                with self.subTest(str(testFile)):
                    self.assertEqual(htmlBody, de.html, 'Stripped output does not match original.')

    @unittest.skipIf(USER_TEST_DIR is None, 'User test files not defined.')
    @unittest.skipIf(USER_TEST_DIR is not None and not (USER_TEST_DIR / 'export-results').exists(), 'User export tests not defined.')
    def testUser_stripRtf(self):
        """
        Tests exporting the user's example files.

        Uses the files in export-results to determine which test files to use.
        """
        self.test_stripRtf(USER_TEST_DIR)
