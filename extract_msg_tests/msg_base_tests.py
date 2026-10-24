__all__ = [
    'MSGBaseTests',
]


import logging
import sys
import unittest

import extract_msg

from .constants import TEST_FILE_DIR
from extract_msg import openMsg
from extract_msg.exceptions import StandardViolationError
from extract_msg.enums import ErrorBehavior


# Handle NoLogs not existing in 3.8 and 3.9.
if sys.version_info[1] < 10:
    class NoLogsHandler:
        def __enter__(self, *args, **kwargs):
            pass

        def __exit__(self, *args, **kwargs):
            pass


class MSGBaseTests(unittest.TestCase):
    # Handle NoLogs not existing in 3.8 and 3.9
    if sys.version_info[1] < 10:
        logging.info('Version does not support checking for an absence of logs.')
        def assertNoLogs(self, *args, **kwargs):
            return NoLogsHandler()

    def testMutipleStreams(self):
        """
        Tests getting multiple-type value streams from the test file.
        """
        msgLogger = extract_msg.msg_classes.msg.logger
        expectedOutputs = {
            'multiBin': [
                b'\x00\x01\x02\x03',
                b'\x04\x05\x06\x07',
                b'\x08\x09\x0A\x0B',
                b'\x0C\x0D\x0E\x0F',
            ],
            'multiStr': [
                'Hello',
                'World',
                'It\'s',
                'Me!',
            ],
            'shortBin': [
                b'\x00\x01\x02\x03',
                b'\x04\x05\x06\x07',
                b'\x08\x09\x0A\x0B',
                #b'\x0C\x0D\x0E\x0F', # This one is missing.
            ],
            'shortStr': [
                'Hello',
                'World',
                'It\'s',
                #'Me!', # This one is missing.
            ],
            'missingBin': [
                b'\x00\x01\x02\x03',
                b'\x04\x05\x06\x07',
                #b'\x08\x09\x0A\x0B', # This one is missing.
                #b'\x0C\x0D\x0E\x0F', # This one should be absent.
            ],
            'missingStr': [
                'Hello',
                'World',
                #'It\'s', # This one is missing.
                #'Me!', # This one should be absent.
            ],
        }

        with openMsg(TEST_FILE_DIR / 'multi-stream.msg', strict = False) as msg:
            # These two should fetch normally.
            with self.subTest('Test multiple binary'):
                with self.assertNoLogs(msgLogger, logging.WARNING):
                    res = msg.getMultipleBinary('__substg1.0_0000')
                self.assertIsNotNone(res, 'Failed to fetch multiple binary stream.')
                self.assertListEqual(res, expectedOutputs['multiBin'])

            with self.subTest('Test multiple string'):
                with self.assertNoLogs(msgLogger, logging.WARNING):
                    res = msg.getMultipleString('__substg1.0_0001')
                self.assertIsNotNone(res, 'Failed to fetch multiple string stream.')
                self.assertListEqual(res, expectedOutputs['multiStr'])

            # These two should be short a single item.
            with self.subTest('Test too few multiple binary'):
                with self.assertLogs(msgLogger, logging.ERROR):
                    res = msg.getMultipleBinary('__substg1.0_0002')
                self.assertIsNotNone(res, 'Failed to fetch multiple binary stream.')
                self.assertListEqual(res, expectedOutputs['shortBin'])

            with self.subTest('Test too few multiple string'):
                with self.assertLogs(msgLogger, logging.ERROR):
                    res = msg.getMultipleString('__substg1.0_0003')
                self.assertIsNotNone(res, 'Failed to fetch multiple string stream.')
                self.assertListEqual(res, expectedOutputs['shortStr'])

            # These two should be missing an item and skip the remaining.
            with self.subTest('Test missing multiple binary'):
                with self.assertLogs(msgLogger, logging.ERROR):
                    res = msg.getMultipleBinary('__substg1.0_0004')
                self.assertIsNotNone(res, 'Failed to fetch multiple binary stream.')
                self.assertListEqual(res, expectedOutputs['missingBin'])

            with self.subTest('Test missing multiple string'):
                with self.assertLogs(msgLogger, logging.ERROR):
                    res = msg.getMultipleString('__substg1.0_0005')
                self.assertIsNotNone(res, 'Failed to fetch multiple string stream.')
                self.assertListEqual(res, expectedOutputs['missingStr'])

            # These two should report a predicted length mismatch.
            with self.subTest('Test multiple binary mismatch length'):
                with self.assertLogs(msgLogger, logging.WARNING):
                    res = msg.getMultipleBinary('__substg1.0_0006')
                self.assertIsNotNone(res, 'Failed to fetch multiple binary stream.')
                self.assertListEqual(res, expectedOutputs['multiBin'])

            with self.subTest('Test multiple string mismatch length'):
                with self.assertLogs(msgLogger, logging.WARNING):
                    res = msg.getMultipleString('__substg1.0_0007')
                self.assertIsNotNone(res, 'Failed to fetch multiple string stream.')
                self.assertListEqual(res, expectedOutputs['multiStr'])

            # These two should fail because the length streams are the wrong
            # size.
            with self.subTest('Test standard violation (multiple binary)'):
                with self.assertRaises(StandardViolationError):
                    msg.getMultipleBinary('__substg1.0_F000')

            with self.subTest('Test standard violation (multiple string)'):
                with self.assertRaises(StandardViolationError):
                    msg.getMultipleString('__substg1.0_F001')

            # Test failure because of missing null byte.
            with self.subTest('Test standard violation (multiple string no null)'):
                with self.assertRaises(StandardViolationError):
                    msg.getMultipleString('__substg1.0_F003')

        # Reopen the msg file with standard violations disabled.
        with openMsg(TEST_FILE_DIR / 'multi-stream.msg', strict = False, errorBehavior = ErrorBehavior.STANDARDS_VIOLATION) as msg:
            with self.subTest('Test no standard violation (multiple binary)'):
                with self.assertLogs(msgLogger, logging.ERROR):
                    msg.getMultipleBinary('__substg1.0_F000')

            with self.subTest('Test no standard violation (multiple string)'):
                with self.assertLogs(msgLogger, logging.ERROR):
                    msg.getMultipleString('__substg1.0_F001')

            with self.subTest('Test no standard violation (multiple string no null)'):
                with self.assertLogs(msgLogger, logging.ERROR):
                    msg.getMultipleString('__substg1.0_F003')