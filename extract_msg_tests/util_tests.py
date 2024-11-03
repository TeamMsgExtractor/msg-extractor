__all__ = [
    'UtilTests',
]


import unittest

from extract_msg import utils


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