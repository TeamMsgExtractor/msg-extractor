__all__ = [
    'MessageTests',
]


import base64
import unittest

from email.message import EmailMessage

from .constants import TEST_FILE_DIR
from extract_msg import openMsg
from extract_msg.msg_classes import Message


class MessageTests(unittest.TestCase):
    def testMultiTo(self):
        """
        Tests parsing a message with multiple To and CC recipients.
        """
        with openMsg(TEST_FILE_DIR / 'multi-to.msg') as msg:
            self.assertIsInstance(msg, Message)

            self.assertTrue(msg.subject.startswith('Test: multiple To recipients'))
            self.assertEqual(msg.sender, 'Bob Sender <bob@example.com>')
            self.assertTrue((msg.body or '').startswith('Test email body.'))

            # Expect at least two To recipients (type 1) and at least one CC (type 2).
            to_recipients = [r for r in msg.recipients if r.type == 1]
            cc_recipients = [r for r in msg.recipients if r.type == 2]

            self.assertGreaterEqual(len(to_recipients), 2)
            self.assertGreaterEqual(len(cc_recipients), 1)

            to_emails = [r.email.strip('\x00') for r in to_recipients]
            self.assertIn('alice@example.com', to_emails)
            self.assertIn('carol@example.com', to_emails)

            cc_emails = [r.email.strip('\x00') for r in cc_recipients]
            self.assertIn('dave@example.com', cc_emails)

    def testMultiToTo(self):
        """
        Tests parsing a message where To headers use mixed casing (e.g. 'To' vs 'TO').
        """
        with openMsg(TEST_FILE_DIR / 'multi-to-to.msg') as msg:
            self.assertIsInstance(msg, Message)
            self.assertTrue(msg.subject.startswith('Test: multiple To recipients'))

            to_recipients = [r for r in msg.recipients if r.type == 1]
            cc_recipients = [r for r in msg.recipients if r.type == 2]

            self.assertGreaterEqual(len(to_recipients), 2)
            self.assertGreaterEqual(len(cc_recipients), 1)

            to_emails = [r.email.strip('\x00') for r in to_recipients]
            self.assertIn('alice@example.com', to_emails)
            self.assertIn('carol@example.com', to_emails)

    def testMultiToToAsEmailMessage(self):
        """
        Tests EML conversion when To headers appear with mixed casing across recipients.
        """
        with openMsg(TEST_FILE_DIR / 'multi-to-to.msg') as msg:
            em = msg.asEmailMessage()

            self.assertIsInstance(em, EmailMessage)
            self.assertEqual(sum(1 for k in em.keys() if k.lower() == 'to'), 1)
            to_header = em['TO'] or em['To']
            self.assertIn('alice@example.com', to_header)
            self.assertIn('carol@example.com', to_header)

    def testUnicodeHeaderAsEmailMessage(self):
        """
        Tests that a message whose To header mixes plain and RFC 2047-encoded
        display names (using a charset with invalid byte sequences) can be
        converted to EML bytes without a UnicodeEncodeError. This crashes on
        Python 3.13 without the fix.
        """
        with openMsg(TEST_FILE_DIR / 'unicode-header.msg') as msg:
            em = msg.asEmailMessage()
            self.assertIsInstance(em, EmailMessage)
            raw = em.as_bytes()
            self.assertIn(b'alice@example.com', raw)

    def testGbkFallbackDisplayName(self):
        """
        Tests that RFC 2047 encoded words declared as GB2312 but containing
        byte sequences only valid in GBK (a strict superset) are decoded
        correctly rather than mangled via latin-1 fallback.

        The encoded word =?gb2312?B?6pCzydXCKG1heGNoZW4p?= decodes to
        '陳成章(maxchen)' in GBK. Without the fix, the display name is
        garbled as latin-1.
        """
        with openMsg(TEST_FILE_DIR / 'unicode-header.msg') as msg:
            em = msg.asEmailMessage()
            raw = em.as_bytes()
            # The correctly GBK-decoded name must appear RFC 2047-encoded as UTF-8.
            # The Chinese characters 陳成章 base64-encoded in UTF-8 is the marker.
            self.assertIn(base64.b64encode('陳成章'.encode('utf-8')), raw)

    def testMultiToAsEmailMessage(self):
        """
        Tests that a message with multiple To recipients converts to EML without error,
        and that duplicate To headers are merged into a single comma-separated value.
        """
        with openMsg(TEST_FILE_DIR / 'multi-to.msg') as msg:
            em = msg.asEmailMessage()

            self.assertIsInstance(em, EmailMessage)
            # EmailMessage must have exactly one TO field (duplicates merged).
            self.assertEqual(sum(1 for k in em.keys() if k == 'TO'), 1)
            to_header = em['TO']
            self.assertIn('alice@example.com', to_header)
            self.assertIn('carol@example.com', to_header)
            self.assertEqual(em['CC'], 'Dave Jones <dave@example.com>')
            self.assertEqual(em['From'], 'Bob Sender <bob@example.com>')
