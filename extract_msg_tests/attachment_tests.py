__all__ = [
    'AttachmentTests',
]


import unittest

from typing import Callable, Type, TypeVar

from .constants import TEST_FILE_DIR
from extract_msg import enums, MSGFile, openMsg, PropertiesStore
from extract_msg.attachments import (
        Attachment, AttachmentBase, BrokenAttachment, UnsupportedAttachment
    )


_T = TypeVar('_T', bound = AttachmentBase)


class AttachmentTests(unittest.TestCase):
    def testBroken(self):
        attFunction = _forceAttachmentType(BrokenAttachment)
        with openMsg(TEST_FILE_DIR / 'unicode.msg', initAttachment = attFunction) as msg:
            self.assertEqual(len(msg.attachments), 2)

            for att in msg.attachments:
                self.assertIsInstance(att, BrokenAttachment)
                self.assertIs(att.type, enums.AttachmentType.BROKEN)
                self.assertIsNone(att.data)
                with self.assertRaises(NotImplementedError):
                    att.save()
                self.assertEqual(att.save(skipNotImplemented = True), (enums.SaveType.NONE, None))
                with self.assertRaises(NotImplementedError):
                    att.getFilename()

    def testNormal(self):
        # Just covers a bit of the attachment class.
        with openMsg(TEST_FILE_DIR / 'unicode.msg') as msg:
            self.assertEqual(len(msg.attachments), 2)

            for att in msg.attachments:
                self.assertIsInstance(att, Attachment)
                self.assertIs(att.type, enums.AttachmentType.DATA)
                self.assertTrue(att.exists('__properties_version1.0'))
                self.assertTrue(att.exists('__substg1.0_0FF90102'))
                self.assertTrue(att.exists('__substg1.0_37010102'))
                self.assertTrue(att.sExists('__substg1.0_3704'))
                self.assertTrue(att.sExists('__substg1.0_3707'))
                self.assertTrue(att.exists('__substg1.0_370A0102'))
                self.assertTrue(att.sExists('__substg1.0_370E'))

                self.assertIsNotNone(att.getFilename())
                self.assertIsNone(att.attachmentEncoding)
                self.assertIsNone(att.additionalInformation)
                self.assertIsNone(att.cid)
                self.assertIs(att.contentId, att.cid)
                self.assertEqual(att.mimetype, 'image/tiff')
                self.assertIs(att.msg, msg)

                weakRefList = att.treePath
                self.assertEqual(len(weakRefList), 2)
                self.assertIs(weakRefList[0](), msg)
                self.assertIs(weakRefList[1](), att)

                self.assertIsNotNone(att.data)
                self.assertIs(att.dataType, bytes)
                self.assertIsInstance(att.data, bytes)

    def testUnsupported(self):
        attFunction = _forceAttachmentType(UnsupportedAttachment)
        with openMsg(TEST_FILE_DIR / 'unicode.msg', initAttachment = attFunction) as msg:
            self.assertEqual(len(msg.attachments), 2)

            for att in msg.attachments:
                self.assertIsInstance(att, UnsupportedAttachment)
                self.assertIs(att.type, enums.AttachmentType.UNSUPPORTED)
                self.assertIsNone(att.data)
                with self.assertRaises(NotImplementedError):
                    att.save()
                self.assertEqual(att.save(skipNotImplemented = True), (enums.SaveType.NONE, None))
                with self.assertRaises(NotImplementedError):
                    att.getFilename()


def _forceAttachmentType(attType : Type[_T]) -> Callable[[MSGFile, str], _T]:
    def createAttachment(msg : MSGFile, dir_ : str) -> _T:
        propertiesStream = msg.getStream([dir_, '__properties_version1.0'])
        propStore = PropertiesStore(propertiesStream, enums.PropertiesType.ATTACHMENT)
        return attType(msg, dir_, propStore)

    return createAttachment
