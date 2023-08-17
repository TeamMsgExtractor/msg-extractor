from __future__ import annotations


__all__ = [
    'JournalAssociatedAttachment',
]


from functools import cached_property
from typing import Optional, TYPE_CHECKING

from . import registerHandler
from .custom_handler import CustomAttachmentHandler
from ...structures.entry_id import EntryID


if TYPE_CHECKING:
    from ..attachment_base import AttachmentBase


class JournalAssociatedAttachment(CustomAttachmentHandler):
    def __init__(self, attachment : AttachmentBase):
        super().__init__(attachment)

    @classmethod
    def isCorrectHandler(cls, attachment : AttachmentBase) -> bool:
        # This only applies to journal objects.
        if not attachment.msg.classType.lower().startswith('ipm.activity'):
            return False
        if attachment.clsid != '00020D09-0000-0000-C000-000000000046':
            return False

        return True

    @cached_property
    def mailMsgAttFld(self) -> Optional[EntryID]:
        """
        The EntryID of the folder of the linked Message object.
        """
        return EntryID.autoCreate(self.getStream('MailMsgAttFld'))

    @cached_property
    def mailMsgAttMdb(self) -> Optional[EntryID]:
        """
        The EntryID of the store of the linked Message object.
        """
        return EntryID.autoCreate(self.getStream('MailMsgAttMdb'))

    @cached_property
    def mailMsgAttMsg(self) -> Optional[EntryID]:
        """
        The EntryID linked Message object; required only if the
        mailMsgAttSrchKey property is None.
        """
        return EntryID.autoCreate(self.getStream('MailMsgAttMsg'))

    @cached_property
    def mailMsgAttSrchFld(self) -> Optional[EntryID]:
        """
        The object EntryID of the Sent Items special folder of the linked
        Message object.
        """
        return EntryID.autoCreate(self.getStream('MailMsgAttSrchFld'))

    @cached_property
    def mailMsgAttSrchKey(self) -> Optional[bytes]:
        """
        The search key for the linked message object; required only if
        mailMsgAttMsg is None.
        """
        return self.getStream('MailMsgAttSrchKey')

    @cached_property
    def metafileBytes(self) -> Optional[bytes]:
        """
        The metafile that contains the icon to be used when rendering the
        attachment.

        From my understanding, this MUST be set, but we are treating it as
        SHOULD be set.
        """
        # The documentation specifies clearly that the filename is "IOlePres000"
        # HOWEVER my tests revealed that the "I" is actually a "\x02" character.
        # This is quite confusing but whatever. We'll just look for both of
        # them.
        return self.getStream('IOlePres000') or self.getStream('\x02OlePres000')



registerHandler(JournalAssociatedAttachment)