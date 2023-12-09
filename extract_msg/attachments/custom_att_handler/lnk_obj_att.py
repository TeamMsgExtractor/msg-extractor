from __future__ import annotations


__all__ = [
    'LinkedObjectAttachment',
]


from functools import cached_property
from typing import List, Optional, TYPE_CHECKING

from . import registerHandler
from .custom_handler import CustomAttachmentHandler
from ...structures.entry_id import EntryID
from ...structures.ole_pres import OLEPresentationStream


if TYPE_CHECKING:
    from ..attachment_base import AttachmentBase


class LinkedObjectAttachment(CustomAttachmentHandler):
    """
    A link to an Outlook object.

    Not *positive* I understand what this attachment type is, but this seems to
    be the most likely name. Contains presentation data about how to render it
    as well as properties with data that link to it. It looks *similar* to what
    the documentation for Journal specifies would be it's custom attachment
    type, however some small details don't perfectly add up.

    I've also only seen this on Journal objects thus far.
    """

    def __init__(self, attachment: AttachmentBase):
        super().__init__(attachment)
        stream = attachment.getStream('__substg1.0_3701000D/\x03MailStream')
        if not stream:
            raise ValueError('MailStream could not be found.')
        if len(stream) != 12:
            raise ValueError('MailStream is the wrong length.')

    @classmethod
    def isCorrectHandler(cls, attachment: AttachmentBase) -> bool:
        if attachment.clsid != '00020D09-0000-0000-C000-000000000046':
            return False

        return True

    def generateRtf(self) -> Optional[bytes]:
        # TODO
        return None

    @property
    def data(self) -> None:
        # This type of attachment has no direct associated data.
        return None

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
        The EntryID linked Message object, required only if the
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
        The search key for the linked message object, required only if
        mailMsgAttMsg is None.
        """
        return self.getStream('MailMsgAttSrchKey')

    @property
    def name(self) -> None:
        # Doesn't save.
        return None

    @property
    def obj(self) -> None:
        # No object to represent this.
        return None



registerHandler(LinkedObjectAttachment)