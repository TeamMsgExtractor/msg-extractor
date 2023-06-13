from __future__ import annotations


__all__ = [
    'CustomAttachmentHandler',
]


import abc

from typing import List, Optional, Tuple, TYPE_CHECKING


if TYPE_CHECKING:
    from ..attachment import Attachment


class CustomAttachmentHandler(abc.ABC):
    """
    A class designed to help with custom attachments that may require parsing in
    special ways that are completely different from one another.
    """

    def __init__(self, attachment : Attachment):
        super().__init__()
        self.__att = attachment

    @classmethod
    @abc.abstractmethod
    def isCorrectHandler(cls, attachment : Attachment) -> bool:
        """
        Checks if this is the correct handler for the attachment.
        """

    @abc.abstractmethod
    def generateRtf(self) -> Optional[bytes]:
        """
        Generates the RTF to inject in place of the \objattph tag.

        If this function should do nothing, returns None.
        """

    @property
    def attachment(self):
        """
        The attachment this handler is associated with.
        """
        return self.__att

    @property
    @abc.abstractmethod
    def data(self) -> bytes:
        """
        Gets the data for the attachment.
        """

    @property
    @abc.abstractmethod
    def name(self) -> str:
        """
        Returns the name to be used when saving the attachment.
        """
