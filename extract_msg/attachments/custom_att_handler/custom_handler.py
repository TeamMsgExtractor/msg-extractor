from __future__ import annotations


__all__ = [
    'CustomAttachmentHandler',
]


import abc

from typing import Any, Callable, Optional, TYPE_CHECKING, TypeVar

from ...constants import MSG_PATH
from ...utils import msgPathToString


if TYPE_CHECKING:
    from ..attachment_base import AttachmentBase

_T = TypeVar('_T')


class CustomAttachmentHandler(abc.ABC):
    """
    A class designed to help with custom attachments that may require parsing in
    special ways that are completely different from one another.
    """

    def __init__(self, attachment : AttachmentBase):
        super().__init__()
        self.__att = attachment

    def getStream(self, path : MSG_PATH) -> Optional[bytes]:
        """
        Gets a stream from the custom data directory.
        """
        return self.attachment.getStream('__substg1.0_3701000D/' + msgPathToString(path))

    def getStreamAs(self, streamID : MSG_PATH, overrideClass : Callable[[Any], _T]) -> Optional[_T]:
        """
        Returns the specified stream, modifying it to the specified class if it
        is found.

        :param overrideClass: Class/function to use to morph the data that was
            read. The data will be the first argument to the class's __init__
            function or the function itself, if that is what is provided. If
            the value is None, this function is not called. If you want it to
            be called regardless, you should handle the data directly.
        """
        value = self.getStream(streamID)

        if value is not None:
            value = overrideClass(value)

        return value

    @classmethod
    @abc.abstractmethod
    def isCorrectHandler(cls, attachment : AttachmentBase) -> bool:
        """
        Checks if this is the correct handler for the attachment.
        """

    @abc.abstractmethod
    def generateRtf(self) -> Optional[bytes]:
        """
        Generates the RTF to inject in place of the \\objattph tag.

        If this function should do nothing, returns None.
        """

    @property
    def attachment(self) -> AttachmentBase:
        """
        The attachment this handler is associated with.
        """
        return self.__att

    @property
    @abc.abstractmethod
    def data(self) -> Optional[bytes]:
        """
        Gets the data for the attachment.

        If an attachment should do nothing when saving, returns None.
        """

    @property
    @abc.abstractmethod
    def name(self) -> Optional[str]:
        """
        Returns the name to be used when saving the attachment.
        """

    @property
    @abc.abstractmethod
    def obj(self) -> Optional[object]:
        """
        Returns an object representing the data. May return the same as
        :property data:.

        If there is no object to represent the custom attachment, including
        bytes, returns None.
        """