from __future__ import annotations


__all__ = [
    'CustomAttachmentHandler',
]


import abc
import functools

from typing import Dict, Optional, TYPE_CHECKING, TypeVar

from ...constants import MSG_PATH, OVERRIDE_CLASS
from ...structures.odt import ODTStruct
from ...structures.ole_pres import OLEPresentationStream
from ...structures.ole_stream_struct import OleStreamStruct
from ...utils import msgPathToString


if TYPE_CHECKING:
    from ..attachment_base import AttachmentBase

_T = TypeVar('_T')


class CustomAttachmentHandler(abc.ABC):
    """
    A class designed to help with custom attachments that may require parsing in
    special ways that are completely different from one another.
    """

    def __init__(self, attachment: AttachmentBase):
        super().__init__()
        self.__att = attachment

    def getStream(self, filename: MSG_PATH) -> Optional[bytes]:
        """
        Gets a stream from the custom data directory.
        """
        return self.attachment.getStream('__substg1.0_3701000D/' + msgPathToString(filename))

    def getStreamAs(self, streamID: MSG_PATH, overrideClass: OVERRIDE_CLASS[_T]) -> Optional[_T]:
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
    def isCorrectHandler(cls, attachment: AttachmentBase) -> bool:
        """
        Checks if this is the correct handler for the attachment.
        """

    @abc.abstractmethod
    def generateRtf(self) -> Optional[bytes]:
        """
        Generates the RTF to inject in place of the \\objattph tag.

        If this function should do nothing, returns ``None``.
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

        If an attachment should do nothing when saving, returns ``None``.
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
        Returns an object representing the data.

        May return the same value as :property data:.

        If there is no object to represent the custom attachment, including
        ``bytes``, returns ``None``.
        """

    @functools.cached_property
    def objInfo(self) -> Optional[ODTStruct]:
        """
        The structure representing the stream "\\x03ObjInfo", if it exists.
        """
        return self.getStreamAs('\x03ObjInfo', ODTStruct)

    @functools.cached_property
    def ole(self) -> Optional[OleStreamStruct]:
        """
        The structure representing the stream "\\x01Ole", if it exists.
        """
        return self.getStreamAs('\x01Ole', OleStreamStruct)

    @functools.cached_property
    def presentationObjs(self) -> Optional[Dict[int, OLEPresentationStream]]:
        """
        Returns a dict of all presentation streams, as bytes.
        """
        return {
            int(x[1][-3:]): self.getStreamAs(x[-1], OLEPresentationStream)
            for x in self.attachment.listDir()
            if x[0] == '__substg1.0_3701000D' and x[1].startswith('\x02OlePres')
        }