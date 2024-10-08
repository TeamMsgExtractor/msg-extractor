__all__ = [
    'ClipboardFormatOrAnsiString',
]


from typing import Optional, Union

from .. import constants
from ._helpers import BytesReader
from ..enums import ClipboardFormat


class ClipboardFormatOrAnsiString:
    def __init__(self, reader: Optional[Union[bytes, BytesReader]] = None):
        if reader is None:
            self.__markerOrLength = 0
            self.__clipboardFormat = None
            self.__ansiString = None
            return

        if isinstance(reader, bytes):
            reader = BytesReader(reader)

        self.__markerOrLength = reader.readUnsignedInt()
        if self.__markerOrLength > 0xFFFFFFFD:
            self.__ansiString = None
            self.__clipboardFormat = ClipboardFormat(reader.readUnsignedInt())
        elif self.__markerOrLength > 0:
            self.__ansiString = reader.read(self.__markerOrLength)
            self.__clipboardFormat = None
        else:
            self.__ansiString = None
            self.__clipboardFormat = None

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        ret = constants.st.ST_LE_UI32.pack(self.markerOrLength)
        if self.markerOrLength > 0xFFFFFFFD:
            ret += constants.st.ST_LE_UI32.pack(self.clipboardFormat)
        elif self.markerOrLength > 0:
            ret += self.__ansiString
        return ret

    @property
    def ansiString(self) -> Optional[bytes]:
        """
        The null-terminated ANSI string, as bytes, of the name of a registered
        clipboard format. Only set if markerOrLength is not ``0x00000000``,
        ``0xFFFFFFFE``, or ``0xFFFFFFFF``.

        Setting this will modify the markerOrLength field automatically.
        """
        return self.__ansiString

    @ansiString.setter
    def ansiString(self, val: bytes) -> None:
        if not val:
            raise ValueError('Cannot set :property ansiString: to None or empty bytes.')

        self.__ansiString = val

    @property
    def clipboardFormat(self) -> Optional[ClipboardFormat]:
        """
        The clipboard format, if any.

        To set this, make sure that :property markerOrLength: is ``0xFFFFFFFE``
        or ``0xFFFFFFFF`` *before* setting.
        """
        return self.__clipboardFormat

    @clipboardFormat.setter
    def clipboardFormat(self, val: ClipboardFormat) -> None:
        if not val:
            raise ValueError('Cannot set clipboard format to None.')
        if self.markerOrLength < 0xFFFFFFFE:
            raise ValueError('Cannot set the clipboard format while the marker or length is not 0xFFFFFFFE or 0xFFFFFFFF')
        self.__clipboardFormat = val

    @property
    def markerOrLength(self) -> int:
        """
        If set the 0x00000000, then neither the format property nor the
        ansiString property will be set. If it is 0xFFFFFFFF or 0xFFFFFFFE, then
        the clipboardFormat property will be set. Otherwise, the ansiString
        property
        will be set.
        """
        return self.__markerOrLength

    @markerOrLength.setter
    def markerOrLength(self, val: int) -> None:
        if val < 0:
            raise ValueError(':property markerOrLength: must be a positive integer.')
        if val > 0xFFFFFFFF:
            raise ValueError(':property markerOrLength: cannot be greater than 0xFFFFFFFF')

        if val == 0:
            self.__ansiString = None
            self.__clipboardFormat = None
        elif val > 0xFFFFFFFD:
            self.__ansiString = None
            self.__clipboardFormat = ClipboardFormat.CF_BITMAP
        else:
            raise ValueError('Cannot set :property markerOrLength: to a length value. Set :property ansiString: instead.')
        self.__markerOrLength = val