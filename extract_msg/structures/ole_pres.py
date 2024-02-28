from __future__ import annotations


__all__ = [
    'OLEPresentationStream',
]


from typing import List, Optional, Union

from ._helpers import BytesReader
from .cfoas import ClipboardFormatOrAnsiString
from ..constants import st
from .dv_target_device import DVTargetDevice
from ..enums import ADVF, ClipboardFormat, DVAspect
from .toc_entry import TOCEntry


class OLEPresentationStream:
    """
    [MS-OLEDS] OLEPresentationStream.
    """

    def __init__(self, data: bytes):
        reader = BytesReader(data)
        acf = self.__ansiClipboardFormat = ClipboardFormatOrAnsiString(reader)

        # Validate the structure based on the documentation.
        if acf.markerOrLength == 0:
            raise ValueError('Invalid OLEPresentationStream (MarkerOrLength is 0).')
        if acf.clipboardFormat is ClipboardFormat.CF_BITMAP:
            raise ValueError('Invalid OLEPresentationStream (Format is CF_BITMAP).')
        if 0x201 < acf.markerOrLength < 0xFFFFFFFE:
            raise ValueError('Invalid OLEPresentationStream (ANSI length was more than 0x201).')

        targetDeviceSize = reader.readUnsignedInt()
        if targetDeviceSize < 0x4:
            raise ValueError('Invalid OLEPresentationStream (TargetDeviceSize was less than 4).')
        if targetDeviceSize > 0x4:
            # Read the TargetDevice field.
            self.__targetDevice = DVTargetDevice(reader.read(targetDeviceSize))
        else:
            self.__targetDevice = None

        self.__aspect = reader.readUnsignedInt()
        self.__lindex = reader.readUnsignedInt()
        self.__advf = reader.readUnsignedInt()

        # Reserved1.
        self.__reserved1 = reader.read(4)

        self.__width = reader.readUnsignedInt()
        self.__height = reader.readUnsignedInt()
        size = reader.readUnsignedInt()
        self.__data = reader.read(size)

        if acf.clipboardFormat is ClipboardFormat.CF_METAFILEPICT:
            self.__reserved2 = reader.read(18)
        else:
            self.__reserved2 = b'\x00' * 18.

        self.__tocSignature = reader.readUnsignedInt()
        self.__tocEntries = []
        if self.__tocSignature == 0x494E414E: # b'NANI' in little endian.
            for _ in range(reader.readUnsignedInt()):
                self.__tocEntries.append(TOCEntry(reader))

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        ret = bytes(self.__ansiClipboardFormat)

        if self.__targetDevice is None:
            ret += b'\x04\x00\x00\x00'
        else:
            dvData = bytes(self.__targetDevice)
            ret += st.ST_LE_UI32.pack(len(dvData))
            ret += dvData

        ret += st.ST_LE_UI32.pack(self.__aspect)
        ret += st.ST_LE_UI32.pack(self.__lindex)
        ret += st.ST_LE_UI32.pack(self.__advf)
        ret += self.__reserved1
        ret += st.ST_LE_UI32.pack(self.__width)
        ret += st.ST_LE_UI32.pack(self.__height)
        ret += st.ST_LE_UI32.pack(len(self.__data)) + self.__data
        if self.reserved2: # Shortcut since this property has protection.
            ret += self.__reserved2

        if self.__tocSignature == 0x494E414E:
            ret += st.ST_LE_UI32.pack(len(self.__tocEntries))
            for entry in self.__tocEntries:
                ret += bytes(entry)
        else:
            ret += b'\x00\x00\x00\x00'

        return ret

    @property
    def advf(self) -> Union[int, ADVF]:
        """
        An implementation specific hint on how to render the presentation data
        on screen. May be ignored on processing.
        """
        return self.__advf

    @advf.setter
    def advf(self, val: Union[int, ADVF]) -> None:
        if not isinstance(val, int):
            raise TypeError(':property advf: must be an int.')
        if val < 0:
            raise ValueError(':property advf: must be positive.')
        if val > 0xFFFFFFFF:
            raise ValueError(':property advf: cannot be greater than 0xFFFFFFFF.')

        self.__advf = val

    @property
    def ansiClipboardFormat(self) -> ClipboardFormatOrAnsiString:
        return self.__ansiClipboardFormat

    @property
    def aspect(self) -> Union[int, DVAspect]:
        """
        An implementation specific hint on how to render the presentation data
        on screen. May be ignored on processing.
        """
        return self.__aspect

    @aspect.setter
    def aspect(self, val: Union[int, DVAspect]) -> None:
        if not isinstance(val, int):
            raise TypeError(':property aspect: must be an int.')
        if val < 0:
            raise ValueError(':property aspect: must be positive.')
        if val > 0xFFFFFFFF:
            raise ValueError(':property aspect: cannot be greater than 0xFFFFFFFF.')

        self.__aspect = val

    @property
    def data(self) -> bytes:
        """
        The presentation data. The form of this data depends on :property clipboardFormat: of :property ansiClipboardFormat:.
        """
        return self.__data

    @data.setter
    def data(self, val: bytes) -> None:
        if not isinstance(val, bytes):
            raise TypeError(':property data: must be bytes.')
        self.__data = val

    @property
    def height(self) -> int:
        """
        The height, in pixels, of the presentation data.
        """
        return self.__height

    @height.setter
    def height(self, val: int) -> None:
        if not isinstance(val, int):
            raise TypeError(':property height: must be an int.')
        if val < 0:
            raise ValueError(':property height: must be positive.')
        if val > 0xFFFFFFFF:
            raise ValueError(':property height: cannot be greater than 0xFFFFFFFF.')

        self.__height = val

    @property
    def lindex(self) -> int:
        """
        An implementation specific hint on how to render the presentation data on screen. May be ignored on processing.
        """
        return self.__lindex

    @lindex.setter
    def lindex(self, val: int) -> None:
        if not isinstance(val, int):
            raise TypeError(':property lindex: must be an int.')
        if val < 0:
            raise ValueError(':property lindex: must be positive.')
        if val > 0xFFFFFFFF:
            raise ValueError(':property lindex: cannot be greater than 0xFFFFFFFF.')

        self.__lindex = val

    @property
    def reserved1(self) -> bytes:
        """
        4 bytes that can contain any arbitrary data. Must be *exactly* 4 bytes
        when setting.
        """
        return self.__reserved1

    @reserved1.setter
    def reserved1(self, val: bytes) -> None:
        if not isinstance(val, bytes):
            raise TypeError(':property reserved1: must by bytes.')
        if len(val) != 4:
            raise ValueError(':property reserved1: must be exactly 4 bytes.')

        self.__reserved1 = val

    @property
    def reserved2(self) -> Optional[bytes]:
        """
        Optional additional data that is only set if the clipboard format of
        :property ansiClipboardFormat: is ``CF_METAFILEPICT``.

        Getting this will automatically correct the value retrieved based on
        the clipboard format, but will *not* modify the underlying data.

        Must be *exactly* 18 bytes when setting.
        """
        if self.__ansiClipboardFormat.clipboardFormat is not ClipboardFormat.CF_METAFILEPICT:
            return None
        return self.__reserved2

    @reserved2.setter
    def reserved2(self, val: bytes) -> None:
        if self.__ansiClipboardFormat.clipboardFormat is not ClipboardFormat.CF_METAFILEPICT:
            raise ValueError(':property reserved2: cannot be set if the clipboard format (from :property ansiClipboardFormat:) is not CF_METAFILEPICT.')
        if not isinstance(val, bytes):
            raise TypeError(':property reserved2: must by bytes.')
        if len(val) != 18:
            raise ValueError(':property reserved2: must be exactly 18 bytes.')

        self.__reserved2 = val

    @property
    def targetDevice(self) -> Optional[DVTargetDevice]:
        return self.__targetDevice

    @targetDevice.setter
    def targetDevice(self, val: Optional[DVTargetDevice]) -> None:
        if val is not None and not isinstance(val, DVTargetDevice):
            raise TypeError(':property targetDevice: must be None or a DVTargetDevice.')

        self.__targetDevice = val

    @property
    def tocEntries(self) -> List[TOCEntry]:
        """
        A list of TOCEntry structures. If :property tocSignature: is not set to
        ``0x494E414E``, accessing this value will clear the list.

        :returns: A direct reference to the list, allowing for modification.
            This class WILL NOT change this reference over the lifetime of the
            object.
        """
        if self.__tocSignature != 0x494E414E:
            self.__tocEntries.clear()
        return self.__tocEntries

    @property
    def tocSignature(self) -> int:
        """
        If this field does not contain ``0x494E414E``, then
        :property tocEntries: MUST be empty. Modifications to the list will be lost when it is next retrieved, meaning changes while this property is not ``0x494E414E`` WILL be lost.

        Setting this to a value other than ``0x494E414E`` will clear the list
        immediately.
        """
        return self.__tocSignature

    @tocSignature.setter
    def tocSignature(self, val: int) -> None:
        if not isinstance(val, int):
            raise TypeError(':property tocSignature: must be an int.')
        if val < 0:
            raise ValueError(':property tocSignature: must be positive.')
        if val > 0xFFFFFFFF:
            raise ValueError(':property tocSignature: cannot be greater than 0xFFFFFFFF.')

        if val != 0x494E414E:
            self.__tocEntries.clear()

        self.__tocSignature = val

    @property
    def width(self) -> int:
        """
        The width, in pixels, of the presentation data.
        """
        return self.__width

    @width.setter
    def width(self, val: int) -> None:
        if not isinstance(val, int):
            raise TypeError(':property width: must be an int.')
        if val < 0:
            raise ValueError(':property width: must be positive.')
        if val > 0xFFFFFFFF:
            raise ValueError(':property width: cannot be greater than 0xFFFFFFFF.')

        self.__width = val


