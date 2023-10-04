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
    __ansiClipboardFormat : ClipboardFormatOrAnsiString
    __targetDevice : Optional[DVTargetDevice]
    __aspect : Union[int, DVAspect]
    __lindex : int
    __advf : Union[int, ADVF]
    __width : int
    __height : int
    __data : bytes
    __reserved2 : Optional[bytes]
    __tocSignature : int
    __tocEntries : List[TOCEntry]

    def __init__(self, data : bytes):
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
            self.__reserved2 = None

        self.__tocSignature = reader.readUnsignedInt()
        self.__tocEntries = []
        if self.__tocSignature == 0x494E414E: # b'NANI' in little endian.
            for _ in range(reader.readUnsignedInt()):
                self.__tocEntries.append(TOCEntry(reader))


    def toBytes(self) -> bytes:
        ret = self.__ansiClipboardFormat.toBytes()

        if self.__targetDevice is None:
            ret += b'\x04\x00\x00\x00'
        else:
            dvData = self.__targetDevice.toBytes()
            ret += st.ST_LE_UI32.pack(len(dvData))
            ret += dvData

        ret += st.ST_LE_UI32.pack(self.__aspect)
        ret += st.ST_LE_UI32.pack(self.__lindex)
        ret += st.ST_LE_UI32.pack(self.__advf)
        ret += self.__reserved1
        ret += st.ST_LE_UI32.pack(self.__width)
        ret += st.ST_LE_UI32.pack(self.__height)
        ret += st.ST_LE_UI32.pack(len(self.__data)) + self.__data
        if self.__reserved2: # Shortcut since this property has protection.
            ret += self.__reserved2

        if self.__tocSignature == 0x494E414E:
            ret += st.ST_LE_UI32.pack(len(self.__tocEntries))
            for entry in self.__tocEntries:
                ret += entry.toBytes()
        else:
            ret += b'\x00\x00\x00\x00'

        return ret

    @property
    def advf(self) -> Union[int, ADVF]:
        """

        """
        return self.__advf

    @advf.setter
    def _(self, val : Union[int, ADVF]) -> None:
        if not isinstance(val, int):
            raise TypeError('advf must be an int.')
        if val < 0:
            raise ValueError('advf must be positive.')
        if val > 4294967295:
            raise ValueError('advf cannot be greater than 4294967295.')

        self.__advf = val

    @property
    def ansiClipboardFormat(self) -> ClipboardFormatOrAnsiString:
        """

        """
        return self.__ansiClipboardFormat

    @property
    def aspect(self) -> Union[int, DVAspect]:
        """

        """
        return self.__aspect

    @aspect.setter
    def _(self, val : Union[int, DVAspect]) -> None:
        if not isinstance(val, int):
            raise TypeError(':property aspect: must be an int.')
        if val < 0:
            raise ValueError(':property aspect: must be positive.')
        if val > 4294967295:
            raise ValueError(':property aspect: cannot be greater than 4294967295.')

        self.__aspect = val

    @property
    def data(self) -> bytes:
        """

        """
        return self.__data

    @data.setter
    def _(self, val : bytes) -> None:
        if not isinstance(val, bytes):
            raise TypeError(':property data: must be bytes.')
        self.__data = val

    @property
    def height(self) -> int:
        """

        """
        return self.__height

    @height.setter
    def _(self, val : int) -> None:
        if val < 0:
            raise ValueError(':property height: must be positive.')
        if val > 4294967295:
            raise ValueError(':property height: cannot be greater than 4294967295.')

        self.__height = val

    @property
    def lindex(self) -> int:
        """

        """
        return self.__lindex

    @lindex.setter
    def _(self, val : int) -> None:
        if val < 0:
            raise ValueError(':property lindex: must be positive.')
        if val > 4294967295:
            raise ValueError(':property lindex: cannot be greater than 4294967295.')

        self.__lindex = val

    @property
    def reserved1(self) -> bytes:
        """

        """
        return self.__reserved1

    @reserved1.setter
    def _(self, val : bytes) -> None:
        if not isinstance(val, bytes):
            raise TypeError(':property reserved1: must by bytes.')
        if len(val) != 4:
            raise ValueError(':property reserved1: must be exactly 4 bytes.')

        self.__reserved1 = val

    @property
    def reserved2(self) -> Optional[bytes]:
        """

        """
        return self.__reserved2

    @reserved2.setter
    def _(self, val : bytes) -> None:
        if not isinstance(val, bytes):
            raise TypeError(':property reserved2: must by bytes.')
        if len(val) != 18:
            raise ValueError(':property reserved2: must be exactly 18 bytes.')

        self.__reserved2 = val

    @property
    def targetDevice(self) -> Optional[DVTargetDevice]:
        """

        """
        return self.__targetDevice

    @targetDevice.setter
    def _(self, val : Optional[DVTargetDevice]) -> None:
        if val is not None and not isinstance(val, DVTargetDevice):
            raise TypeError(':property targetDevice: must be None or a DVTargetDevice.')

        self.__targetDevice = val

    @property
    def tocEntries(self) -> List[TOCEntry]:
        """

        """
        return self.__tocEntries

    @property
    def tocSignature(self) -> int:
        """

        """
        return self.__tocSignature

    @tocSignature.setter
    def _(self, val : int) -> None:
        if val < 0:
            raise ValueError('tocSignature must be positive.')
        if val > 4294967295:
            raise ValueError('tocSignature cannot be greater than 4294967295.')

        self.__tocSignature = val

    @property
    def width(self) -> int:
        """

        """
        return self.__width

    @width.setter
    def _(self, val : int) -> None:
        if val < 0:
            raise ValueError('width must be positive.')
        if val > 4294967295:
            raise ValueError('width cannot be greater than 4294967295.')

        self.__width = val


