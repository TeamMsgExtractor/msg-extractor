__all__ = [
    'TOCEntry',
]


from typing import Optional, Union

from ._helpers import BytesReader
from .cfoas import ClipboardFormatOrAnsiString
from ..constants import st
from .dv_target_device import DVTargetDevice
from ..enums import ADVF, DVAspect


class TOCEntry:
    def __init__(self, reader: Optional[Union[bytes, BytesReader]] = None):
        if reader is None:
            self.__clipFormat = ClipboardFormatOrAnsiString()
            self.__aspect = 0
            self.__lindex = 0
            self.__tymed = 0
            self.__advf = 0
            self.__targetDevice = DVTargetDevice()
            return

        if isinstance(reader, bytes):
            reader = BytesReader(reader)

        self.__clipFormat = ClipboardFormatOrAnsiString(reader)
        targetDeviceSize = reader.readUnsignedInt()
        self.__aspect = reader.readUnsignedInt()
        self.__lindex = reader.readUnsignedInt()
        self.__tymed = reader.readUnsignedInt()
        reader.read(12)
        self.__advf = reader.readUnsignedInt()

        # Based off the wording of the documentation, it seems like this can't
        # actually be 0 bytes, so this should be fine.
        self.__targetDevice = DVTargetDevice(reader.read(targetDeviceSize))


    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        ret = bytes(self.__clipFormat)
        td = bytes(self.__targetDevice)
        ret += st.ST_LE_UI32.pack(len(td))
        ret += st.ST_LE_UI32.pack(self.__aspect)
        ret += st.ST_LE_UI32.pack(self.__lindex)
        ret += st.ST_LE_UI32.pack(self.__tymed)
        ret += b'\x00' * 12
        ret += st.ST_LE_UI32.pack(self.__advf)
        ret += b'\x00' * 4
        ret += td

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
        return self.__clipFormat

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
    def lindex(self) -> int:
        """
        An implementation specific hint on how to render the presentation data
        on screen. May be ignored on processing.
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
    def targetDevice(self) -> DVTargetDevice:
        return self.__targetDevice

    @property
    def tymed(self) -> int:
        return self.__tymed

    @tymed.setter
    def tymed(self, val: int) -> None:
        if not isinstance(val, int):
            raise TypeError(':property lindex: must be an int.')
        if val < 0:
            raise ValueError(':property lindex: must be positive.')
        if val > 0xFFFFFFFF:
            raise ValueError(':property lindex: cannot be greater than 0xFFFFFFFF.')

        self.__tymed = val
