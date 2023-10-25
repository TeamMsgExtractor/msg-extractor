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
    def __init__(self, reader : Union[bytes, BytesReader]):
        if isinstance(reader, bytes):
            reader = BytesReader(reader)
            self.__clipFormat = ClipboardFormatOrAnsiString(reader)
            targetDeviceSize = reader.readUnsignedInt()
            self.__aspect = reader.readUnsignedInt()
            self.__lindex = reader.readUnsignedInt()
            self.__tymed = reader.readUnsignedInt()
            reader.read(12)
            self.__advf = reader.readUnsignedInt()
            reader.read(4)
            if targetDeviceSize == 0:
                self.__targetDevice = None
            else:
                self.__targetDevice = DVTargetDevice(reader.read(targetDeviceSize))
        # TODO

    def toBytes(self) -> bytes:
        ret = self.__clipFormat.toBytes()
        td = self.__targetDevice.toBytes() if self.__targetDevice else b''
        ret += st.ST_LE_UI32.pack(len(td))
        ret += st.ST_LE_UI32.pack(self.__aspect)
        ret += st.ST_LE_UI32.pack(self.__lindex)
        ret += st.ST_LE_UI32.pack(self.__tymed)
        #TODO

        return ret

    @property
    def advf(self) -> Union[int, ADVF]:
        """
        An implementation specific hint on how to render the presentation data 
        on screen. May be ignored on processing.
        """
        return self.__advf

    @advf.setter
    def _(self, val : Union[int, ADVF]) -> None:
        if not isinstance(val, int):
            raise TypeError(':property advf: must be an int.')
        if val < 0:
            raise ValueError(':property advf: must be positive.')
        if val > 4294967295:
            raise ValueError(':property advf: cannot be greater than 4294967295.')

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
    def _(self, val : Union[int, DVAspect]) -> None:
        if not isinstance(val, int):
            raise TypeError(':property aspect: must be an int.')
        if val < 0:
            raise ValueError(':property aspect: must be positive.')
        if val > 4294967295:
            raise ValueError(':property aspect: cannot be greater than 4294967295.')

        self.__aspect = val

    