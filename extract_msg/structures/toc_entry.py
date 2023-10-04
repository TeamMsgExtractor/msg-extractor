__all__ = [
    'TOCEntry',
]


from typing import Optional, Union

from ._helpers import BytesReader
from .cfoas import ClipboardFormatOrAnsiString
from .dv_target_device import DVTargetDevice


class TOCEntry:
    def __init__(self, reader : Union[bytes, BytesReader]):
        if isinstance(reader, bytes):
            reader = BytesReader(reader)
            self.__clipFormat = ClipboardFormatOrAnsiString(reader)
            targetDeviceSize = reader.readUnsignedInt()
            self.__aspect = reader.readUnsignedInt()
            self.__lindex = reader.readUnsignedInt()
            self.__tymed = reader.readUnsignedInt()
            reader.read(4)
            self.__advf = reader.readUnsignedInt()
            reader.read(4)
            if targetDeviceSize == 0:
                self.__targetDevice = None
            else:
                self.__targetDevice = DVTargetDevice(reader.read(targetDeviceSize))
        # TODO

    def toBytes(self) -> bytes:
        ret = self.__clipFormat.toBytes()
        # TODO

        return ret