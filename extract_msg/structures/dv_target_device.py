__all__ = [
    'DVTargetDevice',
]


import struct

from typing import Optional

from ._helpers import BytesReader
from .dev_mode_a import DevModeA


class DVTargetDevice:
    """
    Specifies information about a device that renders the presentation data.

    The creator of this data structure MUST NOT assume that it will be
    understood during processing.
    """

    def __init__(self, data: Optional[bytes]):
        self.__driverName = None
        self.__deviceName = None
        self.__portName = None
        self.__extDevMode = None

        if not data:
            return
        reader = BytesReader(data)

        # We have 4 fields to read, and *technically* they may not all even be
        # present, given that this structure can be 4 bytes? Reading all of
        # these is also much more complicated than other structures, as they can
        # technically overlap. We are just going to be *much* more lenient about
        # this structure.
        offset1 = offset2 = offset3 = offset4 = -1
        try:
            offset1 = reader.readUnsignedShort()
            offset2 = reader.readUnsignedShort()
            offset3 = reader.readUnsignedShort()
            offset4 = reader.readUnsignedShort()
        except IOError:
            pass

        if offset1 != -1 and offset1 < len(data):
            reader.seek(offset1)
            try:
                self.__driverName = reader.readByteString()
            except IOError:
                self.__driverName = reader.read()
                if not self.__driverName:
                    self.__driverName = None

        if offset2 != -1 and offset2 < len(data):
            reader.seek(offset2)
            try:
                self.__deviceName = reader.readByteString()
            except IOError:
                self.__deviceName = reader.read()
                if not self.__deviceName:
                    self.__deviceName = None

        if offset3 != -1 and offset3 < len(data):
            reader.seek(offset3)
            try:
                self.__portName = reader.readByteString()
            except IOError:
                self.__portName = reader.read()
                if not self.__portName:
                    self.__portName = None

        if offset4 != -1 and offset4 < len(data):
            reader.seek(offset4)
            try:
                devmode = DevModeA(reader.read(56))
                if devmode:
                    self.__extDevMode = devmode
            except IOError:
                self.__extDevMode = None

    def __bytes__(self) -> bytes:
        ret = self.toBytes()
        if not isinstance(ret, bytes):
            raise TypeError(f'Cannot convert {self.__class__.__name__} instance to bytes.')
        return ret

    def toBytes(self) -> Optional[bytes]:
        if not (self.driverName or self.deviceName or self.portName or self.extDevMode):
            return None
        currentPosition = 8

        offset1 = 8 if self.__driverName else 0
        if self.__driverName:
            currentPosition += len(self.__driverName) + 1

        offset2 = currentPosition if self.__deviceName else 0
        if self.__deviceName:
            currentPosition += len(self.__deviceName) + 1

        offset3 = currentPosition if self.__portName else 0
        if self.__portName:
            currentPosition += len(self.__portName) + 1

        extDevModeBytes = bytes(self.__extDevMode) if self.__extDevMode else None
        offset4 = currentPosition if extDevModeBytes else 0

        try:
            ret = struct.pack('<HHHH', offset1, offset2, offset3, offset4)
        except struct.error:
            raise ValueError('DVTargetDevice structure contains too much data.')
        if self.__driverName:
            ret += self.__driverName + b'\x00'
        if self.__deviceName:
            ret += self.__deviceName + b'\x00'
        if self.__portName:
            ret += self.__portName + b'\x00'
        if extDevModeBytes:
            ret += extDevModeBytes

        return ret

    @property
    def driverName(self) -> Optional[bytes]:
        """
        Optional ANSI string that contains a hunt on how to display or print
        presentation data.
        """
        return self.__driverName

    @driverName.setter
    def driverName(self, data: Optional[bytes]) -> None:
        self.__driverName = None if not data else data

    @property
    def deviceName(self) -> Optional[bytes]:
        """
        Optional ANSI string that contains a hunt on how to display or print
        presentation data.
        """
        return self.__deviceName

    @deviceName.setter
    def deviceName(self, data: Optional[bytes]) -> None:
        self.__deviceName = None if not data else data

    @property
    def portName(self) -> Optional[bytes]:
        """
        Optional ANSI string that contains any arbitrary value.
        """
        return self.__portName

    @portName.setter
    def portName(self, data: Optional[bytes]) -> None:
        self.__portName = None if not data else data

    @property
    def extDevMode(self) -> Optional[DevModeA]:
        """
        Optional ANSI string that contains a hunt on how to display or print
        presentation data.
        """
        return self.__extDevMode

    @extDevMode.setter
    def extDevMode(self, data: Optional[DevModeA]) -> None:
        self.__extDevMode = None if not data else data