__all__ = [
    'ODTStruct',
]

import struct
from typing import final, Optional

from ..enums import ODTCf, ODTPersist1, ODTPersist2


@final
class ODTStruct:
    def __init__(self, data: Optional[bytes] = None):
        if data:
            values = struct.unpack('<HH', data[:8])
            self.__cf = ODTCf(values[0])
            self.__persist1 = ODTPersist1(values[1])
            if len(data) >= 6:
                self.__persist2 = ODTPersist2(struct.unpack('<H', data[4:6])[0])
            else:
                self.__persist2 = ODTPersist2.NONE
        else:
            self.__cf = ODTCf.UNSPECIFIED
            self.__persist1 = ODTPersist1.NONE
            self.__persist2 = ODTPersist2.NONE

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        return struct.pack('<HHH', self.__cf, self.__persist1, self.__persist2)

    @property
    def cf(self) -> ODTCf:
        """
        An enum value that specifies the format this OLE object uses to
        transmit data to the host application.
        """
        return self.__cf

    @cf.setter
    def cf(self, value: ODTCf) -> None:
        if not isinstance(value, ODTCf):
            raise TypeError(':property cf: MUST be of type ODTCf.')

        self.__cf = value

    @property
    def odtPersist1(self) -> ODTPersist1:
        """
        Flags that specify information about the OLE object.
        """
        return self.__persist1

    @odtPersist1.setter
    def odtPersist1(self, value: ODTPersist1) -> None:
        if not isinstance(value, ODTPersist1):
            raise TypeError(':property odtPersist1: MUST be of type ODTPersist1.')

        self.__persist1 = value

    @property
    def odtPersist2(self) -> ODTPersist2:
        """
        Flags that specify additional information about the OLE object.
        """
        return self.__persist2

    @odtPersist2.setter
    def odtPersist2(self, value: ODTPersist2) -> None:
        if not isinstance(value, ODTPersist2):
            raise TypeError(':property odtPersist2: MUST be of type ODTPersist2.')

        self.__persist2 = value
