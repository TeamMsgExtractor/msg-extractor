__all__ = [
    'Cf',
    \
    'ODTStruct',
]

import struct
from typing import Optional

from ..enums import ODTCf, ODTPersist1, ODTPersist2


class ODTStruct:
    def __init__(self, data : Optional[bytes]):
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
    def setter(self, value : ODTCf) -> None:
        if not isinstance(value, ODTCf):
            raise TypeError(':property cf: MUST be of type ODTCf.')

        self.__cf = value

    @property
    def odtPersist1(self) -> ODTPersist1:
        """
        Flags the specify information about the OLE object.
        """
        return self.__persist1

    @odtPersist1.setter
    def setter(self, value : ODTPersist1) -> None:
        if not isinstance(value, ODTPersist1):
            raise TypeError(':property odtPersist1: MUST be of type ODTPersist1.')

        self.__persist1 = value

    @property
    def odtPersist2(self) -> ODTPersist1:
        """
        Flags the specify additional information about the OLE object.
        """
        return self.__persist2

    @odtPersist2.setter
    def setter(self, value : ODTPersist2) -> None:
        if not isinstance(value, ODTPersist2):
            raise TypeError(':property odtPersist2: MUST be of type ODTPersist2.')

        self.__persist2 = value
