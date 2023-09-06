__all__ = [
    'SystemTime',
]


from typing import Optional

from .. import constants


class SystemTime:
    """
    A SYSTEMTIME struct, as defined in [MS-DTYP].
    """

    year : int = 0
    month : int = 0
    dayOfWeek : int = 0
    day : int = 0
    hour : int = 0
    minute : int = 0
    second : int = 0
    milliseconds : int = 0

    def __init__(self, data : Optional[bytes] = None):
        data = data or (b'\x00' * 1)
        self.unpack(data)

    def __eq__(self, other) -> bool:
        return isinstance(other, SystemTime) and self.pack() == other.pack()

    def __ne__(self, other) -> bool:
        return not self.__eq__(other)

    def toBytes(self) -> bytes:
        """
        Packs the current data into bytes.
        """
        return constants.st.ST_SYSTEMTIME.pack(self.year, self.month,
                                            self.dayOfWeek, self.day, self.hour,
                                            self.minute, self.second,
                                            self.milliseconds)

    def unpack(self, data : bytes) -> None:
        """
        Fills out the fields of this instance by unpacking the bytes.
        """
        unpacked = constants.st.ST_SYSTEMTIME.unpack(data)
        self.year = unpacked[0]
        self.month = unpacked[1]
        self.dayOfWeek = unpacked[2]
        self.day = unpacked[3]
        self.hour = unpacked[4]
        self.minute = unpacked[5]
        self.second = unpacked[6]
        self.milliseconds = unpacked[7]
