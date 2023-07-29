__all__ = [
    'SystemTime',
]


from .. import constants


class SystemTime:
    """
    A SYSTEMTIME struct, as defined in [MS-DTYP].
    """

    year : int = None
    month : int = None
    dayOfWeek : int = None
    day : int = None
    hour : int = None
    minute : int = None
    second : int = None
    milliseconds : int = None

    def __init__(self, data : bytes):
        self.unpack(data)

    def __eq__(self, other):
        return self.pack() == other.pack()

    def __ne__(self, other):
        return not self.__eq__(other)

    def pack(self) -> bytes:
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
