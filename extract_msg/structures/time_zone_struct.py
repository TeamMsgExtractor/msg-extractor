__all__ = [
    'TimeZoneStruct',
]


from .. import constants
from .system_time import SystemTime


class TimeZoneStruct:
    """
    A TimeZoneStruct, as specified in [MS-OXOCAL].
    """

    def __init__(self, data : bytes):
        self.__rawData = data
        unpacked = constants.st.ST_TZ.unpack(data)
        self.__bias = unpacked[0]
        self.__standardBias = unpacked[1]
        self.__daylightBias = unpacked[2]
        self.__standardYear = unpacked[3]
        self.__standardDate = SystemTime(unpacked[4])
        self.__daylightYear = unpacked[5]
        self.__daylightDate = SystemTime(unpacked[6])

    @property
    def bias(self) -> int:
        """
        The time zone's offset in minutes from UTC.
        """
        return self.__bias

    @property
    def daylightBias(self) -> int:
        """
        The offset in minutes from the value of the bias field during daylight
        saving time.
        """
        return self.__daylightBias

    @property
    def daylightDate(self) -> SystemTime:
        """
        The date and local time that indicate when to begin using the value
        specified in the daylightBias field. Uses the same format as
        standardDate.
        """
        return self.__daylightDate

    @property
    def daylightYear(self) -> int:
        """
        The value of the daylightDate field's year.
        """
        return self.__daylightYear

    @property
    def rawData(self) -> bytes:
        """
        The raw bytes used to create this object.
        """
        return self.__rawData

    @property
    def standardBias(self) -> int:
        """
        The offset in minutes from the value of the bias field during standard
        time.
        """
        return self.__standardBias

    @property
    def standardDate(self) -> SystemTime:
        """
        The date and local time that indicate when to begin using the value
        specified in the standardBias field. If the time zone does not support
        daylight's savings time, the month member must be 0. If the year is not
        0, then it is an absolute date than only occurs once, otherwise it is a
        relative date that occurs yearly.

        See [MS-OXOCAL] for details.
        """
        return self.__standardDate

    @property
    def standardYear(self) -> int:
        """
        The value of the standardDate field's year.
        """
        return self.__standardYear
