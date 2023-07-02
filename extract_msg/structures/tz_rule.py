__all__ = [
    'TZRule',
]


from typing import Set

from .. import constants
from ..enums import TZFlag
from ._helpers import BytesReader
from .system_time import SystemTime


class TZRule:
    """
    A TZRule structure, as defined in [MS-OXOCAL].
    """

    __SIZE__ : int = 66

    def __init__(self, data : bytes):
        self.__rawData = data
        reader = BytesReader(data)
        self.__majorVersion = reader.readByte()
        self.__minorVersion = reader.readByte()
        reader.assertRead(b'\x3E\x00')
        self.__flags = TZFlag(reader.readUnsignedShort())
        self.__year = reader.readShort()
        # We *should* be doing this, but Outlook is violating the standard so...
        #reader.assertNull(14)
        self.__bias = reader.readInt()
        self.__standardBias = reader.readInt()
        self.__daylightBias = reader.readInt()
        self.__standardDate = SystemTime(reader.read(16))
        self.__daylightDate = SystemTime(reader.read(16))

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
    def flags(self) -> Set[TZFlag]:
        """
        The flags for this rule.
        """
        return self.__flags

    @property
    def majorVersion(self) -> int:
        """
        The major version.
        """
        return self.__majorVersion

    @property
    def minorVersion(self) -> int:
        """
        The minor version.
        """
        return self.__minorVersion

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
    def year(self) -> int:
        """
        The year this rule is scheduled to take place. A rule will remain in
        effect from January 1 of it's year until January 1 of the next rule's
        year field. If no rules exist for subsequent years, this rule will
        remain in effect indefinately.
        """
        return self.__year
