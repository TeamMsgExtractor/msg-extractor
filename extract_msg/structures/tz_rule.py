__all__ = [
    'TZRule',
]


import logging

from struct import Struct
from typing import Final, final, Optional

from ..enums import TZFlag
from ._helpers import BytesReader
from .system_time import SystemTime


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


@final
class TZRule:
    """
    A TZRule structure, as defined in [MS-OXOCAL].
    """

    __SIZE__: int = 66
    __struct: Final[Struct] = Struct('4B2H14x3i16s16s')

    def __init__(self, data: Optional[bytes] = None):
        if not data:
            self.__majorVersion = 2
            self.__minorVersion = 1
            self.__flags = TZFlag(0)
            self.__year = 0
            self.__bias = 0
            self.__standardBias = 0
            self.__daylightBias = 0
            self.__standardDate = SystemTime()
            self.__daylightDate = SystemTime()
            return

        reader = BytesReader(data)
        self.__majorVersion = reader.readUnsignedByte()
        self.__minorVersion = reader.readUnsignedByte()
        reader.assertRead(b'\x3E\x00')
        self.__flags = TZFlag(reader.readUnsignedShort())
        self.__year = reader.readUnsignedShort()
        # This *MUST* be null, however I've seen Outlook not follow that. Simply
        # log a warning about it even though it's a violation.
        if any(b := reader.read(14)):
            logger.warning(f'Read TZRule with non-null X section (got {b}).')
        self.__bias = reader.readInt()
        self.__standardBias = reader.readInt()
        self.__daylightBias = reader.readInt()
        self.__standardDate = SystemTime(reader.read(16))
        self.__daylightDate = SystemTime(reader.read(16))

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        return self.__struct.pack(self.__majorVersion,
                                  self.__minorVersion,
                                  62,
                                  0,
                                  self.__flags,
                                  self.__year,
                                  self.__bias,
                                  self.__standardBias,
                                  self.__daylightBias,
                                  bytes(self.__standardDate),
                                  bytes(self.__daylightDate))

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
    def flags(self) -> TZFlag:
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
