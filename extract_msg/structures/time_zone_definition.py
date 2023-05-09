__all__ = [
    'TimeZoneDefinition',
]


from typing import Tuple

from ._helpers import BytesReader
from .tz_rule import TZRule


class TimeZoneDefinition:
    """
    Structure for PidLidAppointmentTimeZoneDefinitionRecur from [MS-OXOCAL].
    """

    def __init__(self, data : bytes):
        self.__rawData = data
        reader = BytesReader(data)
        self.__majorVersion = reader.readUnsignedByte()
        self.__minorVersion = reader.readUnsignedByte()
        cbHeader = reader.readUnsignedShort()
        reader.assertRead(b'\x02\x00')
        cchKeyName = reader.readUnsignedShort()
        self.__keyName = reader.read(2 * cchKeyName).decode('utf-16-le')
        cRules = reader.readUnsignedShort()
        self.__rules = tuple(reader.readClass(TZRule) for x in range(cRules))

    @property
    def keyName(self) -> str:
        """
        The name of the associated time zone. Not localized but instead set to
        the unique name of the desired time zone.
        """
        return self.__keyName

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
    def rules(self) -> Tuple[TZRule]:
        """
        A tuple of TZRule structures that specifies a time zone.
        """
        return self.__rules
