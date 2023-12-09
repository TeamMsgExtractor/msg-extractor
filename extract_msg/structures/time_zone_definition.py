__all__ = [
    'TimeZoneDefinition',
]


from typing import List, Optional

from ..constants import st
from ._helpers import BytesReader
from .tz_rule import TZRule


class TimeZoneDefinition:
    """
    Structure for PidLidAppointmentTimeZoneDefinitionRecur from [MS-OXOCAL].
    """

    def __init__(self, data: Optional[bytes] = None):
        if not data:
            self.__majorVersion = 2
            self.__minorVersion = 1
            self.__keyName = ''
            self.__rules = [TZRule()]
            return
        reader = BytesReader(data)
        self.__majorVersion = reader.readUnsignedByte()
        self.__minorVersion = reader.readUnsignedByte()
        cbHeader = reader.readUnsignedShort()
        reader.assertRead(b'\x02\x00')
        cchKeyName = reader.readUnsignedShort()
        self.__keyName = reader.read(2 * cchKeyName).decode('utf-16-le')
        cRules = reader.readUnsignedShort()
        if cRules < 1 or cRules > 1024:
            raise ValueError('Value for cRules was out of range.')
        self.__rules = [reader.readClass(TZRule) for _ in range(cRules)]

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        # Validate some of the data.
        if len(self.__rules) < 1:
            raise ValueError('Cannot pack a TimeZoneDefinition with no rules.')
        if len(self.__rules) > 1024:
            raise ValueError('TimeZoneDefintion can only have up to 1024 rules.')

        ret = bytes((self.__majorVersion, self.__minorVersion))
        ret += st.ST_LE_UI16.pack(6 + 2 * len(self.__keyName))
        ret += b'\x02\x00'
        ret += st.ST_LE_UI16.pack(2* len(self.__keyName))
        ret += st.ST_LE_UI16.pack(len(self.__rules))
        ret += b''.join(bytes(x) for x in self.__rules)

        return ret

    @property
    def keyName(self) -> str:
        """
        The name of the associated time zone.

        Not localized but instead set to the unique name of the desired time
        zone.
        """
        return self.__keyName

    @keyName.setter
    def keyName(self, value: str) -> None:
        value = str(value)
        if len(value) > 260:
            raise ValueError('Key name must be a string less than 261 characters.')

        self.__keyName = value

    @property
    def majorVersion(self) -> int:
        """
        The major version.
        """
        return self.__majorVersion

    @majorVersion.setter
    def majorVersion(self, value: int) -> None:
        if value > 255:
            raise ValueError('Major version cannot be greater than 255')
        if value < 0:
            raise ValueError('Major version must be positive.')

        self.__minorVersion = value

    @property
    def minorVersion(self) -> int:
        """
        The minor version.
        """
        return self.__minorVersion

    @minorVersion.setter
    def minorVersion(self, value: int) -> None:
        if value > 255:
            raise ValueError('Minor version cannot be greater than 255')
        if value < 0:
            raise ValueError('Minor version must be positive.')

        self.__minorVersion = value

    @property
    def rules(self) -> List[TZRule]:
        """
        A tuple of TZRule structures that specifies a time zone.
        """
        return self.__rules
