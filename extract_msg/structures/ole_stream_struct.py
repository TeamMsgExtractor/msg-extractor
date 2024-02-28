__all__ = [
    'OleStreamStruct',
]


from typing import final, Optional

from ..constants import st
from ._helpers import BytesReader
from .mon_stream import MonikerStream


@final
class OleStreamStruct:
    """
    The OLEStream structure, as specified in [MS-OLEDS].

    Specifically, this is *only* the version that is used for embedded objects.
    As such, only some of the fields are ever present.
    """

    def __init__(self, data: Optional[bytes] = None):
        """
        :raises TypeError: The data given is not for an embedded object.
        """
        self.__rms = None
        if not data:
            self.__flags = 0
            self.__linkUpdateOption = 0
            return
        reader = BytesReader(data)
        # Assert the version.
        reader.assertRead(b'\x01\x00\x00\x02', 'Ole stream had invalid version (expected {expected}, got {actual}).')
        self.__flags = reader.readUnsignedInt()
        if self.__flags & 1:
            raise TypeError('Cannot parse an OLEStream structure for a linked object.')
        self.__linkUpdateOption = reader.readUnsignedInt()
        reader.assertNull(4, 'Ole stream reserved was not null (got {actual}).')
        rmsSize = reader.readUnsignedInt()
        if rmsSize > 0:
            self.__rms = MonikerStream(reader.read(rmsSize - 4))

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        ret = b'\x01\x00\x00\x02'
        ret += st.ST_LE_UI32.pack(self.__flags)
        ret += st.ST_LE_UI32.pack(self.__linkUpdateOption)
        ret += b'\x00\x00\x00\x00'
        rmsBytes = b'' if self.__rms is None else bytes(self.__rms)
        rmsLen = (len(rmsBytes) + 4) if rmsBytes else 0
        ret += st.ST_LE_UI32.pack(rmsLen) + rmsBytes

        return ret

    @property
    def flags(self) -> int:
        """
        The flags for the OLEStream.

        The bit with mask ``0x00001000`` is an implementation-specific hint
        supplied by the application or by a higher-level protocol that creates
        the data structure. It MAY be ignored on processing. A server
        implementation which does not ignore this bit MAY cache the storage
        when the bit is set.

        :raises ValueError: The property was set with a bit other than the
            implementation specific bit set.
        """
        return self.__flags

    @flags.setter
    def flags(self, value: int) -> None:
        if not isinstance(value, int):
            raise TypeError(':property flags: MUST be an int.')
        if value != 0 and value != 0x1000:
            raise ValueError('Cannot set bits other than the implementation specific one.')

        self.__flags = value

    @property
    def linkUpdateOption(self) -> int:
        """
        An implementation-specific hint.

        This hint MAY be ignored. On Windows, this field contains values from
        the OLEUPDATE enumeration.
        """
        return self.__linkUpdateOption

    @linkUpdateOption.setter
    def linkUpdateOption(self, value: int) -> None:
        if not isinstance(value, int):
            raise TypeError(':property linkUpdateOption: MUST be an int.')
        if value < 0:
            raise ValueError(':property linkUpdateOption: MUST be positive.')
        if value > 0xFFFFFFFF:
            raise ValueError(':property linkUpdateOption: MUST be less than 0x100000000.')

        self.__linkUpdateOption = value

    @property
    def reservedMonikerStream(self) -> Optional[MonikerStream]:
        """
        A MonikerStream structure that can contain any arbitrary value.
        """
        return self.__rms

    @reservedMonikerStream.setter
    def reservedMonikerStream(self, data: Optional[MonikerStream]) -> None:
        if data is not None and not isinstance(data, MonikerStream):
            raise TypeError('Reserved moniker stream must be a MonikerStream instance or None.')

        self.__rms = data


