__all__ = [
    'OleStreamStruct',
]


from typing import final, Optional

from ..constants import st
from ._helpers import BytesReader
from .mon_stream import MonikerStream


@final
class OleStreamStruct:
    def __init__(self, data : Optional[bytes] = None):
        self.__rms = None
        if not data:
            self.__flags = 0
            self.__linkUpdateOption = 0
            return
        reader = BytesReader(data)
        # Assert the version.
        reader.assertRead(b'\x01\x00\x00\x02', 'Ole stream had invalid version (expected {expected}, got {actual}).')
        self.__flags = reader.readUnsignedInt()
        self.__linkUpdateOption = reader.readUnsignedInt()
        reader.assertNull(4, 'Ole stream reserved was not null (got {actual}).')
        rmsSize = reader.readUnsignedInt()
        if rmsSize > 0:
            self.__rms = MonikerStream(reader.read(rmsSize))

        if self.__flags & 1:
            # Only check this stuff if this is not for an embedded object.
            pass # TODO

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        ret = b'\x01\x00\x00\x02'
        ret += st.ST_LE_UI32.pack(self.__flags)
        ret += st.ST_LE_UI32.pack(self.__linkUpdateOption)
        ret += b'\x00\x00\x00\x00'
        rmsBytes = b'' if self.__rms is None else bytes(self.__rms)
        ret += st.ST_LE_UI32.pack(len(rmsBytes)) + rmsBytes
        # TODO finish this with the optional properties.

        return ret

    @property
    def reservedMonikerStream(self) -> Optional[MonikerStream]:
        """
        A MonikerStream structure that can contain any arbitrary value.
        """
        return self.__rms

    @reservedMonikerStream.setter
    def _(self, data : Optional[MonikerStream]) -> None:
        if data is not None and not isinstance(data, MonikerStream):
            raise TypeError('Reserved moniker stream must be a MonikerStream instance or None.')

        self.__rms = data


