__all__ = [
    'OleStreamStruct',
]


from typing import Optional

from ._helpers import BytesReader
from .mon_stream import MonikerStream


class OleStreamStruct:
    def __init__(self, data : Optional[bytes] = None):
        self.__reservedMonikerStream = None
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
            self.__reservedMonikerStream = MonikerStream(reader.read(rmsSize))

        # TODO implement the rest. It's all optional things.

    def toBytes(self) -> bytes:
        pass # TODO
        

