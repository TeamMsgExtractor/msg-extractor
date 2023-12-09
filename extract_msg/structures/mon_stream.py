__all__ = [
    'MonikerStream',
]


from typing import final, Optional


@final
class MonikerStream:
    def __init__(self, data: Optional[bytes] = None):
        if data:
            self.__clsid = data[:16]
            self.__streamData = data[16:]
        else:
            self.__clsid = b'\x00' * 16
            self.__streamData = b''

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        return self.__clsid + self.__streamData

    @property
    def clsid(self) -> bytes:
        """
        The CLSID, as a stream of 16 bytes, of an implementation specific object
        capable of processing the stream data.
        """
        return self.__clsid

    @clsid.setter
    def clsid(self, data: bytes) -> None:
        if not isinstance(data, bytes):
            raise TypeError('CLSID MUST be bytes.')
        if len(data) != 16:
            raise ValueError('CLSID MUST be 16 bytes.')

        self.__clsid = data

    @property
    def streamData(self) -> bytes:
        """
        An array of bytes that specifies the reference to the linked object.
        """
        return self.__streamData

    @streamData.setter
    def streamData(self, data: bytes) -> None:
        if not isinstance(data, bytes):
            raise TypeError('Stream data MUST be bytes.')

        self.__streamData = data

