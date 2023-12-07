"""
Miscellaneous ID structures used in MSG files that don't fit into any of the
other ID structure classifications.
"""


__all__ = [
    'FolderID',
    'GlobalObjectID',
    'MessageID',
    'ServerID',
]


import datetime

from .. import constants
from ._helpers import BytesReader
from ..utils import filetimeToDatetime


class FolderID:
    """
    A Folder ID structure specified in [MS-OXCDATA].
    """

    __SIZE__: int = 16

    def __init__(self, data: bytes):
        self.__rawData = data
        self.__replicaID = constants.st.ST_LE_UI16.unpack(data[:2])[0]
        # This entry is 6 bytes, so we pull some shenanigans to unpack it.
        self.__globalCounter = constants.st.ST_LE_UI64.unpack(data[2:8] + b'\x00\x00')[0]

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        return self.__rawData

    @property
    def globalCounter(self) -> int:
        """
        An unsigned integer identifying the folder within its Store object.
        """
        return self.__globalCounter

    @property
    def replicaID(self) -> int:
        """
        An unsigned integer identifying a Store object.
        """
        return self.__replicaID



class GlobalObjectID:
    """
    A GlobalObjectID structure, as specified in [MS-OXOCAL].
    """

    def __init__(self, data: bytes):
        self.__rawData = data
        reader = BytesReader(data)
        expectedBytes = b'\x04\x00\x00\x00\x82\x00\xE0\x00\x74\xC5\xB7\x10\x1A\x82\xE0\x08'
        errorMsg = 'Byte Array ID did not match for GlobalObjectID (got {actual}).'
        self.__byteArrayID = reader.assertRead(expectedBytes, errorMsg)
        self.__yh = reader.read(1)
        self.__yl = reader.read(1)
        self.__year = constants.st.ST_BE_UI16.unpack(self.__yh + self.__yl)[0]
        self.__month = reader.readUnsignedByte()
        self.__day = reader.readUnsignedByte()
        self.__creationTime = filetimeToDatetime(reader.readUnsignedLong())
        reader.assertNull(8, 'Reserved was not set to null.')
        size = reader.readUnsignedInt()
        self.__data = reader.read(size)

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        return self.__rawData

    @property
    def byteArrayID(self) -> bytes:
        """
        An array of 16 bytes identifying the bytes this BLOB as a Global Object
        ID.
        """
        return self.__byteArrayID

    @property
    def creationTime(self) -> datetime.datetime:
        """
        The date and time when this Global Object ID was generated.
        """
        return self.__creationTime

    @property
    def data(self) -> bytes:
        """
        An array of bytes that ensures the uniqueness of the Global Object ID
        amoung all Calendar objects in all mailboxes.
        """
        return self.__data

    @property
    def day(self) -> int:
        """
        The day from the PidLidExceptionReplaceTime property if the object
        represents an exception. Otherwise, this is 0.
        """
        return self.__day

    @property
    def month(self) -> int:
        """
        The month from the PidLidExceptionReplaceTime property if the object
        represents an exception. Otherwise, this is 0.
        """
        return self.__month

    @property
    def year(self) -> int:
        """
        The year from the PidLidExceptionReplaceTime property if the object
        represents an exception. Otherwise, this is 0.
        """
        return self.__year



class MessageID:
    """
    A Message ID structure, as defined in [MS-OXCDATA].
    """

    __SIZE__: int = 8

    def __init__(self, data: bytes):
        self.__rawData = data
        self.__replicaID = constants.st.ST_LE_UI16.unpack(data[:2])[0]
        # This entry is 6 bytes, so we pull some shenanigans to unpack it.
        self.__globalCounter = constants.st.ST_LE_UI64.unpack(data[2:8] + b'\x00\x00')[0]

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        return self.__rawData

    @property
    def globalCounter(self) -> int:
        """
        An unsigned integer identifying the folder within its Store object.
        """
        return self.__globalCounter

    @property
    def isFolder(self) -> bool:
        """
        Tells if the object pointed to is actually a folder.
        """
        return self.__globalCounter == 0 and self.__replicaID == 0

    @property
    def replicaID(self) -> int:
        """
        An unsigned integer identifying a Store object.
        """
        return self.__replicaID



class ServerID:
    """
    Class representing a PtypServerId.
    """
    def __init__(self, data: bytes):
        """
        :param data: The data to use to create the ServerID.

        :raises TypeError: The data is not a ServerID.
        """
        # According to the docs, the first byte being a 1 means it follows this
        # structure. A value of 0 means it does not.
        if data[0] != 1:
            raise TypeError('Not a standard ServerID.')
        self.__rawData = data
        self.__folderID = FolderID(data[1:9])
        self.__messageID = MessageID(data[9:17])
        self.__instance = constants.st.ST_LE_UI32.unpack(data[17:21])[0]

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        return self.__rawData

    @property
    def folderID(self) -> FolderID:
        """
        The folder the message will be in.
        """
        return self.__folderID

    @property
    def instance(self) -> int:
        """
        Instance number that is only used in multivalue properties for searching
        for a specific ServerID. Will otherwise be 0.
        """
        return self.__instance

    @property
    def messageID(self) -> MessageID:
        """
        A MessageID identifying a message in a folder identified by the folderID
        instance of this object. If the object pointed to is a folder, both
        properties will be 0.
        """
        return self.__messageID
