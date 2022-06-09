"""
Various small data structures used in extract_msg.
"""

from . import constants


class FolderID:
    """
    A Folder ID structure specified in [MS-OXCDATA].
    """
    def __init__(self, data : bytes):
        self.__replicaID = constants.STUI16.unpack(data[:2])
        # This entry is 6 bytes, so we pull some shenanigans to unpack it.
        self.__globalCounter = constants.STUI64.unpack(data[2:8] + b'\x00\x00')

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



class MessageID:
    """

    """
    def __init__(self, data):
        self.__replicaID = constants.STUI16.unpack(data[:2])
        # This entry is 6 bytes, so we pull some shenanigans to unpack it.
        self.__globalCounter = constants.STUI64.unpack(data[2:8] + b'\x00\x00')

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



class PermanentEntryID:
    def __init__(self, data : bytes):
        self.__data = data
        unpacked = constants.STPEID.unpack(data[:28])
        if unpacked[0] != 0:
            raise TypeError(f'Not a PermanentEntryID (expected 0, got {unpacked[0]}).')
        self.__providerUID = unpacked[1]
        self.__displayTypeString = unpacked[2]
        self.__distinguishedName = data[28:-1].decode('ascii') # Cut off the null character at the end and decode the data as ascii

    @property
    def data(self) -> bytes:
        """
        Returns the raw data used to generate this instance.
        """
        return self.__data

    @property
    def displayTypeString(self) -> int:
        """
        Returns the display type string. This will be one of the display type constants.
        """
        return self.__displayTypeString

    @property
    def distinguishedName(self) -> str:
        """
        Returns the distinguished name.
        """
        return self.__distinguishedName

    @property
    def providerUID(self):
        """
        Returns the provider UID.
        """
        return self.__providerUID



class ServerID:
    """
    Class representing a PtypServerId.
    """
    def __init__(self, data : bytes):
        """
        :param data: The data to use to create the ServerID.

        :raises TypeError: if the data is not a ServerID.
        """
        # According to the docs, the first byte being a 1 means it follows this
        # structure. A value of 0 means it does not.
        if data[0] != 1:
            raise TypeError('Not a standard ServerID.')
        self.__raw = data
        self.__folderID = FolderID(data[1:9])
        self.__messageID = MessageID(data[9:17])
        self.__instance = constants.STUI32.unpack(data[17:21])

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

    @property
    def raw(self) -> bytes:
        """
        The raw data used to create this object.
        """
        return self.__raw
