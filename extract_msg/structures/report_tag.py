__all__ = [
    'ReportTag',
]


from typing import Optional

from ._helpers import BytesReader
from .entry_id import EntryID, FolderEntryID, MessageEntryID, StoreObjectEntryID


class ReportTag:
    """
    A Report Tag structure, as defined in [MS-OXOMSG].
    """

    def __init__(self, data: bytes):
        self.__rawData = data
        reader = BytesReader(data)

        self.__cookie = reader.assertRead(b'PCDFEB09\x00')

        self.__version = reader.readUnsignedInt()
        entrySize = reader.readInt()
        if entrySize:
            self.__storeEntryID = StoreObjectEntryID(reader.read(entrySize))
        else:
            self.__storeEntryID = None

        entrySize = reader.readInt()
        if entrySize:
            self.__folderEntryID = FolderEntryID(reader.read(entrySize))
        else:
            self.__folderEntryID = None

        entrySize = reader.readInt()
        if entrySize:
            self.__messageEntryID = MessageEntryID(reader.read(entrySize))
        else:
            self.__messageEntryID = None

        entrySize = reader.readInt()
        if entrySize:
            self.__searchFolderEntryID = FolderEntryID(reader.read(entrySize))
        else:
            self.__searchFolderEntryID = None

        entrySize = reader.readInt()
        if entrySize:
            self.__messageSearchKey = reader.read(entrySize)
        else:
            self.__messageSearchKey = None

        entrySize = reader.readInt()
        if entrySize:
            self.__ansiText = reader.read(entrySize)
        else:
            self.__ansiText = None

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        return self.__rawData

    @property
    def ansiText(self) -> Optional[bytes]:
        """
        The subject of the original message.

        Set to None if not present.
        """
        return self.__ansiText

    @property
    def cookie(self) -> bytes:
        """
        String used for validation.

        Set to ``b'PCDFEB09\x00'``.
        """
        return self.__cookie

    @property
    def folderEntryID(self) -> Optional[EntryID]:
        """
        The EntryID of the folder than contains the original message.
        """
        return self.__folderEntryID

    @property
    def messageEntryID(self) -> Optional[EntryID]:
        """
        The EntryID of the original message.
        """
        return self.__messageEntryID

    @property
    def messageSearchKey(self) -> Optional[bytes]:
        """
        The search key of the original message.
        """
        return self.__messageSearchKey

    @property
    def searchFolderEntryID(self) -> Optional[EntryID]:
        """
        The EntryID of an alternate folder that contains the original message.
        """
        return self.__searchFolderEntryID

    @property
    def storeEntryID(self) -> Optional[EntryID]:
        """
        The EntryID of the mailbox that contains the original message.
        """
        return self.__storeEntryID

    @property
    def version(self) -> int:
        """
        The version used.

        If SearchFolderEntryID is present, this MUST be ``0x00020001``,
        otherwise it MUST be ``0x00010001``.
        """
        return self.__version
