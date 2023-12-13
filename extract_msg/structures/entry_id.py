from __future__ import annotations


__all__ = [
    'AddressBookEntryID',
    'ContactAddressEntryID',
    'EntryID',
    'FolderEntryID',
    'MessageEntryID',
    'NNTPNewsgroupFolderEntryID',
    'OneOffRecipient',
    'PermanentEntryID',
    'PersonalDistributionListEntryID',
    'StoreObjectEntryID',
    'WrappedEntryID',
]


import abc
import logging

from typing import Optional, Union

from ._helpers import BytesReader
from .. import constants
from ..enums import (
        AddressBookType, ContactAddressIndex, DisplayType, EntryIDType,
        MacintoshEncoding, MessageFormat, MessageType, OORBodyFormat,
        WrappedType
    )
from ..exceptions import FeatureNotImplemented
from ..utils import bitwiseAdjustedAnd, bytesToGuid


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


# First we define the main EntryID structure that is the base for the others.
class EntryID(abc.ABC):
    """
    Base class for all EntryID structures. Use :classmethod autoCreate: to
    automatically create the correct EntryID structure type from the specified
    data.
    """

    @classmethod
    def autoCreate(cls, data: Optional[bytes]) -> Optional[EntryID]:
        """
        Automatically determines the type of EntryID and returns an instance of
        the correct subclass. If the subclass cannot be determined, will return
        a plain EntryID instance.
        """
        if not data:
            return None

        if len(data) < 20:
            raise ValueError('Cannot create an EntryID with less than 20 bytes.')
        providerUID = data[4:20]
        try:
            providerUID = EntryIDType(providerUID)
        except ValueError:
            raise ValueError(f'Unrecognized UID "{"".join(f"{x:02X}" for x in providerUID)}". You should probably report this to the developers.') from None

        # Now check the Provider UID against the known ones.
        if providerUID == EntryIDType.ADDRESS_BOOK_RECIPIENT:
            return AddressBookEntryID(data)
        if providerUID == EntryIDType.CA_OR_PDL_RECIPIENT:
            # Verify that the type signature is correct.
            if data[24:28] not in (b'\x04\x00\x00\x00', b'\x05\x00\x00\x00'):
                raise ValueError(f'Found Entry ID matching ContactAddress or PersonalDistributionList but the type was invalid ({data[24:28]}).')
            if data[24] == 4:
                return ContactAddressEntryID(data)
            else:
                return PersonalDistributionListEntryID(data)
        if providerUID == EntryIDType.NNTP_NEWSGROUP_FOLDER:
            # This is, of course, another one with a shared unique identifier.
            # Technically it's the UID of the provider, but why do multiple
            # structures use the same provider?
            if data[20] == 0xC:
                return NNTPNewsgroupFolderEntryID(data)
            else:
                return StoreObjectEntryID(data)
        if providerUID == EntryIDType.ONE_OFF_RECIPIENT:
            return OneOffRecipient(data)
        if providerUID == EntryIDType.PUBLIC_MESSAGE_STORE:
            if len(data) == 46:
                return FolderEntryID(data)
            else:
                return MessageEntryID(data)
        if providerUID == EntryIDType.WRAPPED:
            return WrappedEntryID(data)

        raise FeatureNotImplemented(f'UID for EntryID found in database, but no class was specified for it: {providerUID}')

    def __init__(self, data: bytes):
        self.__flags = data[:4]
        self.__providerUID = data[4:20]
        self.__rawData = data

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        return self.__rawData

    @property
    def flags(self) -> bytes:
        """
        The flags for this Entry ID.
        """
        return self.__flags

    @property
    def entryIDType(self) -> Union[EntryIDType, bytes]:
        """
        Returns an instance of EntryIDType corresponding to the provider UID of
        this EntryID. If none is found, returns the bytes.
        """
        try:
            return EntryIDType(self.__providerUID)
        except ValueError:
            return self.__providerUID

    @property
    def longTerm(self) -> bool:
        """
        Whether the EntryID is long term or not.
        """
        return self.__flags == b'\x00\x00\x00\x00'

    @property
    @abc.abstractmethod
    def position(self) -> int:
        """
        Used to tell the amount of bytes read in this EntryID. Useful for
        EntryID data that has been chained together with no separator.
        """

    @property
    def providerUID(self) -> bytes:
        """
        The 16 byte UID that identifies the type of Entry ID.
        """
        return self.__providerUID



# Now for the specific types.
class AddressBookEntryID(EntryID):
    """
    An Address Book EntryID structure, as specified in [MS-OXCDATA].
    """

    def __init__(self, data: bytes):
        super().__init__(data)
        reader = BytesReader(data[20:])
        # Version *MUST* be 1.
        self.__version = reader.readUnsignedInt()
        if self.__version != 1:
            raise ValueError(f'Version must be 1 on address book entry IDs (got {self.__version}).')

        self.__type = AddressBookType(reader.readUnsignedInt())
        self.__X500DN = reader.readByteString()
        self.__position = reader.tell() + 20

    @property
    def position(self) -> int:
        return self.__position

    @property
    def type(self) -> AddressBookType:
        """
        The type of the object.
        """
        return self.__type

    @property
    def version(self) -> int:
        """
        The version. MUST be 1.
        """
        return self.__version

    @property
    def X5000DN(self) -> bytes:
        """
        The X500 DN of the Address Book object.
        """
        return self.__X500DN



class ContactAddressEntryID(EntryID):
    """
    A Contact Address EntryID structure, as defined in [MS-OXCDATA]. Specifies a
    set of data representing recipients whose information is stored in a Contact
    object.
    """

    def __init__(self, data: bytes):
        super().__init__(data)
        reader = BytesReader(data[20:])
        if (version := reader.readUnsignedInt()) != 3:
            raise ValueError(f'Version must be 3 (got {version}).')
        if (type_ := reader.readUnsignedInt()) != 4:
            raise ValueError(f'Type must be 4 (got {type_}).')
        self.__index = ContactAddressIndex(reader.readUnsignedInt())
        self.__entryIdCount = reader.readUnsignedInt()
        self.__entryID = MessageEntryID(reader.read(self.__entryIdCount))
        self.__position = reader.tell() + 20

    @property
    def entryID(self) -> MessageEntryID:
        """
        The EntryID contained in this object.
        """
        return self.__entryID

    @property
    def entryIDCount(self) -> int:
        """
        The size, in bytes, of the EntryID contained in this object.
        """
        return self.__entryIdCount

    @property
    def index(self) -> ContactAddressIndex:
        """
        The electronic address in the contact information to use.
        """
        return self.__index

    @property
    def position(self) -> int:
        return self.__position



class FolderEntryID(EntryID):
    """
    A Folder EntryID structure, as defined in [MS-OXCDATA].
    """

    __SIZE__: int = 46

    def __init__(self, data: bytes):
        super().__init__(data)
        reader = BytesReader(data[20:])
        self.__folderType = MessageType(reader.readUnsignedShort())
        self.__databaseGuid = bytesToGuid(reader.read(16))
        # This entry is 6 bytes, so we pull some shenanigans to unpack it.
        self.__globalCounter = constants.st.ST_LE_UI64.unpack(reader.read(6) + b'\x00\x00')[0]
        reader.assertNull(2, 'Pad bytes were not 0.')

    @property
    def databaseGuid(self) -> str:
        """
        A GUID associated with the Store Object and corresponding to the
        ReplicaID field of the FID structure.
        """
        return self.__databaseGuid

    @property
    def folderType(self) -> MessageType:
        """
        The type of folder.
        """
        return self.__folderType

    @property
    def globalCounter(self) -> int:
        """
        An unsigned integer identifying the folder.
        """
        return self.__globalCounter

    @property
    def position(self) -> int:
        return self.__SIZE__



class MessageEntryID(EntryID):
    """
    A Message EntryID structure, as defined in [MS-OXCDATA].
    """

    __SIZE__: int = 70

    def __init__(self, data: bytes):
        super().__init__(data)
        reader = BytesReader(data[20:])
        self.__messageType = MessageType(reader.readUnsignedShort())
        self.__folderDatabaseGuid = bytesToGuid(reader.read(16))
        # This entry is 6 bytes, so we pull some shenanigans to unpack it.
        self.__folderGlobalCounter = constants.st.ST_LE_UI64.unpack(reader.read(6) + b'\x00\x00')[0]
        reader.assertNull(2, 'Pad bytes were not 0.')
        self.__messageDatabaseGuid = bytesToGuid(reader.read(16))
        # This entry is 6 bytes, so we pull some shenanigans to unpack it.
        self.__messageGlobalCounter = constants.st.ST_LE_UI64.unpack(reader.read(6) + b'\x00\x00')[0]
        reader.assertNull(2, 'Pad bytes were not 0.')
        # Not sure why Microsoft decided to say "yes, let's do 2 6-byte integers
        # followed by 2 pad bytes each" instead of just 2 8-byte integers with a
        # maximum value, but here we are.

    @property
    def folderDatabaseGuid(self) -> str:
        """
        A GUID associated with the Store object of the folder in which the
        message resides and corresponding to the ReplicaId field in the folder
        ID structure.
        """
        return self.__folderDatabaseGuid

    @property
    def folderGlobalCounter(self) -> int:
        """
        An unsigned integer identifying the folder in which the message resides.
        """
        return self.__folderGlobalCounter

    @property
    def messageDatabaseGuid(self) -> str:
        """
        A GUID associated with the Store object of the message and corresponding
        to the ReplicaId field of the Message ID structure.
        """
        return self.__messageDatabaseGuid

    @property
    def messageGlobalCounter(self) -> int:
        """
        An unsigned integer identifying the message.
        """
        return self.__messageGlobalCounter

    @property
    def messageType(self) -> MessageType:
        """
        The Store object type.
        """
        return self.__messageType

    @property
    def position(self) -> int:
        return self.__SIZE__



class NNTPNewsgroupFolderEntryID(EntryID):
    """
    A NNTP Newsgroup Folder EntryID structure, as defined in [MS-OXCDATA].
    """

    def __init__(self, data: bytes):
        super().__init__(data)
        reader = BytesReader(data[20:])
        self.__folderType = reader.readUnsignedShort()
        if self.__folderType != 0x000C:
            raise ValueError(f'Folder type was not 0x000C (got {self.__folderType})')
        self.__newsgroupName = reader.readByteString()
        self.__position = reader.tell() + 20

    @property
    def folderType(self) -> int:
        """
        The type of folder. MUST be ``0x000C``.
        """
        return self.__folderType

    @property
    def newsgroupName(self) -> bytes:
        """
        The name of the newsgroup, as an ANSI string.
        """
        return self.__newsgroupName

    @property
    def position(self) -> int:
        return self.__position



class OneOffRecipient(EntryID):
    """
    A One-Off EntryID structure, as specified in [MS-OXCDATA].
    """

    def __init__(self, data: bytes):
        super().__init__(data)
        # Create a reader to easily
        reader = BytesReader(data[20:])
        self.__version = reader.readUnsignedShort()
        # It's not really flags, but I can't come up with a descriptive name for
        # this collection of data, so `flagsThing` it is.
        flagsThing = reader.readUnsignedShort()

        # Just a little forewarning, I am *well* aware that these masks for each
        # flag do not match the specification as you might expect. That is
        # because, unlike with other parts of the documentation, these bytes
        # are, for some reason, *not read together*. This means they are not
        # meant to be swapped for little endian, and as such I had to flip the
        # masks to compensate. Again, this is *despite other portions of the
        # documentation using an identical format and being read in little
        # endian. Took me way too long to figure out why this was not working
        # despite following the documentation to the letter. If I had to guess,
        # the reason this one is not flipped and the others are is because this
        # is not grouped together.

        self.__macintoshEncoding = MacintoshEncoding(bitwiseAdjustedAnd(flagsThing, 0xC))
        self.__format = OORBodyFormat(bitwiseAdjustedAnd(flagsThing, 0x1E))
        # Flag to indicate how messages are to be sent. 0 means TNEF, 1 means
        # MIME.
        self.__messageFormat = MessageFormat(bitwiseAdjustedAnd(flagsThing, 0x1))
        # Whether the strings are UTF-16 or not.
        self.__stringFormatUnicode = bitwiseAdjustedAnd(flagsThing, 0x8000) == 1
        self.__canLookup = bitwiseAdjustedAnd(flagsThing, 0x1000) == 0
        if self.__stringFormatUnicode:
            self.__displayName = reader.readUtf16String()
            self.__addressType = reader.readUtf16String()
            self.__emailAddress = reader.readUtf16String()
        else:
            # Don't actually know how to properly handle this kind of encoding,
            # since the documentation doesn't define exactly what encoding to
            # use for this of even how to find out, so for now we just don't
            # decode it at all and just leave it as bytes.
            self.__displayName = reader.readByteString()
            self.__addressType = reader.readByteString()
            self.__emailAddress = reader.readByteString()

        self.__position = reader.tell() + 20

    @property
    def addressType(self) -> Union[str, bytes]:
        """
        The address type for this Recipient.
        """
        return self.__addressType

    @property
    def areStringUnicode(self) -> bool:
        """
        Whether or not the strings are in UTF-16 format.
        """
        return self.__stringFormatUnicode

    @property
    def canLookup(self) -> bool:
        """
        Whether the server can lookup the user's email in the address book.
        """
        return self.__canLookup

    @property
    def displayName(self) -> Union[str, bytes]:
        """
        The display name for this Recipient.
        """
        return self.__displayName

    @property
    def emailAddress(self) -> Union[str, bytes]:
        """
        The email address for this Recipient.
        """
        return self.__emailAddress

    @property
    def format(self) -> OORBodyFormat:
        """
        The message body format desired for this recipient.
        """
        return self.__format

    @property
    def macintoshEncoding(self) -> MacintoshEncoding:
        """
        The encoding used for Macintosh-specific data attachments.
        """
        return self.__macintoshEncoding

    @property
    def messageFormat(self) -> MessageFormat:
        """
        The message format to use for messages sent to this recipient.
        """
        return self.__messageFormat

    @property
    def position(self) -> int:
        return self.__position



class PermanentEntryID(EntryID):
    """
    A Permanent EntryID structure, as defined in [MS-OXNSPI].
    """

    def __init__(self, data: bytes):
        super().__init__(data)
        reader = BytesReader(data)
        unpacked = reader.readStruct(constants.st.ST_PEID)
        if unpacked[0] != 0:
            raise TypeError(f'Not a PermanentEntryID (expected 0, got {unpacked[0]}).')
        self.__displayTypeString = DisplayType(unpacked[2])
        self.__distinguishedName = reader.readAsciiString()
        self.__position = reader.tell()

    @property
    def displayTypeString(self) -> DisplayType:
        """
        Returns the display type string value.
        """
        return self.__displayTypeString

    @property
    def distinguishedName(self) -> str:
        """
        Returns the distinguished name.
        """
        return self.__distinguishedName

    @property
    def position(self) -> int:
        return self.__position



class PersonalDistributionListEntryID(EntryID):
    """
    A Personal Distribution List EntryID structure, as defined in [MS-OXCDATA].
    """

    def __init__(self, data: bytes):
        super().__init__(data)
        reader = BytesReader(data[20:])
        if (arg := reader.readUnsignedInt()) != 3:
            raise ValueError(f'Version must be 3 (got {arg}).')
        if (arg := reader.readUnsignedInt()) != 5:
            raise ValueError(f'Type must be 5 (got {arg}).')
        if (arg := reader.readUnsignedInt()) != 0xFF:
            raise ValueError(f'Index must be 255 (got {arg}).')
        self.__entryIdCount = reader.readUnsignedInt()
        self.__entryID = MessageEntryID(reader.read(self.__entryIdCount))
        self.__position = reader.tell() + 20

    @property
    def entryID(self) -> MessageEntryID:
        """
        The EntryID contained in this object.
        """
        return self.__entryID

    @property
    def entryIDCount(self) -> int:
        """
        The size, in bytes, of the EntryID contained in this object.
        """
        return self.__entryIdCount

    @property
    def position(self) -> int:
        return self.__position



class StoreObjectEntryID(EntryID):
    """
    A Store Object EntryID structure, as defined in [MS-OXCDATA].
    """

    def __init__(self, data: bytes):
        super().__init__(data)
        reader = BytesReader(data[20:])

        self.__version = reader.readUnsignedByte()
        if self.__version != 0:
            raise ValueError(f'Version was not set to 0 (got {self.__version}).')

        self.__flag = reader.readUnsignedByte()
        if self.__flag != 0:
            raise ValueError(f'Flag was not set to 0 (got {self.__flag}).')

        self.__dllFileName = reader.read(14)

        self.__wrappedFlags = reader.readUnsignedInt()
        if self.__wrappedFlags != 0:
            raise ValueError(f'Wrapped flags was not set to 0 (got {self.__wrappedFlags}).')

        self.__wrappedProviderUID = reader.read(16)
        self.__wrappedType = WrappedType(reader.readUnsignedInt())
        # Don't know how this is encoded, just that it is "single-byte
        # characters".
        self.__serverShortname = reader.readByteString()
        if self.__wrappedProviderUID == b'\x1B\x55\xFA\x20\xAA\x66\x11\xCD\x9B\xC8\x00\xAA\x00\x2F\xC4\x5A':
            self.__mailboxDN = reader.readAsciiString()
        else:
            self.__mailboxDN = None

        self.__position = reader.tell() + 20

    @property
    def dllFileName(self) -> bytes:
        """
        Must be set to b'emsmdb.dll\\x00\\x00\\x00\\x00'.
        """
        return self.__dllFileName

    @property
    def flag(self) -> int:
        return self.__flag

    @property
    def mailboxDN(self) -> Optional[str]:
        """
        A string representing the X500 DN of the mailbox, as specified in
        [MS-OXOAB]. THis field is present only for mailbox databases.
        """
        return self.__mailboxDN

    @property
    def position(self) -> int:
        return self.__position

    @property
    def serverShortname(self) -> bytes:
        """
        A string of single-byte characters indicating the short name or NetBIOS
        name of the server.
        """
        return self.__serverShortname

    @property
    def version(self) -> int:
        return self.__version

    @property
    def wrappedProviderUID(self) -> bytes:
        return self.__wrappedProviderUID

    @property
    def wrappedType(self) -> WrappedType:
        """
        Determined by where the folder is located.
        """
        return self.__wrappedType



class WrappedEntryID(EntryID):
    """
    A WrappedEntryId structure, as specified in [MS-OXOCNTC].
    """

    def __init__(self, data: bytes):
        super().__init__(data)
        # Grab the type byte and parse it.
        self.__type = data[20]
        bits = self.__type & 0xF
        if bits == 0:
            self.__embeddedEntryID = OneOffRecipient(data[21:])
        elif bits == 3 or bits == 4:
            self.__embeddedEntryID = MessageEntryID(data[21:])
        elif bits == 5 or bits == 6:
            self.__embeddedEntryID = AddressBookEntryID(data[21:])
        else:
            raise ValueError(f'Found wrapped entry id with invalid type (type bits were {bits}).')

        self.__embeddedIsOneOff = self.__type & 0x80 == 0
        self.__position = 21 + self.__embeddedEntryID.position

    @property
    def embeddedEntryID(self) -> EntryID:
        """
        The embedded EntryID of this object.
        """
        return self.__embeddedEntryID

    @property
    def embeddedIsOneOff(self) -> bool:
        """
        Whether the embedded EntryID is a One-Off EntryID.
        """
        return self.__embeddedIsOneOff

    @property
    def position(self) -> int:
        return self.__position

    @property
    def type(self) -> int:
        """
        The type bits of this object.
        """
        return self.__type
