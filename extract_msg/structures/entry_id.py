import logging
from typing import Union

from ._helpers import BytesReader
from .. import constants
from ..enums import AddressBookType, DisplayType, EntryIDType, MacintoshEncoding, MessageFormat, OORBodyFormat
from ..utils import bitwiseAdjustedAnd


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


# First we define the main EntryID structure that is the base for the others.
class EntryID:
    """
    Base class for all EntryID structures. Use :classmethod autoCreate: to
    automatically create the correct EntryID structure type from the specified
    data.
    """

    @classmethod
    def autoCreate(cls, data) -> 'EntryID':
        """
        Automatically determines the type of EntryID and returns an instance of
        the correct subclass. If the subclass cannot be determined, will return
        a plain EntryID instance.
        """
        if len(data) < 20:
            raise ValueError('Cannot create an EntryID with less than 20 bytes.')
        providerUID = data[4:20]
        try:
            providerUID = EntryIDType(providerUID)
        except ValueError:
            raise ValueError(f'Unrecognized UID "{"".join(f"{x:02X}" for x in providerUID)}". You should probably report this to the developers.') from None

        # Now check the Provider UID against the known ones.
        if providerUID == EntryIDType.ONE_OFF_RECIPIENT:
            return OneOffRecipient(data)
        elif providerUID == EntryIDType.CA_OR_PDL_RECIPIENT:
            # Verify that the type signature is correct.
            if data[24:28] not in (b'\x04\x00\x00\x00', b'\x05\x00\x00\x00'):
                raise ValueError(f'Found Entry ID matching ContactAddress or PersonalDistributionList but the type was invalid ({data[24:28]}).')
            if data[24] == 4:
                return ContactAddress(data)
            else:
                return PersonalDistributionList(data)

        # If all else fails and we do recognize it, just return a plain EntryID.
        return cls(data)

    def __init__(self, data : bytes):
        self.__flags = data[:4]
        self.__providerUID = data[4:20]
        self.__rawData = data

    @property
    def flags(self) -> bytes:
        """
        The flags for this Entry ID.
        """
        return self.__flags

    @property
    def entryIDType(self) -> EntryIDType:
        """
        Returns an instance of EntryIDType corresponding to the provider UID of
        this EntryID. If none is found, raises a ValueError.
        """
        return EntryIDType(self.__providerUID)

    @property
    def longTerm(self) -> bool:
        """
        Whether the EntryID is long term or not.
        """
        return self__flags == b'\x00\x00\x00\x00'

    @property
    def providerUID(self) -> bytes:
        """
        The 16 byte UID that identifies the type of Entry ID.
        """
        return self.__providerUID

    @property
    def rawData(self) -> bytes:
        """
        The raw bytes used in this Entry ID.
        """
        return self.__rawData



# Now for the specific types.
class AddressBook(EntryID):
    def __init__(self, data : bytes):
        super().__init__(data)
        reader = BytesReader(data[20:])
        # Version *MUST* be 1.
        self.__version = reader.readUnsignedInt()
        if self.__version != 1:
            raise ValueError(f'Version must be 1 on address book entry IDs (got {self.__version}).')

        self.__type = AddressBookType(reader.readUnsignedInt())
        self.__X500DN = reader.readByteString()



class ContactAddress(EntryID):
    def __init__(self, data : bytes):
        super().__init__(data)



class OneOffRecipient(EntryID):
    def __init__(self, data : bytes):
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



class PermanentEntryID(EntryID):
    def __init__(self, data : bytes):
        super().__init__(data)
        unpacked = constants.STPEID.unpack(data[:28])
        if unpacked[0] != 0:
            raise TypeError(f'Not a PermanentEntryID (expected 0, got {unpacked[0]}).')
        self.__displayTypeString = DisplayType(unpacked[2])
        self.__distinguishedName = data[28:-1].decode('ascii') # Cut off the null character at the end and decode the data as ascii

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
