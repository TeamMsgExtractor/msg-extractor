from __future__ import annotations


__all__ = [
    'Named',
    'NamedProperties',
    'NamedPropertyBase',
    'NumericalNamedProperty',
    'StringNamedProperty',
]


import copy
import logging
import pprint

from typing import Any, Dict, List, Optional, Tuple, TYPE_CHECKING, Union

from .. import constants
from ..enums import NamedPropertyType
from ..utils import bytesToGuid, divide, makeWeakRef, properHex
from compressed_rtf.crc32 import crc32


# Allow for nice type checking.
if TYPE_CHECKING:
    from ..msg_classes.msg import MSGFile
    from ..attachments.attachment_base import AttachmentBase

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class Named:
    """
    Class for handling access to the named properties themselves.
    """

    __dir = '__nameid_version1.0'

    def __init__(self, msg : MSGFile):
        self.__msg = makeWeakRef(msg)
        # Get the basic streams. If all are emtpy, then nothing to do.
        guidStream = self._getStream('__substg1.0_00020102') or self._getStream('__substg1.0_00020102', False)
        entryStream = self._getStream('__substg1.0_00030102') or self._getStream('__substg1.0_00030102', False)
        self.guidStream = guidStream
        self.entryStream = entryStream
        self.namesStream = self._getStream('__substg1.0_00040102') or self._getStream('__substg1.0_00040102', False)
        # The if else stuff is for protection against None.
        guidStreamLength = len(guidStream) if guidStream else 0
        entryStreamLength = len(entryStream) if entryStream else 0

        self.__propertiesDict = {}
        self.__properties = []
        self.__guids = tuple()
        self.__names = {}

        # Check that we even have any entries. If there are none, nothing to do.
        if entryStream:
            guids = tuple([None, constants.ps.PS_MAPI, constants.ps.PS_PUBLIC_STRINGS] + [bytesToGuid(x) for x in divide(guidStream, 16)])
            entries : List[Dict[str, Any]]= []
            for rawStream in divide(entryStream, 8):
                tmp = constants.st.STNP_ENT.unpack(rawStream)
                entry = {
                    'id': tmp[0],
                    'pid': tmp[2],
                    'guid_index': tmp[1] >> 1,
                    'pkind': NamedPropertyType(tmp[1] & 1), # 0 if numerical, 1 if string.
                    'rawStream': rawStream,
                }
                entry['guid'] = guids[entry['guid_index']]
                entries.append(entry)

            self.entries = entries
            self.__guids = guids

            for entry in entries:
                self.__properties.append(StringNamedProperty(entry, self.__getName(entry['id'])) if entry['pkind'] == NamedPropertyType.STRING_NAMED else NumericalNamedProperty(entry))

            for property in self.__properties:
                name = property.name if isinstance(property, StringNamedProperty) else property.propertyID
                self.__propertiesDict[(name, property.guid)] = property

    def __contains__(self, key) -> bool:
        return key in self.__propertiesDict

    def __getitem__(self, propertyName : Tuple[str, str]):
        # Validate the key.
        if not hasattr(propertyName, '__len__') or len(propertyName) != 2:
            raise TypeError('Named property key must be a tuple of two strings.')

        # Case insensitive search of the dictionary.
        propertyName = (propertyName[0].upper(), propertyName[1].upper())
        for key in self.__propertiesDict.keys():
             if propertyName == (key[0].upper(), key[1].upper()):
                    return self.__propertiesDict[key]

        raise KeyError(propertyName)

    def __iter__(self):
        return self.__propertiesDict.__iter__()

    def __len__(self) -> int:
        return self.__propertiesDict.__len__()

    def __getName(self, offset : int) -> str:
        """
        Parses the offset into the named stream and returns the name found.
        """
        # We used to parse names by handing it as an array, as specified by the
        # documentation, but this new method allows for a little bit more wiggle
        # room in terms of what is accepted by the module.
        if offset & 3 != 0:
            # If the offset is not a multiple of 4, that is an error, but we are
            # reducing it to a warning.
            logger.warning(f'Malformed named properties detected due to bad offset ({offset}). Ignoring.')
        # Check that offset is in string stream.
        if offset > len(self.namesStream):
            raise ValueError('Failed to parse named property: offset was not in string stream.')

        # Get the length, in bytes, of the string.
        length = constants.st.STNP_NAM.unpack(self.namesStream[offset:offset + 4])[0]
        offset += 4

        # Make sure the string can be read entirely. If it can't, something was
        # corrupt.
        if offset + length > len(self.namesStream):
            raise ValueError(f'Failed to parse named property: length ({length}) of string overflows the string stream. This is probably due to a bad offset.')

        return self.namesStream[offset:offset + length].decode('utf-16-le')

    def _getStream(self, filename, prefix = True) -> Optional[bytes]:
        """
        Gets a binary representation of the requested filename.

        This should ALWAYS return a bytes object if it was found, otherwise
        returns None.

        :raises ReferenceError: The associated MSGFile instance has been garbage
            collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The msg file for this Named instance has been garbage collected.')
        return msg._getStream([self.__dir, filename], prefix = prefix)

    def _getStringStream(self, filename, prefix = True) -> Optional[str]:
        """
        Gets a string representation of the requested filename.

        Rather than the full filename, you should only feed this function the
        filename sans the type. So if the full name is "__substg1.0_001A001F",
        the filename this function should receive should be "__substg1.0_001A".

        This should ALWAYS return a string if it was found, otherwise returns
        None.

        :raises ReferenceError: The associated MSGFile instance has been garbage
            collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The msg file for this Named instance has been garbage collected.')
        return msg._getStringStream([self.__dir, filename], prefix = prefix)

    def exists(self, filename) -> bool:
        """
        Checks if stream exists inside the named properties folder.

        :raises ReferenceError: The associated MSGFile instance has been garbage
            collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The msg file for this Named instance has been garbage collected.')
        return msg.exists([self.__dir, filename])

    def sExists(self, filename) -> bool:
        """
        Checks if the string stream exists inside the named properties folder.

        :raises ReferenceError: The associated MSGFile instance has been garbage
            collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The msg file for this Named instance has been garbage collected.')
        return msg.sExists([self.__dir, filename])

    def get(self, propertyName, default = None):
        """
        Tries to get a named property based on its key. Returns :param default:
        if not found. Key is a tuple of the name and the property set GUID.
        """
        try:
            return self[propertyName]
        except KeyError:
            return default

    def keys(self):
        return self.__propertiesDict.keys()

    def pprintKeys(self):
        """
        Uses the pprint function on a sorted list of keys.
        """
        pprint.pprint(sorted(self.__propertiesDict.keys()))

    def values(self):
        return self.__propertiesDict.values()

    @property
    def dir(self):
        """
        Returns the directory inside the msg file where the named properties are located.
        """
        return self.__dir

    @property
    def msg(self) -> MSGFile:
        """
        Returns the Message instance the attachment belongs to.

        :raises ReferenceError: The associated MSGFile instance has been garbage
            collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The msg file for this Named instance has been garbage collected.')
        return msg

    @property
    def namedProperties(self) -> Dict:
        """
        Returns a copy of the dictionary containing all the named properties.
        """
        return copy.deepcopy(self.__propertiesDict)



class NamedProperties:
    """
    An instance that uses a Named instance and an extract-msg class to read the
    data of named properties.
    """
    def __init__(self, named, streamSource : Union[MSGFile, AttachmentBase]):
        """
        :param named: The named instance to refer to for named properties
            entries.
        :param streamSource: The source to use for acquiring the data of a named
            property.
        """
        self.__named = named
        self.__streamSource = makeWeakRef(streamSource)

    def __getitem__(self, item):
        """
        Get a named property using the [] operator. Item must be a named
        property instance or a tuple with 2 items: the name and the GUID string.

        :raises ReferenceError: The associated instance for getting actual
            property data has been garbage collected.
        """
        if (source := self.__streamSource()) is None:
            raise ReferenceError('The stream source for the NamedProperties instance has been garbage collected.')
        if isinstance(item, NamedPropertyBase):
            return source._getTypedData(item.propertyStreamID)
        else:
            return source._getTypedData(self.__named[item].propertyStreamID)

    def get(self, item, default = None):
        """
        Get a named property, returning the value of :param default: if not
        found. Item must be a tuple with 2 items: the name and the GUID string.

        :raises ReferenceError: The associated instance for getting actual
            property data has been garbage collected.
        """
        try:
            return self[item]
        except KeyError:
            return default



class NamedPropertyBase:
    def __init__(self, entry):
        self.__entry = entry
        self.__guidIndex = entry['guid_index']
        self.__namedPropertyID = entry['pid']
        self.__guid = entry['guid']
        self.__propertyStreamID = f'{0x8000 + self.__namedPropertyID:04X}'

    @property
    def guid(self) -> str:
        """
        The guid of the property's property set.
        """
        return self.__guid

    @property
    def guidIndex(self) -> int:
        """
        The guid index of the property's property set.
        """
        return self.__guidIndex

    @property
    def namedPropertyID(self) -> int:
        """
        The named property id.
        """
        return self.__namedPropertyID

    @property
    def propertyStreamID(self) -> str:
        """
        An ID usable for grabbing the value stream.
        """
        return self.__propertyStreamID

    @property
    def rawEntry(self) -> dict:
        return copy.deepcopy(self.__entry)

    @property
    def rawEntryStream(self) -> bytes:
        """
        The raw data used for the entry.
        """
        return self.__entry['rawStream']

    @property
    def type(self) -> NamedPropertyType:
        """
        The type of named property.
        """
        raise NotImplementedError('NamedPropertyBase cannot be used directly. Subclass it before using it.')



class StringNamedProperty(NamedPropertyBase):
    def __init__(self, entry, name):
        super().__init__(entry)
        self.__name = name

        # Finally got this to be correct after asking about it on a Microsoft
        # forum. Apparently it uses the same CRC-32 as the Compressed RTF
        # standard does, so we can just use the function defined in the
        # compressed-rtf Python module.
        #
        # First thing to note is that the name should only ever be lowered if it
        # is part of the PS_INTERNET_HEADERS property set **AND** it is
        # generated by certain versions of Outlook. As such, a little bit of
        # additional code will need to run to determine exactly what the stream
        # ID should be if it is in that property set.
        if self.guid == constants.ps.PS_INTERNET_HEADERS:
            # To be sure if it needs to be lower the most effective method would
            # be to just get the Stream ID and then check if the entry is in
            # there. If it isn't, then check the regular case and see. If it is
            # not in either... well, we don't use it for anything so it will
            # just be a warning, and the Stream ID will be set to 0.
            #
            # TODO: Unfortunately, doing this will need to be put off until a
            # different version, preferably after Python 2 support is removed,
            # as this will require restructuring a lot of internal code. For now
            # we just assume that it is lowercase.
            self.__streamID = 0x1000 + (crc32(name.lower().encode('utf-16-le')) ^ (self.guidIndex << 1 | 1)) % 0x1F

        else:
            # No special logic here to determine what to do.
            self.__streamID = 0x1000 + (crc32(name.encode('utf-16-le')) ^ (self.guidIndex << 1 | 1)) % 0x1F

    @property
    def name(self) -> str:
        """
        The name of the property.
        """
        return self.__name

    @property
    def streamID(self) -> int:
        """
        Returns the streamID of the named property. This may not be accurate.
        """
        return self.__streamID

    @property
    def type(self) -> NamedPropertyType:
        """
        Returns the type of the named property. This will either be NUMERICAL_NAMED or STRING_NAMED.
        """
        return NamedPropertyType.STRING_NAMED



class NumericalNamedProperty(NamedPropertyBase):
    def __init__(self, entry):
        super().__init__(entry)
        self.__propertyID = properHex(entry['id'], 4).upper()
        self.__streamID = 0x1000 + (entry['id'] ^ (self.guidIndex << 1)) % 0x1F

    @property
    def propertyID(self) -> str:
        """
        The actualy property id of the named property.
        """
        return self.__propertyID

    @property
    def streamID(self) -> int:
        """
        Returns the streamID of the named property. This may not be accurate
        """
        return self.__streamID

    @property
    def type(self) -> NamedPropertyType:
        """
        Returns the type of the named property. This will either be NUMERICAL_NAMED or STRING_NAMED.
        """
        return NamedPropertyType.NUMERICAL_NAMED
