from __future__ import annotations


__all__ = [
    'Named',
    'NamedProperties',
    'NamedPropertyBase',
    'NumericalNamedProperty',
    'StringNamedProperty',
]


import abc
import copy
import logging
import pprint
import weakref

from typing import (
        Any, Dict, Iterable, Iterator, List, Optional, Tuple, TYPE_CHECKING,
        TypeVar, Union
    )

from .. import constants
from ..enums import ErrorBehavior, NamedPropertyType
from ..exceptions import InvalidPropertyIdError, StandardViolationError
from ..utils import bytesToGuid, divide, msgPathToString, verifyPropertyId
from compressed_rtf.crc32 import crc32


# Allow for nice type checking.
if TYPE_CHECKING:
    from ..msg_classes.msg import MSGFile
    from ..attachments.attachment_base import AttachmentBase

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

_T = TypeVar('_T')


class Named:
    """
    Class for handling access to the named properties themselves.
    """

    __dir = '__nameid_version1.0'

    def __init__(self, msg: MSGFile):
        self.__msg = weakref.ref(msg)
        # Get the basic streams. If all are emtpy, then nothing to do.
        guidStream = self.getStream('__substg1.0_00020102')
        entryStream = self.getStream('__substg1.0_00030102')
        self.guidStream = guidStream
        self.entryStream = entryStream
        self.namesStream = self.getStream('__substg1.0_00040102')

        self.__propertiesDict: Dict[Tuple[str, str], NamedPropertyBase] = {}

        self.__streamIDDict: Dict[str, Tuple[str, str]] = {}

        if guidStream is None:
            if ErrorBehavior.STANDARDS_VIOLATION in msg.errorBehavior:
                logger.warning('Standards Violation: Guid stream missing from named properties.')
                guidStream = b''
            else:
                raise StandardViolationError('Guid stream missing from named properties.')

        if entryStream is None:
            if ErrorBehavior.STANDARDS_VIOLATION in msg.errorBehavior:
                logger.warning('Standards Violation: Entry stream missing from named properties.')
                entryStream = b''
            else:
                raise StandardViolationError('Entry stream missing from named properties.')

        if self.namesStream is None:
            if ErrorBehavior.STANDARDS_VIOLATION in msg.errorBehavior:
                logger.error('Standards Violation: Guid stream missing from named properties. Will not parse named properties.')
                # Return immediately since the entry stream will likely fail.
                return
            else:
                raise StandardViolationError('Guid stream missing from named properties.')

        # Check that we even have any entries. If there are none, nothing to do.
        if entryStream:
            guids = tuple([None, constants.ps.PS_MAPI, constants.ps.PS_PUBLIC_STRINGS] + [bytesToGuid(x) for x in divide(guidStream, 16)])
            entries: List[Dict[str, Any]]= []
            for rawStream in divide(entryStream, 8):
                tmp = constants.st.ST_NP_ENT.unpack(rawStream)
                entry = {
                    'id': tmp[0],
                    'pid': tmp[2],
                    'guid_index': tmp[1] >> 1,
                    'pkind': NamedPropertyType(tmp[1] & 1), # 0 if numerical, 1 if string.
                    'rawStream': rawStream,
                }
                entry['guid'] = guids[entry['guid_index']]
                entries.append(entry)

            properties: List[NamedPropertyBase] = []

            for entry in entries:
                if entry['pkind'] == NamedPropertyType.STRING_NAMED:
                    name = None
                    try:
                        name = self.__getName(entry['id'])
                    except ValueError as e:
                        if ErrorBehavior.NAMED_NAME_STREAM in msg.errorBehavior:
                            logger.error(f'Dropping named property because it failed to acquire name from name stream: {e}')
                        else:
                            raise

                    if name:
                        properties.append(StringNamedProperty(entry, name))
                else:
                    properties.append(NumericalNamedProperty(entry))

            for property in properties:
                id_ = property.identifier
                self.__propertiesDict[id_] = property
                self.__streamIDDict[property.propertyStreamID] = id_

    def __contains__(self, key) -> bool:
        return key in self.__propertiesDict

    def __getitem__(self, propertyName: Tuple[str, str]) -> NamedPropertyBase:
        # Validate the key.
        if not hasattr(propertyName, '__len__') or len(propertyName) != 2:
            raise TypeError('Named property key must be a tuple of two strings.')

        # Case insensitive search of the dictionary.
        propertyName = (propertyName[0].upper(), propertyName[1].upper())
        for key in self.__propertiesDict.keys():
             if propertyName == (key[0].upper(), key[1].upper()):
                    return self.__propertiesDict[key]

        raise KeyError(propertyName)

    def __iter__(self) -> Iterator[Tuple[str, str]]:
        return self.__propertiesDict.__iter__()

    def __len__(self) -> int:
        return self.__propertiesDict.__len__()

    def __getName(self, offset: int) -> str:
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
        length = constants.st.ST_LE_I32.unpack(self.namesStream[offset:offset + 4])[0]
        offset += 4

        # Make sure the string can be read entirely. If it can't, something was
        # corrupt.
        if offset + length > len(self.namesStream):
            raise ValueError(f'Failed to parse named property: length ({length}) of string overflows the string stream. This is probably due to a bad offset.')

        return self.namesStream[offset:offset + length].decode('utf-16-le')

    def exists(self, filename: constants.MSG_PATH) -> bool:
        """
        Checks if stream exists inside the named properties folder.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Named instance has been garbage collected.')
        return msg.exists([self.__dir, msgPathToString(filename)], False)

    def get(self, propertyName: Tuple[str, str], default: _T = None) -> Union[NamedPropertyBase, _T]:
        """
        Tries to get a named property based on its key.

        Returns :param default: if not found. Key is a tuple of the name and
        the property set GUID.
        """
        try:
            return self[propertyName]
        except KeyError:
            return default

    def getPropNameByStreamID(self, streamID: Union[int, str]) -> Optional[Tuple[str, str]]:
        """
        Gets the name of a property (as a key for the internal dict) that is stored in the specified stream.

        Useful for determining if a stream/property stream entry is a named
        property.

        :param streamID: A 4 hex character identifier that will be checked. May
            also be an integer that can convert to 4 hex characters.

        :returns: The name, if the stream is a named property, otherwise
            ``None``.

        :raises InvalidPropertyIdError: The Stream ID is invalid.
        :raises TypeError: The Stream ID is not a valid type.
        """
        if isinstance(streamID, int):
            if streamID < 0x8000:
                # Definitely doesn't refer to a name property, so just return
                # None.
                return None
            if streamID > 0xFFFF:
                raise InvalidPropertyIdError('Stream ID is out of range.')
            streamID = f'{streamID:04X}'
        elif isinstance(streamID, str):
            verifyPropertyId(streamID)
        else:
            raise TypeError(':param streamID: MUST be an int or str.')

        return self.__streamIDDict.get(streamID)

    def getStream(self, filename: constants.MSG_PATH) -> Optional[bytes]:
        """
        Gets a binary representation of the requested stream.

        This should ALWAYS return a ``bytes`` object if it was found, otherwise
        returns ``None``.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Named instance has been garbage collected.')
        return msg.getStream([self.__dir, msgPathToString(filename)], False)

    def items(self) -> Iterable[Tuple[Tuple[str, str], NamedPropertyBase]]:
        return self.__propertiesDict.items()

    def keys(self) -> Iterable[Tuple[str, str]]:
        return self.__propertiesDict.keys()

    def pprintKeys(self) -> None:
        """
        Uses the pprint function on a sorted list of keys.
        """
        pprint.pprint(sorted(self.__propertiesDict.keys()))

    def values(self) -> Iterable[NamedPropertyBase]:
        return self.__propertiesDict.values()

    @property
    def dir(self):
        """
        Returns the directory inside the MSG file where the named properties are located.
        """
        return self.__dir

    @property
    def msg(self) -> MSGFile:
        """
        Returns the Message instance the attachment belongs to.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Named instance has been garbage collected.')
        return msg

    @property
    def namedProperties(self) -> Dict[Tuple[str, str], NamedPropertyBase]:
        """
        Returns a copy of the dictionary containing all the named properties.
        """
        return copy.deepcopy(self.__propertiesDict)



class NamedProperties:
    """
    An instance that uses a Named instance and an extract-msg class to read the
    data of named properties.
    """
    def __init__(self, named: Named, streamSource: Union[MSGFile, AttachmentBase]):
        """
        :param named: The Named instance to refer to for named properties
            entries.
        :param streamSource: The source to use for acquiring the data of a named
            property.
        """
        self.__named = named
        self.__streamSource = weakref.ref(streamSource)

    def __getitem__(self, item: Union[Tuple[str, str], NamedPropertyBase]):
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

    def get(self, item: Union[Tuple[str, str], NamedPropertyBase], default: _T = None) -> Union[Any, _T]:
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



class NamedPropertyBase(abc.ABC):
    def __init__(self, entry: Dict[str, Any]):
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
    @abc.abstractmethod
    def identifier(self) -> Tuple[str, str]:
        """
        An identifier that can be used to full identify the property.
        """

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
    def rawEntry(self) -> Dict[str, Any]:
        return copy.deepcopy(self.__entry)

    @property
    def rawEntryStream(self) -> bytes:
        """
        The raw data used for the entry.
        """
        return self.__entry['rawStream']

    @property
    @abc.abstractmethod
    def type(self) -> NamedPropertyType:
        """
        Returns the type of the named property. This will be a member of the
        NamedPropertyType enum.
        """



class StringNamedProperty(NamedPropertyBase):
    def __init__(self, entry: Dict, name: str):
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
    def identifier(self) -> Tuple[str, str]:
        return (self.name, self.guid)

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
        Returns the type of the named property. This will be a member of the
        NamedPropertyType enum.
        """
        return NamedPropertyType.STRING_NAMED



class NumericalNamedProperty(NamedPropertyBase):
    def __init__(self, entry: Dict):
        super().__init__(entry)
        self.__propertyID = f'{entry["id"]:04X}'
        self.__streamID = 0x1000 + (entry['id'] ^ (self.guidIndex << 1)) % 0x1F

    @property
    def identifier(self) -> Tuple[str, str]:
        return (self.propertyID, self.guid)

    @property
    def propertyID(self) -> str:
        """
        The actualy property id of the named property.
        """
        return self.__propertyID

    @property
    def streamID(self) -> int:
        """
        Returns the streamID of the named property. This may not be accurate.
        """
        return self.__streamID

    @property
    def type(self) -> NamedPropertyType:
        """
        Returns the type of the named property. This will be a member of the
        NamedPropertyType enum.
        """
        return NamedPropertyType.NUMERICAL_NAMED
