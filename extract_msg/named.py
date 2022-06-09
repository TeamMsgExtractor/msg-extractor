import copy
import logging
import pprint

from . import constants
from .enums import Guid, NamedPropertyType
from .utils import bytesToGuid, divide, properHex, roundUp
from compressed_rtf.crc32 import crc32


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class Named:
    __dir = '__nameid_version1.0'
    def __init__(self, msg):
        self.__msg = msg
        # Get the basic streams. If all are emtpy, then nothing to do.
        guidStream = self._getStream('__substg1.0_00020102') or self._getStream('__substg1.0_00020102', False)
        entryStream = self._getStream('__substg1.0_00030102') or self._getStream('__substg1.0_00030102', False)
        namesStream = self._getStream('__substg1.0_00040102') or self._getStream('__substg1.0_00040102', False)
        self.guidStream = guidStream
        self.entryStream = entryStream
        self.namesStream = namesStream
        # The if else stuff is for protection against None.
        guidStreamLength = len(guidStream) if guidStream else 0
        entryStreamLength = len(entryStream) if entryStream else 0
        namesStreamLength = len(namesStream) if namesStream else 0

        self.__propertiesDict = {}
        self.__properties = []
        self.__guids = tuple()
        self.__names = {}

        # Check that we even have any entries. If there are none, nothing to do.
        if entryStream:
            guids = tuple([None, Guid.PS_MAPI.value, Guid.PS_PUBLIC_STRINGS.value] + [bytesToGuid(x) for x in divide(guidStream, 16)])
            entries = []
            for rawStream in divide(entryStream, 8):
                tmp = constants.STNP_ENT.unpack(rawStream)
                entry = {
                    'id': tmp[0],
                    'pid': tmp[2],
                    'guid_index': tmp[1] >> 1,
                    'pkind': NamedPropertyType(tmp[1] & 1), # 0 if numerical, 1 if string.
                    'rawStream': rawStream,
                    }
                entry['guid'] = guids[entry['guid_index']]
                entries.append(entry)

            # Parse the names stream.
            names = self.__names
            pos = 0
            while pos < namesStreamLength:
                nameLength = constants.STNP_NAM.unpack(namesStream[pos:pos+4])[0]
                pos += 4 # Move to the start of the entry.
                names[pos - 4] = namesStream[pos:pos+nameLength].decode('utf-16-le') # Names are stored in the dictionary as the position they start at.
                pos += roundUp(nameLength, 4)

            self.entries = entries
            self.__guids = guids

            for entry in entries:
                streamID = properHex(0x8000 + entry['pid'])
                self.__properties.append(StringNamedProperty(entry, names[entry['id']]) if entry['pkind'] == NamedPropertyType.STRING_NAMED else NumericalNamedProperty(entry))

            for property in self.__properties:
                self.__propertiesDict[property.name if isinstance(property, StringNamedProperty) else property.propertyID] = property

    def __getitem__(self, key):
        return self.__propertiesDict[key]

    def __iter__(self):
        return self.__propertiesDict.__iter__()

    def __len__(self):
        return self.__propertiesDict.__len__()

    def _getStream(self, filename, prefix = True):
        return self.__msg._getStream([self.__dir, filename], prefix = prefix)

    def _getStringStream(self, filename, prefix = True):
        """
        Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then :param prefer: specifies which will be
        returned.
        """
        return self.__msg._getStringStream([self.__dir, filename], prefix = prefix)

    def exists(self, filename):
        """
        Checks if stream exists inside the named properties folder.
        """
        return self.__msg.exists([self.__dir, filename])

    def sExists(self, filename):
        """
        Checks if the string stream exists inside the named properties folder.
        """
        return self.__msg.sExists([self.__dir, filename])

    def get(self, propertyName, default = None):
        """
        Tries to get a named property based on its name. Returns :param default:
        if not found.
        """
        try:
            return self.__propertiesDict[propertyName]
        except KeyError:
            propertyName = propertyName.upper()
            for key in self.__propertiesDict.keys():
                if propertyName == key.upper():
                    return self.__propertiesDict[key]
            return default

    def keys(self):
        return self.__propertiesDict.keys()

    def pprintKeys(self):
        """
        Uses the pprint function on a sorted list of keys.
        """
        pprint.pprint(sorted(tuple(self.__propertiesDict.keys())))

    def values(self):
        return self.__propertiesDict.values()

    @property
    def dir(self):
        """
        Returns the directory inside the msg file where the named properties are located.
        """
        return self.__dir

    @property
    def msg(self):
        """
        Returns the Message instance the attachment belongs to.
        """
        return self.__msg

    @property
    def namedProperties(self):
        """
        Returns a copy of the dictionary containing all the named properties.
        """
        return copy.deepcopy(self.__propertiesDict)



class NamedProperties:
    """
    An instance that uses a Named instance and an extract-msg class to read the
    data of named properties.
    """
    def __init__(self, named, streamSource):
        """
        :param named: The named instance to refer to for named properties
            entries.
        :param streamSource: The source to use for acquiring the data of a named
            property.
        """
        self.__named = named
        self.__streamSource = streamSource

    def __getitem__(self, item):
        if isinstance(item, NamedPropertyBase):
            return self.__streamSource._getTypedData(item.propertyStreamID)
        else:
            return self.__streamSource._getTypedData(self.__named[item].propertyStreamID)

    def get(self, item, default = None):
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
    def guid(self):
        """
        The guid of the property's property set.
        """
        return self.__guid

    @property
    def guidIndex(self):
        """
        The guid index of the property's property set.
        """
        return self.__guidIndex

    @property
    def namedPropertyID(self):
        """
        The named property id.
        """
        return self.__namedPropertyID

    @property
    def propertyStreamID(self):
        """
        An ID usable for grabbing the value stream.
        """
        return self.__propertyStreamID

    @property
    def rawEntry(self) -> dict:
        return copy.deepcopy(entry)

    @property
    def rawEntryStream(self):
        """
        The raw data used for the entry.
        """
        return self.__entry['rawStream']



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
        if self.guid == Guid.PS_INTERNET_HEADERS.value:
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
    def name(self):
        """
        The name of the property.
        """
        return self.__name

    @property
    def streamID(self):
        """
        Returns the streamID of the named property. This may not be accurate.
        """
        return self.__streamID

    @property
    def type(self):
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
    def propertyID(self):
        """
        The actualy property id of the named property.
        """
        return self.__propertyID

    @property
    def streamID(self):
        """
        Returns the streamID of the named property. This may not be accurate
        """
        return self.__streamID

    @property
    def type(self):
        """
        Returns the type of the named property. This will either be NUMERICAL_NAMED or STRING_NAMED.
        """
        return NamedPropertyType.NUMERICAL_NAMED
