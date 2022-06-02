import copy
import logging
import pprint

from . import constants
from .utils import bytesToGuid, divide, properHex, roundUp
from compressed_rtf.crc32 import crc32

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

class Named:
    __dir = '__nameid_version1.0'
    def __init__(self, msg):
        self.__msg = msg
        guidStream = self._getStream('__substg1.0_00020102') or self._getStream('__substg1.0_00020102', False)
        entryStream = self._getStream('__substg1.0_00030102') or self._getStream('__substg1.0_00030102', False)
        namesStream = self._getStream('__substg1.0_00040102') or self._getStream('__substg1.0_00040102', False)
        self.guidStream = guidStream
        self.entryStream = entryStream
        self.namesStream = namesStream
        guidStreamLength = len(guidStream)
        entryStreamLength = len(entryStream)
        namesStreamLength = len(namesStream)
        guids = tuple([None, constants.GUID_PS_MAPI, constants.GUID_PS_PUBLIC_STRINGS] + [bytesToGuid(x) for x in divide(guidStream, 16)])
        entries = []
        for rawStream in divide(entryStream, 8):
            tmp = constants.STNP_ENT.unpack(rawStream)
            entry = {
                'id': tmp[0],
                'pid': tmp[2],
                'guid_index': tmp[1] >> 1,
                'pkind': tmp[1] & 1, # 0 if numerical, 1 if string
                'rawStream': rawStream,
                }
            entry['guid'] = guids[entry['guid_index']]
            entries.append(entry)

        # Parse the names stream.
        names = {}
        pos = 0
        while pos < namesStreamLength:
            nameLength = constants.STNP_NAM.unpack(namesStream[pos:pos+4])[0]
            pos += 4 # Move to the start of the entry.
            names[pos - 4] = namesStream[pos:pos+nameLength].decode('utf-16-le') # Names are stored in the dictionary as the position they start at.
            pos += roundUp(nameLength, 4)

        self.entries = entries
        self.__names = names
        self.__guids = guids
        self.__properties = []
        for entry in entries:
            streamID = properHex(0x8000 + entry['pid'])
            msg._registerNamedProperty(entry, entry['pkind'], names[entry['id']] if entry['pkind'] == constants.STRING_NAMED else None)
            if msg.existsTypedProperty(streamID):
                self.__properties.append(StringNamedProperty(entry, names[entry['id']], msg._getTypedData(streamID)) if entry['pkind'] == constants.STRING_NAMED else NumericalNamedProperty(entry, msg._getTypedData(streamID)))
        self.__propertiesDict = {}
        for property in self.__properties:
            self.__propertiesDict[property.name if isinstance(property, StringNamedProperty) else property.propertyID] = property

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

    def getNamed(self, propertyName):
        """
        Tries to get a named property based on its name. Returns None if not found.
        """
        try:
            return self.namedProperties[propertyName]
        except KeyError:
            propertyName = propertyName.upper()
            for key in self.namedProperties.keys():
                if propertyName == key.upper():
                    return self.namedProperties[key]
            return None

    def getNamedValue(self, propertyName):
        """
        Tries to get a the value of a named property based on its name. Returns None if it is not found.
        """
        prop = self.getNamed(propertyName)
        return prop.data if prop is not None else None

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

    def pprintKeys(self):
        """
        Uses the pprint function on a sorted list of keys.
        """
        pprint.pprint(sorted(tuple(self.__propertiesDict.keys())))

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



class NamedAttachmentProperties:
    """
    The named properties associated with a specific attachment.
    """
    def __init__(self, attachment):
        self.__attachment = attachment
        self.__properties = []
        self.__propertiesDict = {}

    def defineProperty(self, entry, _type, name = None):
        """
        Informs the class of a named property that needs to be loaded.
        """
        streamID = properHex(0x8000 + entry['pid']).upper()
        if self.__attachment.existsTypedProperty(streamID)[0]:
            data = self.__attachment._getTypedData(streamID)
            property = StringNamedProperty(entry, name, data) if _type == constants.STRING_NAMED else NumericalNamedProperty(entry, data)
            self.__properties.append(property)
            self.__propertiesDict[property.name if isinstance(property, StringNamedProperty) else property.propertyID] = property

    @property
    def attachment(self):
        """
        The attachment that this NamedAttachmentProperties instance is associated with.
        """
        return self.__attachment

    @property
    def namedProperties(self):
        """
        Returns a copy of the dictionary containing all the named properties.
        """
        return copy.deepcopy(self.__propertiesDict)



class StringNamedProperty:
    def __init__(self, entry, name, data):
        self.__entry = entry
        self.__name = name
        self.__guidIndex = entry['guid_index']
        self.__guid = entry['guid']
        self.__namedPropertyID = entry['pid']

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
        if self.__guid == constants.GUID_PS_INTERNET_HEADERS:
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
            self.__streamID = 0x1000 + (crc32(name.lower().encode('utf-16-le')) ^ (self.__guidIndex << 1 | 1)) % 0x1F

        else:
            # No special logic here to determine what to do.
            self.__streamID = 0x1000 + (crc32(name.encode('utf-16-le')) ^ (self.__guidIndex << 1 | 1)) % 0x1F
        self.__data = data

    @property
    def data(self):
        """
        The data of the property.
        """
        return copy.deepcopy(self.__data)

    @property
    def guid(self):
        """
        The guid of the property's property set.
        """
        return self.__guid

    @property
    def name(self):
        """
        The name of the property.
        """
        return self.__name

    @property
    def namedPropertyID(self):
        """
        The named property id.
        """
        return self.__namedPropertyID

    @property
    def rawEntryStream(self):
        """
        The raw data used for the entry.
        """
        return self.__entry['rawStream']

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
        return constants.STRING_NAMED



class NumericalNamedProperty:
    def __init__(self, entry, data):
        self.__propertyID = properHex(entry['id'], 4).upper()
        self.__guidIndex = entry['guid_index']
        self.__namedPropertyID = entry['pid']
        self.__guid = entry['guid']
        self.__streamID = 0x1000 + (entry['id'] ^ (self.__guidIndex << 1)) % 0x1F
        self.__data = data
        self.__entry = entry

    @property
    def data(self):
        """
        The data of the property.
        """
        return copy.deepcopy(self.__data)

    @property
    def guid(self):
        """
        The guid of the property's property set.
        """
        return self.__guid

    @property
    def namedPropertyID(self):
        """
        The named property id.
        """
        return self.__namedPropertyID

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
        return constants.NUMERICAL_NAMED
