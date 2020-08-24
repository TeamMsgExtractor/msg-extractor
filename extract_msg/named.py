import copy
import logging
import pprint
import zlib

from extract_msg import constants
from extract_msg.utils import bytesToGuid, divide, properHex, roundUp



logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

class Named(object):
    __dir = '__nameid_version1.0'
    def __init__(self, msg):
        super(Named, self).__init__()
        self.__msg = msg
        guid_stream = self._getStream('__substg1.0_00020102')
        entry_stream = self._getStream('__substg1.0_00030102')
        names_stream = self._getStream('__substg1.0_00040102')
        guid_stream = self._getStream('__substg1.0_00020102', False) if guid_stream is None else guid_stream
        entry_stream = self._getStream('__substg1.0_00030102', False) if entry_stream is None else entry_stream
        names_stream = self._getStream('__substg1.0_00040102', False) if names_stream is None else names_stream
        self.guid_stream = guid_stream
        self.entry_stream = entry_stream
        self.names_stream = names_stream
        guid_stream_length = len(guid_stream)
        entry_stream_length = len(entry_stream)
        names_stream_length = len(names_stream)
        # TODO guid stream parsing
        guids = tuple([None, constants.GUID_PS_MAPI, constants.GUID_PS_PUBLIC_STRINGS] + [bytesToGuid(x) for x in divide(guid_stream, 16)])
        # TODO entry_stream parsing
        entries = []
        for x in divide(entry_stream, 8):
            tmp = constants.STNP_ENT.unpack(x)
            entry = {
                'id': tmp[0],
                'pid': tmp[2],
                'guid_index': tmp[1] >> 1,
                'pkind': tmp[1] & 1, # 0 if numerical, 1 if string
                }
            entry['guid'] = guids[entry['guid_index']]
            entries.append(entry)

        # Parse the names stream.
        names = {}
        pos = 0
        while pos < names_stream_length:
            name_length = constants.STNP_NAM.unpack(names_stream[pos:pos+4])[0]
            pos += 4 # Move to the start of the
            names[pos - 4] = names_stream[pos:pos+name_length].decode('utf_16_le') # Names are stored in the dictionary as the position they start at
            pos += roundUp(name_length, 4)

        self.entries = entries
        self.__names = names
        self.__guids = guids
        self.__properties = []
        for entry in entries:
            streamID = properHex(0x8000 + entry['pid'])
            msg._registerNamedProperty(entry, entry['pkind'], names[entry['id']] if entry['pkind'] == constants.STRING_NAMED else None)
            if msg.ExistsTypedProperty(streamID):
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

    def Exists(self, filename):
        """
        Checks if stream exists inside the named properties folder.
        """
        return self.__msg.Exists([self.__dir, filename])

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



class NamedAttachmentProperties(object):
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
        if self.__attachment.ExistsTypedProperty(streamID)[0]:
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



class StringNamedProperty(object):
    def __init__(self, entry, name, data):
        super(StringNamedProperty, self).__init__()
        self.__entry = entry
        self.__name = name
        self.__guidIndex = entry['guid_index']
        self.__guid = entry['guid']
        self.__namedPropertyID = entry['pid']
        # WARNING From the way the documentation is worded, this SHOULD work, but it doesn't.
        self.__streamID = 0x1000 + (zlib.crc32(name.lower().encode('utf-16-le')) ^ (self.__guidIndex << 1 | 1)) % 0x1F
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
        return constants.STRING_NAMED



class NumericalNamedProperty(object):
    def __init__(self, entry, data):
        super(NumericalNamedProperty, self).__init__()
        self.__propertyID = properHex(entry['id'], 4).upper()
        self.__guidIndex = entry['guid_index']
        self.__namedPropertyID = entry['pid']
        self.__streamID = 0x1000 + (entry['id'] ^ (self.__guidIndex << 1)) % 0x1F
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
