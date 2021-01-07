import logging

from extract_msg import constants
from extract_msg.named import NamedAttachmentProperties
from extract_msg.prop import FixedLengthProp
from extract_msg.properties import Properties
from extract_msg.utils import verifyPropertyId, verifyType

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class AttachmentBase(object):
    """
    Stores the attachment data of a Message instance.
    Should the attachment be an embeded message, the
    class used to create it will be the same as the
    Message class used to create the attachment.
    """

    def __init__(self, msg, dir_):
        """
        :param msg: the Message instance that the attachment belongs to.
        :param dir_: the directory inside the msg file where the attachment is located.
        """
        object.__init__(self)
        self.__msg = msg
        self.__dir = dir_
        self.__props = Properties(self._getStream('__properties_version1.0'), constants.TYPE_ATTACHMENT)
        self.__namedProperties = NamedAttachmentProperties(self)


    def _ensureSet(self, variable, streamID, stringStream = True):
        """
        Ensures that the variable exists, otherwise will set it using the specified stream.
        After that, return said variable.

        If the specified stream is not a string stream, make sure to set :param string stream: to False.
        """
        try:
            return getattr(self, variable)
        except AttributeError:
            if stringStream:
                value = self._getStringStream(streamID)
            else:
                value = self._getStream(streamID)
            setattr(self, variable, value)
            return value

    def _ensureSetNamed(self, variable, propertyName):
        """
        Ensures that the variable exists, otherwise will set it using the named property.
        After that, return said variable.
        """
        try:
            return getattr(self, variable)
        except AttributeError:
            value = self.named.getNamedValue(propertyName)
            setattr(self, variable, value)
            return value

    def _ensureSetProperty(self, variable, propertyName):
        """
        Ensures that the variable exists, otherwise will set it using the property.
        After that, return said variable.
        """
        try:
            return getattr(self, variable)
        except AttributeError:
            try:
                value = self.props[propertyName].value
            except (KeyError, AttributeError):
                value = None
            setattr(self, variable, value)
            return value

    def _getStream(self, filename):
        return self.__msg._getStream([self.__dir, filename])

    def _getStringStream(self, filename):
        """
        Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then :param prefer: specifies which will be
        returned.
        """
        return self.__msg._getStringStream([self.__dir, filename])

    def _getTypedData(self, id, _type = None):
        """
        Gets the data for the specified id as the type that it is
        supposed to be. :param id: MUST be a 4 digit hexadecimal
        string.

        If you know for sure what type the data is before hand,
        you can specify it as being one of the strings in the
        constant FIXED_LENGTH_PROPS_STRING or
        VARIABLE_LENGTH_PROPS_STRING.
        """
        verifyPropertyId(id)
        id = id.upper()
        found, result = self._getTypedStream('__substg1.0_' + id, _type)
        if found:
            return result
        else:
            found, result = self._getTypedProperty(id, _type)
            return result if found else None

    def _getTypedProperty(self, propertyID, _type = None):
        """
        Gets the property with the specified id as the type that it
        is supposed to be. :param id: MUST be a 4 digit hexadecimal
        string.

        If you know for sure what type the property is before hand,
        you can specify it as being one of the strings in the
        constant FIXED_LENGTH_PROPS_STRING or
        VARIABLE_LENGTH_PROPS_STRING.
        """
        verifyPropertyId(propertyID)
        verifyType(_type)
        propertyID = propertyID.upper()
        for x in (propertyID + _type,) if _type is not None else self.props:
            if x.startswith(propertyID):
                prop = self.props[x]
                return True, (prop.value if isinstance(prop, FixedLengthProp) else prop)
        return False, None

    def _getTypedStream(self, filename, _type = None):
        """
        Gets the contents of the specified stream as the type that
        it is supposed to be.

        Rather than the full filename, you should only feed this
        function the filename sans the type. So if the full name
        is "__substg1.0_001A001F", the filename this function
        should receive should be "__substg1.0_001A".

        If you know for sure what type the stream is before hand,
        you can specify it as being one of the strings in the
        constant FIXED_LENGTH_PROPS_STRING or
        VARIABLE_LENGTH_PROPS_STRING.

        If you have not specified the type, the type this function
        returns in many cases cannot be predicted. As such, when
        using this function it is best for you to check the type
        that it returns. If the function returns None, that means
        it could not find the stream specified.
        """
        return self.__msg._getTypedStream([self.__dir, filename], True, _type)

    def _registerNamedProperty(self, entry, _type, name = None):
        self.__namedProperties.defineProperty(entry, _type, name)

    def Exists(self, filename):
        """
        Checks if stream exists inside the attachment folder.
        """
        return self.__msg.Exists([self.__dir, filename])

    def sExists(self, filename):
        """
        Checks if the string stream exists inside the attachment folder.
        """
        return self.__msg.sExists([self.__dir, filename])

    def ExistsTypedProperty(self, id, _type = None):
        """
        Determines if the stream with the provided id exists. The return of this
        function is 2 values, the first being a boolean for if anything was found,
        and the second being how many were found.
        """
        return self.__msg.ExistsTypedProperty(id, self.__dir, _type, True, self.__props)

    @property
    def dir(self):
        """
        Returns the directory inside the msg file where the attachment is located.
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
        The NamedAttachmentProperties instance for this attachment.
        """
        return self.__namedProperties

    @property
    def props(self):
        """
        Returns the Properties instance of the attachment.
        """
        return self.__props
