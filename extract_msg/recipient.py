import logging

from extract_msg import constants
from extract_msg.data import PermanentEntryID
from extract_msg.properties import Properties
from extract_msg.utils import verifyPropertyId, verifyType


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class Recipient(object):
    """
    Contains the data of one of the recipients in an msg file.
    """
    def __init__(self, _dir, msg):
        object.__init__(self)
        self.__msg = msg  # Allows calls to original msg file
        self.__dir = _dir
        self.__props = Properties(self._getStream('__properties_version1.0'), constants.TYPE_RECIPIENT)
        self.__email = self._getStringStream('__substg1.0_39FE')
        if not self.__email:
            self.__email = self._getStringStream('__substg1.0_3003')
        self.__name = self._getStringStream('__substg1.0_3001')
        self.__type = self.__props.get('0C150003').value
        self.__formatted = u'{0} <{1}>'.format(self.__name, self.__email)

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
        self.__msg._getTypedStream(self, [self.__dir, filename], True, _type)

    def Exists(self, filename):
        """
        Checks if stream exists inside the recipient folder.
        """
        return self.__msg.Exists([self.__dir, filename])

    def sExists(self, filename):
        """
        Checks if the string stream exists inside the recipient folder.
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
    def account(self):
        """
        Returns the account of this recipient.
        """
        return self._ensureSet('_account', '__substg1.0_3A00')

    @property
    def email(self):
        """
        Returns the recipient's email.
        """
        return self.__email

    @property
    def entryID(self):
        """
        Returns the recipient's name.
        """
        try:
            return self.__entryID
        except AttributeError:
            self.__entryID = PermanentEntryID(self._getStream('__substg1.0_0FFF0102'))
            return self.__entryID

    @property
    def formatted(self):
        """
        Returns the formatted recipient string.
        """
        return self.__formatted

    @property
    def instanceKey(self):
        """
        Returns the instance key of this recipient.
        """
        return self._ensureSet('_instanceKey', '__substg1.0_0FF60102', False)

    @property
    def name(self):
        """
        Returns the recipient's name.
        """
        return self.__name

    @property
    def props(self):
        """
        Returns the Properties instance of the recipient.
        """
        return self.__props

    @property
    def recordKey(self):
        """
        Returns the instance key of this recipient.
        """
        return self._ensureSet('_recordKey', '__substg1.0_0FF90102', False)

    @property
    def searchKey(self):
        """
        Returns the search key of this recipient.
        """
        return self._ensureSet('_searchKey', '__substg1.0_300B0102', False)

    @property
    def smtpAddress(self):
        """
        Returns the SMTP address of this recipient.
        """
        return self._ensureSet('_smtpAddress', '__substg1.0_39FE')

    @property
    def transmittableDisplayName(self):
        """
        Returns the transmittable display name of this recipient.
        """
        return self._ensureSet('_transmittableDisplayName', '__substg1.0_3A20')

    @property
    def type(self):
        """
        Returns the recipient type. Type is:
            * Sender if `type & 0xf == 0`
            * To if `type & 0xf == 1`
            * Cc if `type & 0xf == 2`
            * Bcc if `type & 0xf == 3`
        """
        return self.__type
