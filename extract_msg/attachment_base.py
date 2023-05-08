from __future__ import annotations


__all__ = [
    'AttachmentBase',
]


import datetime
import logging

from functools import partial
from typing import Optional, Tuple, TYPE_CHECKING

from .enums import AttachmentType, ErrorBehavior, PropertiesType
from .exceptions import StandardViolationError
from .named import NamedProperties
from .prop import FixedLengthProp
from .properties import Properties
from .utils import tryGetMimetype, verifyPropertyId, verifyType


# Allow for nice type checking.
if TYPE_CHECKING:
    from .msg import MSGFile

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class AttachmentBase:
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
        self.__msg = msg
        self.__dir = dir_
        if not self.exists('__properties_version1.0'):
            if (msg.errorBehavior & ErrorBehavior.STANDARDS_VIOLATION):
                logger.error('Attachments MUST have a property stream.')
            else:
                raise StandardViolationError('Attachments MUST have a property stream.') from None
        self.__props = Properties(self._getStream('__properties_version1.0'), PropertiesType.ATTACHMENT)
        self.__namedProperties = NamedProperties(msg.named, self)
        self.__treePath = msg.treePath + (self,)

    def _ensureSet(self, variable, streamID, stringStream = True, **kwargs):
        """
        Ensures that the variable exists, otherwise will set it using the
        specified stream. After that, return said variable.

        If the specified stream is not a string stream, make sure to set
        :param stringStream: to False.

        :param overrideClass: Class/function to use to morph the data that was
            read. The data will be the first argument to the class's __init__
            function or the function itself, if that is what is provided. By
            default, this will be completely ignored if the value was not found.
        :param preserveNone: If true (default), causes the function to ignore
            :param overrideClass: when the value could not be found (is None).
            If this is changed to False, then the value will be used regardless.
        """
        try:
            return getattr(self, variable)
        except AttributeError:
            if stringStream:
                value = self._getStringStream(streamID)
            else:
                value = self._getStream(streamID)
            # Check if we should be overriding the data type for this instance.
            if kwargs:
                overrideClass = kwargs.get('overrideClass')
                if overrideClass is not None and (value is not None or not kwargs.get('preserveNone', True)):
                    value = overrideClass(value)
            setattr(self, variable, value)
            return value

    def _ensureSetNamed(self, variable, propertyName : str, guid : str, **kwargs):
        """
        Ensures that the variable exists, otherwise will set it using the named
        property. After that, return said variable.

        :param overrideClass: Class/function to use to morph the data that was
            read. The data will be the first argument to the class's __init__
            function or the function itself, if that is what is provided. By
            default, this will be completely ignored if the value was not found.
        :param preserveNone: If true (default), causes the function to ignore
            :param overrideClass: when the value could not be found (is None).
            If this is changed to False, then the value will be used regardless.
        """
        try:
            return getattr(self, variable)
        except AttributeError:
            value = self.namedProperties.get((propertyName, guid))
            # Check if we should be overriding the data type for this instance.
            if kwargs:
                overrideClass = kwargs.get('overrideClass')
                if overrideClass is not None and (value is not None or not kwargs.get('preserveNone', True)):
                    value = overrideClass(value)
            setattr(self, variable, value)
            return value

    def _ensureSetProperty(self, variable, propertyName, **kwargs):
        """
        Ensures that the variable exists, otherwise will set it using the
        property. After that, return said variable.

        :param overrideClass: Class/function to use to morph the data that was
            read. The data will be the first argument to the class's __init__
            function or the function itself, if that is what is provided. By
            default, this will be completely ignored if the value was not found.
        :param preserveNone: If true (default), causes the function to ignore
            :param overrideClass: when the value could not be found (is None).
            If this is changed to False, then the value will be used regardless.
        """
        try:
            return getattr(self, variable)
        except AttributeError:
            try:
                value = self.props[propertyName].value
            except (KeyError, AttributeError):
                value = None
            # Check if we should be overriding the data type for this instance.
            if kwargs:
                overrideClass = kwargs.get('overrideClass')
                if overrideClass is not None and (value is not None or not kwargs.get('preserveNone', True)):
                    value = overrideClass(value)
            setattr(self, variable, value)
            return value

    def _ensureSetTyped(self, variable, _id, **kwargs):
        """
        Like the other ensure set functions, but designed for when something
        could be multiple types (where only one will be present). This way you
        have no need to set the type, it will be handled for you.

        :param overrideClass: Class/function to use to morph the data that was
            read. The data will be the first argument to the class's __init__
            function or the function itself, if that is what is provided. By
            default, this will be completely ignored if the value was not found.
        :param preserveNone: If true (default), causes the function to ignore
            :param overrideClass: when the value could not be found (is None).
            If this is changed to False, then the value will be used regardless.
        """
        try:
            return getattr(self, variable)
        except AttributeError:
            value = self._getTypedData(_id)
            # Check if we should be overriding the data type for this instance.
            if kwargs:
                overrideClass = kwargs.get('overrideClass')
                if overrideClass is not None and (value is not None or not kwargs.get('preserveNone', True)):
                    value = overrideClass(value)
            setattr(self, variable, value)
            return value

    def _getStream(self, filename) -> Optional[bytes]:
        return self.__msg._getStream([self.__dir, filename])

    def _getStringStream(self, filename) -> Optional[str]:
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

    def exists(self, filename) -> bool:
        """
        Checks if stream exists inside the attachment folder.
        """
        return self.__msg.exists([self.__dir, filename])

    def sExists(self, filename) -> bool:
        """
        Checks if the string stream exists inside the attachment folder.
        """
        return self.__msg.sExists([self.__dir, filename])

    def existsTypedProperty(self, id, _type = None) -> bool:
        """
        Determines if the stream with the provided id exists. The return of this
        function is 2 values, the first being a boolean for if anything was
        found, and the second being how many were found.
        """
        return self.__msg.existsTypedProperty(id, self.__dir, _type, True, self.__props)

    @property
    def attachmentEncoding(self) -> Optional[bytes]:
        """
        The encoding information about the attachment object. Will return
        b'*\x86H\x86\xf7\x14\x03\x0b\x01' if encoded in MacBinary format,
        otherwise it is unset.
        """
        return self._ensureSet('_attachmentEncoding', '__substg1.0_37020102', False)

    @property
    def additionalInformation(self) -> Optional[str]:
        """
        The additional information about the attachment. This property MUST be
        an empty string if attachmentEncoding is not set. Otherwise it MUST be
        set to a string of the format ":CREA:TYPE" where ":CREA" is the
        four-letter Macintosh file creator code and ":TYPE" is a four-letter
        Macintosh type code.
        """
        return self._ensureSet('_additionalInformation', '__substg1.0_370F')

    @property
    def cid(self) -> Optional[str]:
        """
        Returns the Content ID of the attachment, if it exists.
        """
        return self._ensureSet('_cid', '__substg1.0_3712')

    contendId = cid

    @property
    def dir(self):
        """
        Returns the directory inside the msg file where the attachment is
        located.
        """
        return self.__dir

    @property
    def displayName(self) -> Optional[str]:
        """
        Returns the display name of the folder.
        """
        return self._ensureSet('_displayName', '__substg1.0_3001')

    @property
    def exceptionReplaceTime(self) -> Optional[datetime.datetime]:
        """
        The original date and time at which the instance in the recurrence
        pattern would have occurred if it were not an exception.

        Only applicable if the attachment is an Exception object.
        """
        return self._ensureSetProperty('_exceptionReplaceTime', '7FF90040')

    @property
    def extension(self) -> Optional[str]:
        """
        The reported extension for the file.
        """
        return self._ensureSet('_extension', '__substg1.0_3703')

    @property
    def hidden(self) -> bool:
        """
        Indicates whether an Attachment object is hidden from the end user.
        """
        return self._ensureSetProperty('_hidden', '7FFE000B', overrideClass = bool, preserveNone = False)

    @property
    def isAttachmentContactPhoto(self) -> bool:
        """
        Whether the attachment is a contact photo for a Contact object.
        """
        return self._ensureSetProperty('_isAttachmentContactPhoto', '7FFF000B', overrideClass = bool, preserveNone = False)

    @property
    def longFilename(self) -> Optional[str]:
        """
        Returns the long file name of the attachment, if it exists.
        """
        return self._ensureSet('_longFilename', '__substg1.0_3707')

    @property
    def mimetype(self) -> Optional[str]:
        """
        The content-type mime header of the attachment, if specified.
        """
        return self._ensureSet('_mimetype', '__substg1.0_370E', overrideClass = partial(tryGetMimetype, self), preserveNone = False)

    @property
    def msg(self) -> MSGFile:
        """
        Returns the Message instance the attachment belongs to.
        """
        return self.__msg

    @property
    def name(self) -> Optional[str]:
        """
        The best name available for the file. Uses long filename before short.
        """
        if self.type is AttachmentType.MSG:
            if self.displayName:
                return self.displayName + '.msg'
        return self.longFilename or self.shortFilename

    @property
    def namedProperties(self) -> NamedProperties:
        """
        The NamedAttachmentProperties instance for this attachment.
        """
        return self.__namedProperties

    @property
    def payloadClass(self) -> Optional[str]:
        """
        The class name of an object that can display the contents of the
        message.
        """
        return self._ensureSet('_payloadClass', '__substg1.0_371A')

    @property
    def props(self) -> Properties:
        """
        Returns the Properties instance of the attachment.
        """
        return self.__props

    @property
    def renderingPosition(self) -> Optional[int]:
        """
        The offset, in redered characters, to use when rendering the attachment
        within the main message text. A value of 0xFFFFFFFF indicates a hidden
        attachment that is not to be rendered.
        """
        return self._ensureSetProperty('_renderingPosition', '370B0003')

    @property
    def shortFilename(self) -> Optional[str]:
        """
        Returns the short file name of the attachment, if it exists.
        """
        return self._ensureSet('_shortFilename', '__substg1.0_3704')

    @property
    def treePath(self) -> Tuple:
        """
        A path, as a tuple of instances, needed to get to this instance through
        the MSGFile-Attachment tree.
        """
        return self.__treePath

    @property
    def type(self) -> AttachmentType:
        """
        Returns the (internally used) type of the data.
        """
        return AttachmentType.UNKNOWN
