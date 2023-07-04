from __future__ import annotations


__all__ = [
    'AttachmentBase',
]


import abc
import datetime
import functools
import logging
import os
import pathlib
import weakref

from functools import cached_property, partial
from typing import List, Optional, Tuple, Type, TYPE_CHECKING

from .. import constants
from ..enums import AttachmentType
from ..properties.named import NamedProperties
from ..properties.prop import FixedLengthProp
from ..properties.properties_store import PropertiesStore
from ..utils import makeWeakRef, tryGetMimetype, verifyPropertyId, verifyType


# Allow for nice type checking.
if TYPE_CHECKING:
    from ..msg_classes.msg import MSGFile

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class AttachmentBase(abc.ABC):
    """
    The base class for all Attachments used by the module, if not overriden.
    """

    def __init__(self, msg : MSGFile, dir_, propStore : PropertiesStore):
        """
        :param msg: the Message instance that the attachment belongs to.
        :param dir_: the directory inside the msg file where the attachment is located.
        :param propStore: The PropertiesStore instance for the attachment. If
            not provided, it will be found automatically.
        """
        self.__msg = makeWeakRef(msg)
        self.__dir = dir_
        self.__props = propStore
        self.__namedProperties = NamedProperties(msg.named, self)
        self.__treePath = msg.treePath + [makeWeakRef(self)]

    def _getNamedAs(self, propertyName : str, guid : str, overrideClass = None, preserveNone : bool = True):
        """
        Returns the named property, setting the class if specified.

        :param overrideClass: Class/function to use to morph the data that was
            read. The data will be the first argument to the class's __init__
            function or the function itself, if that is what is provided. By
            default, this will be completely ignored if the value was not found.
        :param preserveNone: If true (default), causes the function to ignore
            :param overrideClass: when the value could not be found (is None).
            If this is changed to False, then the value will be used regardless.
        """
        value = self.namedProperties.get((propertyName, guid))
        # Check if we should be overriding the data type for this instance.
        if overrideClass is not None:
            if value is not None or not preserveNone:
                value = overrideClass(value)

        return value

    def _getPropertyAs(self, propertyName, overrideClass = None, preserveNone : bool = True):
        """
        Returns the property, setting the class if specified.

        :param overrideClass: Class/function to use to morph the data that was
            read. The data will be the first argument to the class's __init__
            function or the function itself, if that is what is provided. By
            default, this will be completely ignored if the value was not found.
        :param preserveNone: If True (default), causes the function to ignore
            :param overrideClass: when the value could not be found (is None).
            If this is changed to False, then the value will be used regardless.
        """
        try:
            value = self.props[propertyName].value
        except (KeyError, AttributeError):
            value = None
        # Check if we should be overriding the data type for this instance.
        if overrideClass is not None:
            if (value is not None or not preserveNone):
                value = overrideClass(value)

        return value

    def _getStream(self, filename) -> Optional[bytes]:
        """
        Gets a binary representation of the requested filename.

        This should ALWAYS return a bytes object if it was found, otherwise
        returns None.

        :raises ReferenceError: The associated MSGFile instance has been garbage
            collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The msg file for this Attachment instance has been garbage collected.')
        return msg._getStream([self.__dir, filename])

    def _getStreamAs(self, streamID, stringStream : bool = True, overrideClass = None, preserveNone : bool = True):
        """
        Returns the specified stream, modifying it to the class if specified.

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
        if stringStream:
            value = self._getStringStream(streamID)
        else:
            value = self._getStream(streamID)

        # Check if we should be overriding the data type for this instance.
        if overrideClass is not None:
            if value is not None or not preserveNone:
                value = overrideClass(value)

        return value

    def _getStringStream(self, filename) -> Optional[str]:
        """
        Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then :param prefer: specifies which will be
        returned.

        :raises ReferenceError: The associated MSGFile instance has been garbage
            collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The msg file for this Attachment instance has been garbage collected.')
        return msg._getStringStream([self.__dir, filename])

    def _getTypedAs(self, _id : str, overrideClass = None, preserveNone : bool = True):
        """
        Like the other get as functions, but designed for when something
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
        value = self._getTypedData(_id)
        # Check if we should be overriding the data type for this instance.
        if overrideClass is not None:
            if value is not None or not preserveNone:
                value = overrideClass(value)

        return value

    def _getTypedData(self, id, _type = None):
        """
        Gets the data for the specified id as the type that it is
        supposed to be. :param id: MUST be a 4 digit hexadecimal
        string.

        If you know for sure what type the data is before hand,
        you can specify it as being one of the strings in the
        constant FIXED_LENGTH_PROPS_STRING or
        VARIABLE_LENGTH_PROPS_STRING.

        :raises ReferenceError: The associated MSGFile instance has been garbage
            collected.
        """
        verifyPropertyId(id)
        id = id.upper()
        found, result = self._getTypedStream('__substg1.0_' + id, _type)
        if found:
            return result
        else:
            found, result = self._getTypedProperty(id, _type)
            return result if found else None

    def _getTypedProperty(self, propertyID, _type = None) -> Tuple[bool, Optional[object]]:
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

        :raises ReferenceError: The associated MSGFile instance has been garbage
            collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The msg file for this Attachment instance has been garbage collected.')
        return msg._getTypedStream([self.__dir, filename], True, _type)

    def _handleFnc(self, _zip, filename, customPath, kwargs) -> pathlib.Path:
        """
        "Handle Filename Conflict"

        Internal function for use in determining how to modify the saving path
        when a file with the same name already exists. This is mainly because
        any save function that uses files will need to do this functionality.

        :returns: A pathlib.Path object to where the file should be saved.
        """
        fullFilename = customPath / filename

        overwriteExisting = kwargs.get('overwriteExisting', False)

        if _zip:
            # If we are writing to a zip file and are not overwriting.
            if not overwriteExisting:
                name, ext = os.path.splitext(filename)
                nameList = _zip.namelist()
                if str(fullFilename).replace('\\', '/') in nameList:
                    for i in range(2, 100):
                        testName = customPath / f'{name} ({i}){ext}'
                        if str(testName).replace('\\', '/') not in nameList:
                            return testName
                    else:
                        # If we couldn't find one that didn't exist.
                        raise FileExistsError(f'Could not create the specified file because it already exists ("{fullFilename}").')
        else:
            if not overwriteExisting and fullFilename.exists():
                # Try to split the filename into a name and extention.
                name, ext = os.path.splitext(filename)
                # Try to add a number to it so that we can save without overwriting.
                for i in range(2, 100):
                    testName = customPath / f'{name} ({i}){ext}'
                    if not testName.exists():
                        return testName
                else:
                    # If we couldn't find one that didn't exist.
                    raise FileExistsError(f'Could not create the specified file because it already exists ("{fullFilename}").')

        return fullFilename

    def exists(self, filename) -> bool:
        """
        Checks if stream exists inside the attachment folder.

        :raises ReferenceError: The associated MSGFile instance has been garbage
            collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The msg file for this Attachment instance has been garbage collected.')
        return msg.exists([self.__dir, filename])

    def sExists(self, filename) -> bool:
        """
        Checks if the string stream exists inside the attachment folder.

        :raises ReferenceError: The associated MSGFile instance has been garbage
            collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The msg file for this Attachment instance has been garbage collected.')
        return msg.sExists([self.__dir, filename])

    def existsTypedProperty(self, id, _type = None) -> bool:
        """
        Determines if the stream with the provided id exists. The return of this
        function is 2 values, the first being a boolean for if anything was
        found, and the second being how many were found.

        :raises ReferenceError: The associated MSGFile instance has been garbage
            collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The msg file for this Attachment instance has been garbage collected.')
        return msg.existsTypedProperty(id, self.__dir, _type, True, self.__props)

    @abc.abstractmethod
    def getFilename(self, **kwargs) -> str:
        """
        Returns the filename to use for the attachment.

        :param contentId:      Use the contentId, if available.
        :param customFilename: A custom name to use for the file.

        If the filename starts with "UnknownFilename" then there is no guarentee
        that the files will have exactly the same filename.
        """

    @abc.abstractmethod
    def save(self, **kwargs) -> constants.SAVE_TYPE:
        """
        Saves the attachment data.

        The name of the file is determined by the logic of the getFilename
        function. If you are a developer, ensure that you use this behavior.

        To change the directory that the attachment is saved to, set the value
        of :param customPath: when calling this function. The default save
        directory is the working directory.

        If you want to save the contents into a ZipFile or similar object,
        either pass a path to where you want to create one or pass an instance
        to :param zip:. If :param zip: is an instance, :param customPath: will
        refer to a location inside the zip file.

        :param extractEmbedded: If True, causes the attachment, should it be an
            embedded MSG file, to save as a .msg file instead of calling it's
            save function.
        :param skipEmbedded: If True, skips saving this attachment if it is an
            embedded MSG file.

        :returns: A tuple that specifies how the data was saved. The value of
            the first item specifies what the second value will be.
        """

    @functools.cached_property
    def attachmentEncoding(self) -> Optional[bytes]:
        """
        The encoding information about the attachment object. Will return
        b'*\\x86H\\x86\\xf7\\x14\\x03\\x0b\\x01' if encoded in MacBinary format,
        otherwise it is unset.
        """
        return self._getStream('__substg1.0_37020102')

    @functools.cached_property
    def additionalInformation(self) -> Optional[str]:
        """
        The additional information about the attachment. This property MUST be
        an empty string if attachmentEncoding is not set. Otherwise it MUST be
        set to a string of the format ":CREA:TYPE" where ":CREA" is the
        four-letter Macintosh file creator code and ":TYPE" is a four-letter
        Macintosh type code.
        """
        return self._getStringStream('__substg1.0_370F')

    @functools.cached_property
    def cid(self) -> Optional[str]:
        """
        Returns the Content ID of the attachment, if it exists.
        """
        return self._getStringStream('__substg1.0_3712')

    @cached_property
    def clsid(self) -> str:
        """
        Returns the CLSID for the data stream/storage of the attachment.
        """
        # Set some default values.
        clsid = '00000000-0000-0000-0000-000000000000'
        dataStream = None

        # See if we can find the data stream/storage.
        if self.type in (AttachmentType.CUSTOM, AttachmentType.MSG):
            dataStream = [self.__dir, '__substg1.0_3701000D']
        elif self.type is AttachmentType.DATA:
            dataStream = [self.__dir, '__substg1.0_37010102']
        elif self.type is AttachmentType.UNSUPPORTED:
            # Special check for custom attachments.
            if self.exists('__substg1.0_3701000D'):
                dataStream = [self.__dir, '__substg1.0_3701000D']
            elif self.exists('__substg1.0_37010102'):
                dataStream = [self.__dir, '__substg1.0_37010102']

        # If we found the right item, get the CLSID.
        if dataStream:
            if (msg := self.__msg()) is None:
                raise ReferenceError('The msg file for this Attachment instance has been garbage collected.')
            clsid = msg._getOleEntry(dataStream).clsid or clsid

        return clsid

    @property
    def contentId(self) -> Optional[str]:
        return self.cid

    @property
    @abc.abstractmethod
    def data(self) -> Optional[object]:
        """
        The attachment data, if any. Returns None if there is no data to save.
        """

    @functools.cached_property
    def dataType(self) -> Optional[Type[type]]:
        """
        The class that the data type will use, if it can be retrieved.

        This is a safe way to do type checking on data before knowing if it will
        raise an exception. Returns None if no data will be returns or if an
        exception will be raised.
        """
        try:
            return None if self.data is None else self.data.__class__
        except Exception:
            # All exceptions that accessing data would cause should be silenced.
            return None

    @property
    def dir(self) -> str:
        """
        Returns the directory inside the MSG file where the attachment is
        located.
        """
        return self.__dir

    @functools.cached_property
    def displayName(self) -> Optional[str]:
        """
        Returns the display name of the folder.
        """
        return self._getStringStream('__substg1.0_3001')

    @functools.cached_property
    def exceptionReplaceTime(self) -> Optional[datetime.datetime]:
        """
        The original date and time at which the instance in the recurrence
        pattern would have occurred if it were not an exception.

        Only applicable if the attachment is an Exception object.
        """
        return self._getPropertyAs('7FF90040')

    @functools.cached_property
    def extension(self) -> Optional[str]:
        """
        The reported extension for the file.
        """
        return self._getStringStream('__substg1.0_3703')

    @functools.cached_property
    def hidden(self) -> bool:
        """
        Indicates whether an Attachment object is hidden from the end user.
        """
        return self._getPropertyAs('7FFE000B', bool, False)

    @functools.cached_property
    def isAttachmentContactPhoto(self) -> bool:
        """
        Whether the attachment is a contact photo for a Contact object.
        """
        return self._getPropertyAs('7FFF000B', bool, False)

    @functools.cached_property
    def longFilename(self) -> Optional[str]:
        """
        Returns the long file name of the attachment, if it exists.
        """
        return self._getStringStream('__substg1.0_3707')

    @functools.cached_property
    def longPathname(self) -> Optional[str]:
        """
        The fully qualified path and file name with extension.
        """
        return self._getStringStream('__substg1.0_370D')

    @functools.cached_property
    def mimetype(self) -> Optional[str]:
        """
        The content-type mime header of the attachment, if specified.
        """
        return self._getStreamAs('__substg1.0_370E', partial(tryGetMimetype, self), False)

    @property
    def msg(self) -> MSGFile:
        """
        Returns the Message instance the attachment belongs to.

        :raises ReferenceError: The associated MSGFile instance has been garbage
            collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The msg file for this Attachment instance has been garbage collected.')
        return msg

    @functools.cached_property
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

    @functools.cached_property
    def payloadClass(self) -> Optional[str]:
        """
        The class name of an object that can display the contents of the
        message.
        """
        return self._getStringStream('__substg1.0_371A')

    @property
    def props(self) -> PropertiesStore:
        """
        Returns the Properties instance of the attachment.
        """
        return self.__props

    @functools.cached_property
    def renderingPosition(self) -> Optional[int]:
        """
        The offset, in redered characters, to use when rendering the attachment
        within the main message text. A value of 0xFFFFFFFF indicates a hidden
        attachment that is not to be rendered.
        """
        return self._getPropertyAs('370B0003')

    @property
    def shortFilename(self) -> Optional[str]:
        """
        Returns the short file name of the attachment, if it exists.
        """
        return self._getStringStream('__substg1.0_3704')

    @property
    def treePath(self) -> List[weakref.ReferenceType]:
        """
        A path, as a tuple of instances, needed to get to this instance through
        the MSGFile-Attachment tree.
        """
        return self.__treePath

    @property
    @abc.abstractmethod
    def type(self) -> AttachmentType:
        """
        Returns an enum value that identifies the type of attachment.
        """
