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

from functools import cached_property
from typing import (
        Any, List, Optional, Tuple, Type, TYPE_CHECKING, TypeVar, Union
    )

from ..constants import MSG_PATH, OVERRIDE_CLASS, SAVE_TYPE
from ..enums import AttachmentType
from ..properties.named import NamedProperties
from ..properties.properties_store import PropertiesStore
from ..utils import (
        msgPathToString, tryGetMimetype, verifyPropertyId, verifyType
    )


# Allow for nice type checking.
if TYPE_CHECKING:
    from ..msg_classes.msg import MSGFile

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

_T = TypeVar('_T')


class AttachmentBase(abc.ABC):
    """
    The base class for all standard Attachments used by the module.
    """

    def __init__(self, msg: MSGFile, dir_: str, propStore: PropertiesStore):
        """
        :param msg: the Message instance that the attachment belongs to.
        :param dir_: the directory inside the MSG file where the attachment is
            located.
        :param propStore: The PropertiesStore instance for the attachment.
        """
        self.__msg = weakref.ref(msg)
        self.__dir = dir_
        self.__props = propStore
        self.__namedProperties = NamedProperties(msg.named, self)
        self.__treePath = msg.treePath + [weakref.ref(self)]

    def _getTypedAs(self, _id: str, overrideClass = None, preserveNone: bool = True):
        """
        Like the other get as functions, but designed for when something
        could be multiple types (where only one will be present).

        This way you have no need to set the type, it will be handled for you.

        :param overrideClass: Class/function to use to morph the data that was
            read. The data will be the first argument to the class's
            ``__init__`` method or the function itself, if that is what is
            provided. By default, this will be completely ignored if the value
            was not found.
        :param preserveNone: If ``True`` (default), causes the function to
            ignore :param overrideClass: when the value could not be found (is
            ``None``). If this is changed to ``False``, then the value will be
            used regardless.
        """
        value = self._getTypedData(_id)
        # Check if we should be overriding the data type for this instance.
        if overrideClass is not None:
            if value is not None or not preserveNone:
                value = overrideClass(value)

        return value

    def _getTypedData(self, id, _type = None):
        """
        Gets the data for the specified id as the type that it is supposed to
        be.

        :param id: MUST be a 4 digit hexadecimal string.

        If you know for sure what type the data is before hand, you can specify
        it as being one of the strings in the constant
        ``FIXED_LENGTH_PROPS_STRING`` or ``VARIABLE_LENGTH_PROPS_STRING``.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
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
        Gets the property with the specified id as the type that it is supposed
        to be.

        :param id: MUST be a 4 digit hexadecimal string.

        If you know for sure what type the property is before hand, you can
        specify it as being one of the strings in the constant
        ``FIXED_LENGTH_PROPS_STRING`` or ``VARIABLE_LENGTH_PROPS_STRING``.
        """
        verifyPropertyId(propertyID)
        if _type:
            verifyType(_type)
            propertyID += _type

        notFound = object()
        ret = self.getPropertyVal(propertyID, notFound)
        if ret is notFound:
            return False, None

        return True, ret

    def _getTypedStream(self, filename: MSG_PATH, _type = None):
        """
        Gets the contents of the specified stream as the type that it is
        supposed to be.

        Rather than the full filename, you should only feed this function the
        filename sans the type. So if the full name is "__substg1.0_001A001F",
        the filename this function should receive should be "__substg1.0_001A".

        If you know for sure what type the stream is before hand, you can
        specify it as being one of the strings in the constant
        ``FIXED_LENGTH_PROPS_STRING`` or ``VARIABLE_LENGTH_PROPS_STRING``.

        If you have not specified the type, the type this method returns in many
        cases cannot be predicted. As such, when using this method it is best
        for you to check the type that it returns. If the function returns
        ``None``, that means it could not find the stream specified.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
        return msg._getTypedStream([self.__dir, msgPathToString(filename)], True, _type)

    def _handleFnc(self, _zip, filename, customPath: pathlib.Path, kwargs) -> pathlib.Path:
        """
        "Handle Filename Conflict"

        Internal function for use in determining how to modify the saving path
        when a file with the same name already exists. This is mainly because
        any save function that uses files will need to do this functionality.

        :returns: A ``pathlib.Path`` object to where the file should be saved.
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
                # Try to split the filename into a name and extension.
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

    def exists(self, filename: MSG_PATH) -> bool:
        """
        Checks if stream exists inside the attachment folder.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
        return msg.exists([self.__dir, msgPathToString(filename)])

    def sExists(self, filename: MSG_PATH) -> bool:
        """
        Checks if the string stream exists inside the attachment folder.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
        return msg.sExists([self.__dir, msgPathToString(filename)])

    def existsTypedProperty(self, id, _type = None) -> Tuple[bool, int]:
        """
        Determines if the stream with the provided id exists.

        The return of this function is 2 values, the first being a ``bool`` for
        if anything was found, and the second being how many were found.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
        return msg.existsTypedProperty(id, self.__dir, _type, True, self.__props)

    @abc.abstractmethod
    def getFilename(self, **kwargs) -> str:
        """
        Returns the filename to use for the attachment.

        :param contentId: Use the contentId, if available.
        :param customFilename: A custom name to use for the file.

        If the filename starts with "UnknownFilename" then there is no guarantee
        that the files will have exactly the same filename.
        """

    def getMultipleBinary(self, filename: MSG_PATH) -> Optional[List[bytes]]:
        """
        Gets a multiple binary property as a list of ``bytes`` objects.

        Like :meth:`getStringStream`, the 4 character type suffix should be
        omitted. So if you want the stream "__substg1.0_00011102" then the
        filename would simply be "__substg1.0_0001".

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
        return msg.getMultipleBinary([self.__dir, msgPathToString(filename)])

    def getMultipleString(self, filename: MSG_PATH) -> Optional[List[str]]:
        """
        Gets a multiple string property as a list of ``str`` objects.

        Like :meth:`getStringStream`, the 4 character type suffix should be
        omitted. So if you want the stream "__substg1.0_00011102" then the
        filename would simply be "__substg1.0_0001".

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
        return msg.getMultipleString([self.__dir, msgPathToString(filename)])

    def getNamedAs(self, propertyName: str, guid: str, overrideClass: OVERRIDE_CLASS[_T]) -> Optional[_T]:
        """
        Returns the named property, setting the class if specified.

        :param overrideClass: Class/function to use to morph the data that was
            read. The data will be the first argument to the class's
            ``__init__`` method or the function itself, if that is what is
            provided. If the value is ``None``, this function is not called. If
            you want it to be called regardless, you should handle the data
            directly.
        """
        value = self.getNamedProp(propertyName, guid)
        if value is not None:
            value = overrideClass(value)
        return value

    def getNamedProp(self, propertyName: str, guid: str, default: _T = None) -> Union[Any, _T]:
        """
        ``instance.namedProperties.get((propertyName, guid), default)``

        Can be overriden to create new behavior.
        """
        return self.namedProperties.get((propertyName, guid), default)

    def getPropertyAs(self, propertyName: Union[int, str], overrideClass: OVERRIDE_CLASS[_T]) -> Optional[_T]:
        """
        Returns the property, setting the class if found.

        :param overrideClass: Class/function to use to morph the data that was
            read. The data will be the first argument to the class's
            ``__init__`` method or the function itself, if that is what is
            provided. If the value is ``None``, this function is not called. If
            you want it to be called regardless, you should handle the data
            directly.
        """
        value = self.getPropertyVal(propertyName)

        if value is not None:
            value = overrideClass(value)

        return value

    def getPropertyVal(self, name: Union[int, str], default: _T = None) -> Union[Any, _T]:
        """
        instance.props.getValue(name, default)

        Can be overridden to create new behavior.
        """
        return self.props.getValue(name, default)

    def getSingleOrMultipleBinary(self, filename: MSG_PATH) -> Optional[Union[List[bytes], bytes]]:
        """
        A combination of :meth:`getStringStream` and :meth:`getMultipleString`.

        Checks to see if a single binary stream exists to return, otherwise
        tries to return the multiple binary stream of the same ID.

        Like :meth:`getStringStream`, the 4 character type suffix should be
        omitted. So if you want the stream "__substg1.0_00010102" then the
        filename would simply be "__substg1.0_0001".

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
        return msg.getSingleOrMultipleBinary([self.__dir, msgPathToString(filename)])

    def getSingleOrMultipleString(self, filename: MSG_PATH) -> Optional[Union[List[str], str]]:
        """
        A combination of :meth:`getStringStream` and :meth:`getMultipleString`.

        Checks to see if a single string stream exists to return, otherwise
        tries to return the multiple string stream of the same ID.

        Like :meth:`getStringStream`, the 4 character type suffix should be
        omitted. So if you want the stream "__substg1.0_0001001F" then the
        filename would simply be "__substg1.0_0001".

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
        return msg.getSingleOrMultipleString([self.__dir, msgPathToString(filename)])

    def getStream(self, filename: MSG_PATH) -> Optional[bytes]:
        """
        Gets a binary representation of the requested stream.

        This should ALWAYS return a ``bytes`` object if it was found, otherwise
        returns ``None``.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
        return msg.getStream([self.__dir, msgPathToString(filename)])

    def getStreamAs(self, streamID: MSG_PATH, overrideClass: OVERRIDE_CLASS[_T]) -> Optional[_T]:
        """
        Returns the specified stream, modifying it to the specified class if it
        is found.

        :param overrideClass: Class/function to use to morph the data that was
            read. The data will be the first argument to the class's
            ``__init__`` method or the function itself, if that is what is
            provided. If the value is ``None``, this function is not called. If
            you want it to be called regardless, you should handle the data
            directly.
        """
        value = self.getStream(streamID)

        if value is not None:
            value = overrideClass(value)

        return value

    def getStringStream(self, filename: MSG_PATH) -> Optional[str]:
        """
        Gets a string representation of the requested stream.

        Rather than the full filename, you should only feed this function the
        filename sans the type. So if the full name is "__substg1.0_001A001F",
        the filename this function should receive should be "__substg1.0_001A".

        This should ALWAYS return a string if it was found, otherwise returns
        ``None``.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
        return msg.getStringStream([self.__dir, msgPathToString(filename)])

    def getStringStreamAs(self, streamID: MSG_PATH, overrideClass: OVERRIDE_CLASS[_T]) -> Optional[_T]:
        """
        Returns the specified string stream, modifying it to the specified
        class if it is found.

        :param overrideClass: Class/function to use to morph the data that was
            read. The data will be the first argument to the class's
            ``__init__`` method or the function itself, if that is what is
            provided. If the value is ``None``, this function is not called. If
            you want it to be called regardless, you should handle the data
            directly.
        """
        value = self.getStream(streamID)

        if value is not None:
            value = overrideClass(value)

        return value

    def listDir(self, streams: bool = True, storages: bool = False) -> List[List[str]]:
        """
        Lists the streams and/or storages that exist in the attachment
        directory.

        Returns the paths *excluding* the attachment directory, allowing the
        paths to be directly used for accessing a stream.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
        return [path[1:] for path in msg.listDir(streams, storages, False)
                if len(path) > 1 and path[0] == self.__dir]

    def slistDir(self, streams: bool = True, storages: bool = False) -> List[str]:
        """
        Like listDir, except it returns the paths as strings.
        """
        return ['/'.join(path) for path in self.listDir(streams, storages)]

    @abc.abstractmethod
    def save(self, **kwargs) -> SAVE_TYPE:
        """
        Saves the attachment data.

        The name of the file is determined by the logic of :meth:`getFilename`.
        If you are a developer, ensure that you use this behavior.

        To change the directory that the attachment is saved to, set the value
        of :param customPath: when calling this function. The default save
        directory is the working directory.

        If you want to save the contents into a ``ZipFile`` or similar object,
        either pass a path to where you want to create one or pass an instance
        to :param zip:. If :param zip: is an instance, :param customPath: will
        refer to a location inside the zip file.

        :param extractEmbedded: If ``True``, causes the attachment, should it
            be an embedded MSG file, to save as a .msg file instead of calling its save function.
        :param skipEmbedded: If ``True``, skips saving this attachment if it is
            an embedded MSG file.

        :returns: A tuple that specifies how the data was saved. The value of
            the first item specifies what the second value will be.
        """

    @cached_property
    def attachmentEncoding(self) -> Optional[bytes]:
        """
        The encoding information about the attachment object.

        Will return ``b'\\x2A\\x86\\x48\\x86\\xf7\\x14\\x03\\x0b\\x01'`` if
        encoded in MacBinary format, otherwise it is unset.
        """
        return self.getStream('__substg1.0_37020102')

    @cached_property
    def additionalInformation(self) -> Optional[str]:
        """
        The additional information about the attachment.

        This property MUST be an empty string if attachmentEncoding is not set.
        Otherwise it MUST be set to a string of the format ":CREA:TYPE" where
        ":CREA" is the four-letter Macintosh file creator code and ":TYPE" is a
        four-letter Macintosh type code.
        """
        return self.getStringStream('__substg1.0_370F')

    @cached_property
    def cid(self) -> Optional[str]:
        """
        Returns the Content ID of the attachment, if it exists.
        """
        return self.getStringStream('__substg1.0_3712')

    @cached_property
    def clsid(self) -> str:
        """
        Returns the CLSID for the data stream/storage of the attachment.
        """
        # Set some default values.
        clsid = '00000000-0000-0000-0000-000000000000'
        dataStream = None

        if self.exists('__substg1.0_3701000D'):
            dataStream = [self.__dir, '__substg1.0_3701000D']
        elif self.exists('__substg1.0_37010102'):
            dataStream = [self.__dir, '__substg1.0_37010102']

        # If we found the right item, get the CLSID.
        if dataStream:
            if (msg := self.__msg()) is None:
                raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
            clsid = msg._getOleEntry(dataStream).clsid or clsid

        return clsid

    @property
    def contentId(self) -> Optional[str]:
        """
        Alias of :attr:`cid`.
        """
        return self.cid

    @property
    def createdAt(self) -> Optional[datetime.datetime]:
        """
        Alias of :attr:`creationTime`.
        """
        return self.creationTime

    @cached_property
    def creationTime(self) -> Optional[datetime.datetime]:
        """
        The time the attachment was created.
        """
        return self.getPropertyVal('30070040')

    @property
    @abc.abstractmethod
    def data(self) -> Optional[object]:
        """
        The attachment data, if any.

        Returns ``None`` if there is no data to save.
        """

    @cached_property
    def dataType(self) -> Optional[Type[object]]:
        """
        The class that the data type will use, if it can be retrieved.

        This is a safe way to do type checking on data before knowing if it will
        raise an exception. Returns ``None`` if no data will be returned or if
        an exception will be raised.
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

    @cached_property
    def displayName(self) -> Optional[str]:
        """
        Returns the display name of the folder.
        """
        return self.getStringStream('__substg1.0_3001')

    @cached_property
    def exceptionReplaceTime(self) -> Optional[datetime.datetime]:
        """
        The original date and time at which the instance in the recurrence
        pattern would have occurred if it were not an exception.

        Only applicable if the attachment is an Exception object.
        """
        return self.getPropertyVal('7FF90040')

    @cached_property
    def extension(self) -> Optional[str]:
        """
        The reported extension for the file.
        """
        return self.getStringStream('__substg1.0_3703')

    @cached_property
    def hidden(self) -> bool:
        """
        Indicates whether an Attachment object is hidden from the end user.
        """
        return bool(self.getPropertyVal('7FFE000B'))

    @cached_property
    def isAttachmentContactPhoto(self) -> bool:
        """
        Whether the attachment is a contact photo for a Contact object.
        """
        return bool(self.getPropertyVal('7FFF000B'))

    @cached_property
    def lastModificationTime(self) -> Optional[datetime.datetime]:
        """
        The last time the attachment was modified.
        """
        return self.getPropertyVal('30080040')

    @cached_property
    def longFilename(self) -> Optional[str]:
        """
        Returns the long file name of the attachment, if it exists.
        """
        return self.getStringStream('__substg1.0_3707')

    @cached_property
    def longPathname(self) -> Optional[str]:
        """
        The fully qualified path and file name with extension.
        """
        return self.getStringStream('__substg1.0_370D')

    @cached_property
    def mimetype(self) -> Optional[str]:
        """
        The content-type mime header of the attachment, if specified.
        """
        return tryGetMimetype(self, self.getStringStream('__substg1.0_370E'))

    @property
    def modifiedAt(self) -> Optional[datetime.datetime]:
        """
        Alias of :attr:`lastModificationTime`.
        """
        return self.lastModificationTime

    @property
    def msg(self) -> MSGFile:
        """
        Returns the ``MSGFile`` instance the attachment belongs to.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
        return msg

    @cached_property
    def name(self) -> Optional[str]:
        """
        The best name available for the file.

        Uses long filename before short.
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

    @cached_property
    def payloadClass(self) -> Optional[str]:
        """
        The class name of an object that can display the contents of the
        message.
        """
        return self.getStringStream('__substg1.0_371A')

    @property
    def props(self) -> PropertiesStore:
        """
        Returns the Properties instance of the attachment.
        """
        return self.__props

    @cached_property
    def renderingPosition(self) -> Optional[int]:
        """
        The offset, in rendered characters, to use when rendering the attachment
        within the main message text.

        A value of ``0xFFFFFFFF`` indicates a hidden attachment that is not to
        be rendered.
        """
        return self.getPropertyVal('370B0003')

    @cached_property
    def shortFilename(self) -> Optional[str]:
        """
        The short file name of the attachment, if it exists.
        """
        return self.getStringStream('__substg1.0_3704')

    @property
    def treePath(self) -> List[weakref.ReferenceType[Any]]:
        """
        A path, as a tuple of instances, needed to get to this instance through
        the MSGFile-Attachment tree.
        """
        return self.__treePath

    @property
    @abc.abstractmethod
    def type(self) -> AttachmentType:
        """
        An enum value that identifies the type of attachment.
        """
