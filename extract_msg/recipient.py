from __future__ import annotations


__all__ = [
    'Recipient',
]


import enum
import functools
import logging
import weakref

from typing import (
        Any, Generic, List, Optional, Tuple, TYPE_CHECKING, Type, TypeVar, Union
    )

from .constants import MSG_PATH, OVERRIDE_CLASS
from .enums import ErrorBehavior, PropertiesType
from .exceptions import StandardViolationError
from .properties.properties_store import PropertiesStore
from .structures.entry_id import PermanentEntryID
from .utils import msgPathToString


if TYPE_CHECKING:
    from .msg_classes.msg import MSGFile

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

_T = TypeVar('_T')
_RT = TypeVar('_RT', bound = enum.IntEnum)


class Recipient(Generic[_RT]):
    """
    Contains the data of one of the recipients in an MSG file.
    """

    def __init__(self, _dir: str, msg: MSGFile, recipientTypeClass: Type[_RT]):
        self.__msg = weakref.ref(msg) # Allows calls to original msg file.
        self.__dir = _dir
        if not self.exists('__properties_version1.0'):
            if ErrorBehavior.STANDARDS_VIOLATION in msg.errorBehavior:
                logger.error('Recipients MUST have a property stream.')
            else:
                raise StandardViolationError('Recipients MUST have a property stream.') from None
        self.__props = PropertiesStore(self.getStream('__properties_version1.0'), PropertiesType.RECIPIENT)
        self.__email = self.getStringStream('__substg1.0_39FE')
        if not self.__email:
            self.__email = self.getStringStream('__substg1.0_3003')
        self.__name = self.getStringStream('__substg1.0_3001')
        self.__typeFlags = self.__props.getValue('0C150003', 0)
        self.__type = recipientTypeClass(0xF & self.__typeFlags)
        self.__formatted = f'{self.__name} <{self.__email}>'

    def exists(self, filename: MSG_PATH) -> bool:
        """
        Checks if stream exists inside the recipient folder.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Recipient instance has been garbage collected.')
        return msg.exists([self.__dir, msgPathToString(filename)])

    def sExists(self, filename: MSG_PATH) -> bool:
        """
        Checks if the string stream exists inside the recipient folder.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Recipient instance has been garbage collected.')
        return msg.sExists([self.__dir, msgPathToString(filename)])

    def existsTypedProperty(self, id, _type = None) -> Tuple[bool, int]:
        """
        Determines if the stream with the provided id exists.

        The return of this function is 2 values, the first being a boolean for if anything was found, and the second being how many were found.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Recipient instance has been garbage collected.')
        return msg.existsTypedProperty(id, self.__dir, _type, True, self.__props)

    def getMultipleBinary(self, filename: MSG_PATH) -> Optional[List[bytes]]:
        """
        Gets a multiple binary property as a list of bytes objects.

        Like :meth:`getStringStream`, the 4 character type suffix should be
        omitted. So if you want the stream "__substg1.0_00011102" then the
        filename would simply be "__substg1.0_0001".

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Recipient instance has been garbage collected.')
        return msg.getMultipleBinary([self.__dir, msgPathToString(filename)])

    def getMultipleString(self, filename: MSG_PATH) -> Optional[List[str]]:
        """
        Gets a multiple string property as a list of str objects.

        Like :meth:`getStringStream`, the 4 character type suffix should be
        omitted. So if you want the stream "__substg1.0_00011102" then the
        filename would simply be "__substg1.0_0001".

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Recipient instance has been garbage collected.')
        return msg.getMultipleString([self.__dir, msgPathToString(filename)])

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
        ``instance.props.getValue(name, default)``

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
            raise ReferenceError('The MSGFile for this Recipient instance has been garbage collected.')
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
            raise ReferenceError('The MSGFile for this Recipient instance has been garbage collected.')
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
            raise ReferenceError('The MSGFile for this Recipient instance has been garbage collected.')
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
        None.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Recipient instance has been garbage collected.')
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
        Lists the streams and or storages that exist in the recipient directory.

        :returns: The paths *excluding* the recipient directory, allowing the
            paths to be directly used for accessing a file.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Recipient instance has been garbage collected.')
        return [path[1:] for path in msg.listDir(streams, storages, False)
                if len(path) > 1 and path[0] == self.__dir]

    def slistDir(self, streams: bool = True, storages: bool = False) -> List[str]:
        """
        Like :meth:`listDir`, except it returns the paths as strings.
        """
        return ['/'.join(path) for path in self.listDir(streams, storages)]

    @functools.cached_property
    def account(self) -> Optional[str]:
        """
        The account of this recipient.
        """
        return self.getStringStream('__substg1.0_3A00')

    @property
    def email(self) -> Optional[str]:
        """
        The recipient's email.
        """
        return self.__email

    @functools.cached_property
    def entryID(self) -> Optional[PermanentEntryID]:
        """
        The recipient's Entry ID.
        """
        return self.getStreamAs('__substg1.0_0FFF0102', PermanentEntryID)

    @property
    def formatted(self) -> str:
        """
        The formatted recipient string.
        """
        return self.__formatted

    @functools.cached_property
    def instanceKey(self) -> Optional[bytes]:
        """
        The instance key of this recipient.
        """
        return self.getStream('__substg1.0_0FF60102')

    @property
    def name(self) -> Optional[str]:
        """
        The recipient's name.
        """
        return self.__name

    @property
    def props(self) -> PropertiesStore:
        """
        The Properties instance of the recipient.
        """
        return self.__props

    @functools.cached_property
    def recordKey(self) -> Optional[bytes]:
        """
        The record key of this recipient.
        """
        return self.getStream('__substg1.0_0FF90102')

    @functools.cached_property
    def searchKey(self) -> Optional[bytes]:
        """
        The search key of this recipient.
        """
        return self.getStream('__substg1.0_300B0102')

    @functools.cached_property
    def smtpAddress(self) -> Optional[str]:
        """
        The SMTP address of this recipient.
        """
        return self.getStringStream('__substg1.0_39FE')

    @functools.cached_property
    def transmittableDisplayName(self) -> Optional[str]:
        """
        The transmittable display name of this recipient.
        """
        return self.getStringStream('__substg1.0_3A20')

    @property
    def type(self) -> _RT:
        """
        The recipient type.
        """
        return self.__type

    @property
    def typeFlags(self) -> int:
        """
        The raw recipient type value and all the flags it includes.
        """
        return self.__typeFlags
