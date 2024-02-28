from __future__ import annotations


__all__ = [
    'PropertiesStore',
]


import copy
import datetime
import logging
import pprint

from typing import (
        Any, Dict, Iterable, Iterator, List, Optional, Tuple, TypeVar, Union
    )

from .. import constants
from ..enums import PropertiesType
from ..exceptions import NotWritableError
from .prop import createProp, FixedLengthProp, PropBase
from ..utils import divide


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

_T = TypeVar('_T')


class PropertiesStore:
    """
    Parser for msg properties files.
    """

    def __init__(self, data: Optional[bytes], type_: PropertiesType, writable: bool = False):
        """
        Reads a properties stream or creates a brand new ``PropertiesStore``
        object.

        :param data: The bytes of the properties instance. Setting to ``None``
            or empty bytes will cause the properties instance to not be valid
            *unless* writable is set to ``True``. If that is the case, the
            instance will be setup for creating a new properties stream.
        :param type_: The type of properties stream this instance represents.
        :param writable: Whether this properties stream should accept
            modification.
        """
        if not isinstance(type_, PropertiesType):
            raise TypeError(':param type_: MUST be a value of PropertiesType.')

        self.__type = type_

        # Setup early variables.
        self.__props: Dict[str, PropBase] = {}
        # This maps short IDs to all properties that use that ID. More than one
        # property with the same ID but a different type may exist.
        self.__idMapping: Dict[str, List[str]] = {}
        self.__naid = None
        self.__nrid = None
        self.__ac = None
        self.__rc = None
        self.__writable = writable
        # Set this now and unset it if everything goes well.
        self.__isError = True

        # Check if data is None or empty bytes.
        if not data:
            # Two paths here. If writable, we are just meant to be creating a
            # new storage for properties, so initialize and return. Otherwise,
            # we're dealing with an error situation, but we want to silence it.
            if writable:
                self.__rawData = b''
                if type_ is not PropertiesType.ATTACHMENT:
                    self.__naid = 0
                    self.__nrid = 0
                    self.__ac = 0
                    self.__rc = 0
            return

        if not isinstance(data, bytes):
            raise TypeError(':param data: MUST be bytes or None.')
        self.__rawData = data

        if type_ == PropertiesType.MESSAGE:
            skip = 32
            self.__nrid, self.__naid, self.__rc, self.__ac = constants.st.ST_PROPSTORE_HEADER.unpack(data[:24])
        elif type_ == PropertiesType.MESSAGE_EMBED:
            skip = 24
            self.__nrid, self.__naid, self.__rc, self.__ac = constants.st.ST_PROPSTORE_HEADER.unpack(data[:24])
        else:
            skip = 8
        streams = divide(self.__rawData[skip:], 16)
        for st in streams:
            if len(st) == 16:
                prop = createProp(st)
                self.__props[prop.name] = prop

                # Add the ID to our mapping list.
                id_ = prop.name[:4]
                if id_ not in self.__idMapping:
                    self.__idMapping[id_] = []
                self.__idMapping[id_].append(prop.name)
            else:
                logger.warning(f'Found stream from divide that was not 16 bytes: {st}. Ignoring.')
        self.__isError = False

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def __contains__(self, key: Any) -> bool:
        return self.__props.__contains__(key)

    def __delitem__(self, key: str) -> None:
        """
        Removes an item using the del operator.

        :raises KeyError: The key was not found.
        :raises TypeError: The key was not a string.
        """

        if not isinstance(key, str):
            raise TypeError('Del operator can only remove a property by string.')

        key = key.upper()

        del self.__props[key]

        # If the deletion was successful, we need to remove the related ID
        # mapping.
        shortKey = key[:4]
        self.__idMapping[shortKey].remove(key)
        if len(self.__idMapping[shortKey]) == 0:
            del self.__idMapping[shortKey]

    def __getitem__(self, key: Union[str, int]) -> PropBase:
        if (found := self._mapId(key)):
            return self.__props.__getitem__(found)
        raise KeyError(key)

    def __iter__(self) -> Iterator[str]:
        return self.__props.__iter__()

    def __len__(self) -> int:
        """
        Returns the number of properties.
        """
        return len(self.__props)

    def __repr__(self) -> str:
        return self.__props.__repr__()

    def _mapId(self, id_: Union[int, str]) -> str:
        """
        Converts an input into an appropriate property ID.

        This is a complex function, allowing the user to specify an int or
        string. If the input is a string that is not 4 characters, it is
        returned. Otherwise, a series of checks will be
        performed. If the input is an int that is less than ``0x10000``, it is
        considered a property ID without a type and converted to a 4 character
        hexadecimal string. Otherwise, it is converted to an 8 character
        hexadecimal string and returned.

        Once the input is a 4 character string from the other paths, it will
        then be checked against the list of found property IDs, and the first
        full ID will be returned.

        If a valid conversion could not be done, returns an empty string.

        All strings returned will be uppercase.
        """
        # See if we need to convert to 4 character string and map or if we just
        # need to return quickly.
        if isinstance(id_, str):
            id_ = id_.upper()
            if len(id_) != 4:
                return id_
        elif isinstance(id_, int):
            if id_ >= 0x10000:
                return f'{id_:08X}'
            else:
                id_ = f'{id_:04X}'
        else:
            return ''

        return self.__idMapping.get(id_, ('',))[0]

    def addProperty(self, prop: PropBase, force: bool = False) -> None:
        """
        Adds the property if it does not exist.

        :param prop: The property to add.
        :param force: If ``True``, the writable property will be ignored. This
            will not be reflected when converting to ``bytes`` if the instance
            is not readable.

        :raises KeyError: A property already exists with the chosen name.
        :raises NotWritableError: The method was used on an unwritable instance.
        """
        if not (force or self.__writable):
            raise

        if prop.name in self.__props:
            raise KeyError('A property with that name already exists.')
        self.__props[prop.name.upper()] = prop
        self.__idMapping.setdefault(prop.name[:4], list()).append(prop.name.upper())

    def get(self, name: Union[str, int], default: _T = None) -> Union[PropBase, _T]:
        """
        Retrieve the property of :param name:.

        :returns: The property, or the value of :param default: if the property
            could not be found.
        """
        if (name := self._mapId(name)):
            return self.__props.get(name, default)
        else:
            return default

    def getProperties(self, id_: Union[str, int]) -> List[PropBase]:
        """
        Gets all properties with the specified ID.

        :param ID: An 4 digit hexadecimal string or an int that is less than
            0x10000.
        """
        if isinstance(id_, int):
            if id_ >= 0x10000:
                return []
            else:
                id_ = f'{id_:04X}'
        elif isinstance(id_, str):
            if len(id_) == 4:
                id_ = id_.upper()
            else:
                return []

        return [self[x] for x in self.__idMapping.get(id_, [])]

    def getValue(self, name: Union[str, int], default: _T = None) -> Union[Any, _T]:
        """
        Attempts to get the first property
        """
        if isinstance(name, int):
            if name >= 0x10000:
                name = f'{name:08X}'
            else:
                name = f'{name:04X}'
        if len(name) == 4:
            for prop in self.getProperties(name):
                if isinstance(prop, FixedLengthProp):
                    return prop.value
            return default
        elif len(name) == 8:
            if (prop := self.get(name)):
                if isinstance(prop, FixedLengthProp):
                    return prop.value
                else:
                    return default
            return default
        else:
            raise ValueError('Property name must be an int less than 0x100000000, a 4 character hex string, or an 8 character hex string.')

    def items(self) -> Iterable[Tuple[str, PropBase]]:
        return self.__props.items()

    def keys(self) -> Iterable[str]:
        return self.__props.keys()

    def makeWritable(self) -> PropertiesStore:
        """
        Returns a copy of this PropertiesStore object that allows modification.

        If the instance is already writable, this will return the object.
        """
        if self.__writable:
            return self
        return PropertiesStore(self.__rawData, self.__type, True)

    def pprintKeys(self) -> None:
        """
        Uses the pprint function on a sorted list of the keys.
        """
        pprint.pprint(sorted(self.__props.keys()))

    def removeProperty(self, nameOrProp: Union[str, PropBase]) -> None:
        """
        Removes the property by name or by instance.

        Due to possible ambiguities, this function does *not* accept an int
        argument nor will it be able to find a property based on the 4 character
        hex ID.

        :raises KeyError: The property was not found.
        :raises NotWritableError: The instance is not writable.
        :raises TypeError: The type for :param nameOrProp: was wrong.
        """
        if isinstance(nameOrProp, str):
            del self[nameOrProp]
        elif isinstance(nameOrProp, PropBase):
            del self[nameOrProp.name]
        else:
            raise TypeError(f'Cannot remove property using type {type(nameOrProp)}.')

    def toBytes(self) -> bytes:
        if self.__writable:
            # The reserved field is present on all of them.
            ret = b'\x00' * 8

            # Add additional fields depending on type.
            if self.__type is not PropertiesType.ATTACHMENT:
                ret += constants.st.ST_PROPSTORE_HEADER.pack(self.__nrid, self.__naid, self.__rc, self.__ac)
                if self.__type is PropertiesType.MESSAGE:
                    ret += b'\x00' * 8

            # Convert all the properties to bytes.
            ret += b''.join(bytes(prop) for prop in self.__props.values())

            return ret
        else:
            return self.__rawData

    def values(self) -> Iterable[PropBase]:
        return self.__props.values()

    items.__doc__ = dict.items.__doc__
    keys.__doc__ = dict.keys.__doc__
    values.__doc__ = dict.values.__doc__

    @property
    def attachmentCount(self) -> int:
        """
        The number of Attachment objects for the ``MSGFile`` object.

        :raises NotWritableError: The setter was used on an unwritable instance.
        :raises TypeError: The Properties instance is not for an ``MSGFile``
            object.
        """
        if self.__ac is None:
            raise TypeError('Attachment properties do not contain an attachment count.')
        return self.__ac

    @attachmentCount.setter
    def attachmentCount(self, value: int) -> None:
        if not self.__writable:
            raise NotWritableError('PropertiesStore object is not writable.')

        if not isinstance(value, int):
            raise TypeError(':property attachmentCount: must be an int.')

        if self.__ac is None:
            raise TypeError('Attachment properties do not contain an attachment count.')

        self.__ac = value

    @property
    def date(self) -> Optional[datetime.datetime]:
        """
        Returns the send date contained in the Properties file.
        """
        try:
            return self.__date
        except AttributeError:
            self.__date = None
            if '00390040' in self:
                dateValue = self.getValue('00390040')
                # A date can be bytes if it fails to initialize, so we check it
                # first.
                if isinstance(dateValue, datetime.datetime):
                    self.__date = dateValue
            return self.__date

    @property
    def isError(self) -> bool:
        """
        Whether the instance is in an invalid state.

        If the instance is not writable and was given no data, this will be
        ``True``.
        """
        return self.__isError

    @property
    def nextAttachmentId(self) -> int:
        """
        The ID to use for naming the next Attachment object storage if one is
        created inside the .msg file.

        :raises NotWritableError: The setter was used on an unwritable instance.
        :raises TypeError: The Properties instance is not for an ``MSGFile``
            object.
        """
        if self.__naid is None:
            raise TypeError('Attachment properties do not contain a next attachment ID.')
        return self.__naid

    @nextAttachmentId.setter
    def nextAttachmentId(self, value: int) -> None:
        if not self.__writable:
            raise NotWritableError('PropertiesStore object is not writable.')

        if not isinstance(value, int):
            raise TypeError(':property nextAttachmentId: must be an int.')

        if self.__ac is None:
            raise TypeError('Attachment properties do not contain a next attachment ID.')

        self.__naid = value

    @property
    def nextRecipientId(self) -> int:
        """
        The ID to use for naming the next Recipient object storage if one is
        created inside the .msg file.

        :raises NotWritableError: The setter was used on an unwritable instance.
        :raises TypeError: The Properties instance is not for an ``MSGFile``
            object.
        """
        if self.__nrid is None:
            raise TypeError('Attachment properties do not contain a next recipient ID.')
        return self.__nrid

    @nextRecipientId.setter
    def nextRecipientId(self, value: int) -> None:
        if not self.__writable:
            raise NotWritableError('PropertiesStore object is not writable.')

        if not isinstance(value, int):
            raise TypeError(':property nextRecipientId: must be an int.')

        if self.__ac is None:
            raise TypeError('Attachment properties do not contain a next recipient ID.')

        self.__nrid = value

    @property
    def props(self) -> Dict[str, PropBase]:
        """
        Returns a copy of the internal properties dict.
        """
        return copy.deepcopy(self.__props)

    @property
    def recipientCount(self) -> int:
        """
        The number of Recipient objects for the ``MSGFile`` object.

        :raises NotWritableError: The setter was used on an unwritable instance.
        :raises TypeError: The Properties instance is not for an ``MSGFile``
            object.
        """
        if self.__rc is None:
            raise TypeError('Attachment properties do not contain a recipient count.')
        return self.__rc

    @recipientCount.setter
    def recipientCount(self, value: int) -> None:
        if not self.__writable:
            raise NotWritableError('PropertiesStore object is not writable.')

        if not isinstance(value, int):
            raise TypeError(':property recipientCount: must be an int.')

        if self.__ac is None:
            raise TypeError('Attachment properties do not contain a recipient count.')

        self.__nrid = value

    @property
    def writable(self) -> bool:
        """
        Whether the instance accepts modification.
        """
        return self.__writable
