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
from ..enums import Intelligence, PropertiesType
from .prop import createProp, FixedLengthProp, PropBase
from ..utils import divide


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

_T = TypeVar('_T')


class PropertiesStore:
    """
    Parser for msg properties files.
    """

    def __init__(self, data : Optional[bytes] = b'', _type : Optional[PropertiesType] = None, skip : Optional[int] = None):
        if not data:
            # If data comes back false, make sure is is empty bytes.
            data = b''
        if not isinstance(data, bytes):
            raise TypeError(':param data: MUST be bytes.')
        self.__rawData = data
        self.__len = len(data)
        self.__props : Dict[str, PropBase] = {}
        # This maps short IDs to all properties that use that ID. More than one
        # property with the same ID but a different type may exist.
        self.__idMapping : Dict[str, List[str]] = {}
        self.__naid = None
        self.__nrid = None
        self.__ac = None
        self.__rc = None
        # Handle an empty properties stream.
        if self.__len == 0:
            self.__intel = Intelligence.ERROR
            skip = 0
        elif _type is not None:
                _type = PropertiesType(_type)
                self.__intel = Intelligence.SMART
                if _type == PropertiesType.MESSAGE:
                    skip = 32
                    self.__nrid, self.__naid, self.__rc, self.__ac = constants.st.ST1.unpack(self.__rawData[:24])
                elif _type == PropertiesType.MESSAGE_EMBED:
                    skip = 24
                    self.__nrid, self.__naid, self.__rc, self.__ac = constants.st.ST1.unpack(self.__rawData[:24])
                else:
                    skip = 8
        else:
            self.__intel = Intelligence.DUMB
            if skip is None:
                # This section of the skip handling is not very good. While it
                # does work, it is likely to create extra properties that are
                # created from the properties file's header data. While that
                # won't actually mess anything up, it is far from ideal.
                # Basically, this is the dumb skip length calculation.
                # Preferably, we want the type to have been specified so all of
                # the additional fields will have been filled out.
                #
                # If the skip would end up at 0, set it to 32.
                skip = (self.__len % 16) or 32
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
        self.__pl = len(self.__props)

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def __contains__(self, key) -> bool:
        return self.__props.__contains__(key)

    def __getitem__(self, key : Union[str, int]) -> PropBase:
        if (found := self._mapId(key)):
            return self.__props.__getitem__(found)
        raise KeyError(key)

    def __iter__(self) -> Iterator[str]:
        return self.__props.__iter__()

    def __len__(self) -> int:
        """
        Returns the number of properties.
        """
        return self.__pl

    def __repr__(self) -> str:
        return self.__props.__repr__()

    def _mapId(self, id_ : Union[int, str]) -> str:
        """
        Converts an input into an appropriate property ID.

        This is a complex function, allowing the user to specify an int or
        string. If the input is a string that is not 4 characters, it is
        returned. Otherwise, a series of checks will be
        performed. If the input is an int that is less than 0x10000, it is
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

    def get(self, name : Union[str, int], default : _T = None) -> Union[PropBase, _T]:
        """
        Retrieve the property of :param name:. Returns the value of
        :param default: if the property could not be found.
        """
        if (name := self._mapId(name)):
            return self.__props.get(name, default)
        else:
            return default

    def getProperties(self, id_ : Union[str, int]) -> List[PropBase]:
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

    def getValue(self, name : Union[str, int], default : _T = None) -> Union[Any, _T]:
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

    def pprintKeys(self) -> None:
        """
        Uses the pprint function on a sorted list of keys.
        """
        pprint.pprint(sorted(tuple(self.__props.keys())))

    def toBytes(self) -> bytes:
        return self.__rawData

    def values(self) -> Iterable[PropBase]:
        return self.__props.values()

    items.__doc__ = dict.items.__doc__
    keys.__doc__ = dict.keys.__doc__
    values.__doc__ = dict.values.__doc__

    @property
    def attachmentCount(self) -> int:
        """
        The number of Attachment objects for the MSGFile object.

        :raises TypeError: The Properties instance is not for an MSGFile object.
        """
        if self.__ac is None:
            raise TypeError('Properties instance must be intelligent and of type MESSAGE to get attachment count.')
        return self.__ac

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
                dateValue = self.get('00390040').value
                # A date can be bytes if it fails to initialize, so we check it
                # first.
                if isinstance(dateValue, datetime.datetime):
                    self.__date = dateValue
            return self.__date

    @property
    def intelligence(self) -> Intelligence:
        """
        Returns the inteligence level of the Properties instance.
        """
        return self.__intel

    @property
    def nextAttachmentId(self) -> int:
        """
        The ID to use for naming the next Attachment object storage if one is
        created inside the .msg file.

        :raises TypeError: The Properties instance is not for an MSGFile object.
        """
        if self.__naid is None:
            raise TypeError('Properties instance must be intelligent and of type MESSAGE to get next attachment id.')
        return self.__naid

    @property
    def nextRecipientId(self) -> int:
        """
        The ID to use for naming the next Recipient object storage if one is
        created inside the .msg file.

        :raises TypeError: The Properties instance is not for an MSGFile object.
        """
        if self.__nrid is None:
            raise TypeError('Properties instance must be intelligent and of type MESSAGE to get next recipient id.')
        return self.__nrid

    @property
    def props(self) -> Dict[str, PropBase]:
        """
        Returns a copy of the internal properties dict.
        """
        return copy.deepcopy(self.__props)

    @property
    def _propDict(self) -> Dict[str, PropBase]:
        """
        A direct reference to the underlying property dictionary. Used in one
        place in the code, and not recommended to be used if you are not a
        developer. Use `Properties.props` instead for a safe reference.
        """
        return self.__props

    @property
    def recipientCount(self) -> int:
        """
        The number of Recipient objects for the MSGFile object.

        :raises TypeError: The Properties instance is not for an MSGFile object.
        """
        if self.__rc is None:
            raise TypeError('Properties instance must be intelligent and of type MESSAGE to get recipient count.')
        return self.__rc
