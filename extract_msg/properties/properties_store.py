__all__ = [
    'PropertiesStore',
]


import copy
import datetime
import logging
import pprint

from typing import Any, Dict, Optional, Union
from warnings import warn

from .. import constants
from ..enums import Intelligence, PropertiesType
from .prop import createProp, PropBase
from ..utils import divide, properHex


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class PropertiesStore:
    """
    Parser for msg properties files.
    """

    def __init__(self, data : Optional[bytes], _type : Optional[PropertiesType] = None, skip : Optional[int] = None):
        if not data:
            # If data comes back false, make sure is is empty bytes.
            data = b''
        if not isinstance(data, bytes):
            raise TypeError(':param data: MUST be bytes.')
        self.__rawData = data
        self.__len = len(data)
        self.__props : Dict[PropBase] = {}
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
            else:
                logger.warning(f'Found stream from divide that was not 16 bytes: {st}. Ignoring.')
        self.__pl = len(self.__props)

    def __contains__(self, key) -> bool:
        return self.__props.__contains__(key)

    def __getitem__(self, key) -> PropBase:
        return self.__props.__getitem__(key)

    def __iter__(self):
        return self.__props.__iter__()

    def __len__(self) -> int:
        """
        Returns the number of properties.
        """
        return self.__pl

    def __repr__(self) -> str:
        return self.__props.__repr__()

    def get(self, name, default = None) -> Optional[Union[PropBase, Any]]:
        """
        Retrieve the property of :param name:. Returns the value of
        :param default: if the property could not be found.
        """
        try:
            return self.__props[name]
        except KeyError:
            # DEBUG
            logger.debug('KeyError exception.')
            logger.debug(properHex(self.__rawData))
            logger.debug(self.__props)
            return default

    def has_key(self, key) -> bool:
        """
        Checks if :param key: is a key in the properties dictionary.
        """
        warn('`Properties.has_key` is deprecated. Use the `in` keyword instead.', DeprecationWarning)
        return key in self.__props

    def items(self):
        return self.__props.items()

    def keys(self):
        return self.__props.keys()

    def pprintKeys(self) -> None:
        """
        Uses the pprint function on a sorted list of keys.
        """
        pprint.pprint(sorted(tuple(self.__props.keys())))

    def values(self):
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
                # A date can by bytes if it fails to initialize, so we check it
                # first.
                if isinstance(dateValue, datetime.datetime):
                    self.__date = dateValue.__format__('%a, %d %b %Y %H:%M:%S %z')
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
    def props(self) -> Dict:
        """
        Returns a copy of the internal properties dict.
        """
        return copy.deepcopy(self.__props)

    @property
    def _propDict(self) -> Dict:
        """
        A direct reference to the underlying property dictionary. Used in one
        place in the code, and not recommended to be used if you are not a
        developer. Use `Properties.props` instead for a safe reference.
        """
        return self.__props

    @property
    def rawData(self) -> bytes:
        """
        The raw bytes used to create this object.
        """
        return self.__rawData

    @property
    def recipientCount(self) -> int:
        """
        The number of Recipient objects for the MSGFile object.

        :raises TypeError: The Properties instance is not for an MSGFile object.
        """
        if self.__rc is None:
            raise TypeError('Properties instance must be intelligent and of type MESSAGE to get recipient count.')
        return self.__rc
