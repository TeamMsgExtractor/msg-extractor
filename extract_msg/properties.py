import copy
import logging
import pprint

from extract_msg import constants
from extract_msg.prop import create_prop
from extract_msg.utils import divide, fromTimeStamp, msgEpoch, properHex

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class Properties(object):
    """
    Parser for msg properties files.
    """

    def __init__(self, stream, type=None, skip=None):
        object.__init__(self)
        self.__stream = stream
        self.__pos = 0
        self.__len = len(stream)
        self.__props = {}
        self.__naid = None
        self.__nrid = None
        self.__ac = None
        self.__rc = None
        if type is not None:
            self.__intel = constants.INTELLIGENCE_SMART
            if type == constants.TYPE_MESSAGE:
                skip = 32
                self.__naid, self.__nrid, self.__ac, self.__rc = constants.ST1.unpack(self.__stream[:24])
            elif type == constants.TYPE_MESSAGE_EMBED:
                skip = 24
                self.__naid, self.__nrid, self.__ac, self.__rc = constants.ST1.unpack(self.__stream[:24])
            else:
                skip = 8
        else:
            self.__intel = constants.INTELLIGENCE_DUMB
            if skip is None:
                # This section of the skip handling is not very good.
                # While it does work, it is likely to create extra
                # properties that are created from the properties file's
                # header data. While that won't actually mess anything
                # up, it is far from ideal. Basically, this is the dumb
                # skip length calculation. Preferably, we want the type
                # to have been specified so all of the additional fields
                # will have been filled out
                skip = self.__len % 16
                if skip == 0:
                    skip = 32
        streams = divide(self.__stream[skip:], 16)
        for st in streams:
            a = create_prop(st)
            self.__props[a.name] = a
        self.__pl = len(self.__props)

    def __contains__(self, key):
        self.__props.__contains__(key)

    def __getitem__(self, key):
        return self.__props.__getitem__(key)

    def __iter__(self):
        return self.__props.__iter__()

    def __len__(self):
        """
        Returns the number of properties.
        """
        return self.__pl

    @property
    def __repr__(self):
        return self.__props.__repr__

    def get(self, name):
        """
        Retrieve the property of :param name:.
        """
        try:
            return self.__props[name]
        except KeyError:
            # DEBUG
            logger.debug('KeyError exception.')
            logger.debug(properHex(self.__stream))
            logger.debug(self.__props)
            raise

    def has_key(self, key):
        """
        Checks if :param key: is a key in the properties dictionary.
        """
        return key in self.__props

    def items(self):
        return self.__props.items()

    def keys(self):
        return self.__props.keys()

    def pprintKeys(self):
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
    def attachment_count(self):
        if self.__ac is None:
            raise TypeError('Properties instance must be intelligent and of type TYPE_MESSAGE to get attachment count.')
        return self.__ac

    @property
    def date(self):
        """
        Returns the send date contained in the Properties file.
        """
        try:
            return self.__date
        except AttributeError:
            if self.has_key('00390040'):
                self.__date = self.get('00390040').value.__format__('%a, %d %b %Y %H:%M:%S %z')
            elif self.has_key('30080040'):
                self.__date = self.get('30080040').value.__format__('%a, %d %b %Y %H:%M:%S %z')
            elif self.has_key('30070040'):
                self.__date = self.get('30070040').value.__format__('%a, %d %b %Y %H:%M:%S %z')
            else:
                # DEBUG
                logger.warning(
                    'Error retrieving date. Setting as "Unknown". Please send the following data to developer:\n--------------------')
                logger.warning(properHex(self.__stream))
                logger.warning(self.keys())
                logger.warning('--------------------')
                self.__date = 'Unknown'
            return self.__date

    @property
    def intelligence(self):
        """
        Returns the inteligence level of the Properties instance.
        """
        return self.__intel

    @property
    def next_attachment_id(self):
        if self.__naid is None:
            raise TypeError(
                'Properties instance must be intelligent and of type TYPE_MESSAGE to get next attachment id.')
        return self.__naid

    @property
    def next_recipient_id(self):
        if self.__nrid is None:
            raise TypeError(
                'Properties instance must be intelligent and of type TYPE_MESSAGE to get next recipient id.')
        return self.__nrid

    @property
    def props(self):
        """
        Returns a copy of the internal properties dict.
        """
        return copy.deepcopy(self.__props)

    @property
    def recipient_count(self):
        if self.__rc is None:
            raise TypeError('Properties instance must be intelligent and of type TYPE_MESSAGE to get recipient count.')
        return self.__rc

    @property
    def stream(self):
        """
        Returns the data stream used to generate this Properties instance.
        """
        return self.__stream
