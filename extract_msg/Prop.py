from extract_msg import constants
from extract_msg.debug import debug
from extract_msg.utils import properHex



class Prop(object):
    """
    Class to contain the data for a single property.

    Currently a work in progress.
    """

    def __init__(self, string):
        n = string[:4][::-1]
        self.__raw = string
        self.__name = properHex(n).upper()
        self.__type, self.__flags, self.__value = constants.ST2.unpack(string)
        self.__value = self.parse_type(self.__type, self.__value)
        self.__fm = self.__flags & 1 == 1
        self.__fr = self.__flags & 2 == 2
        self.__fw = self.__flags & 4 == 4

    def parse_type(self, _type, stream):
        """
        Converts the data in :param stream: to a
        much more accurate type, specified by
        :param _type:, if possible.

        WARNING: Not done.
        """
        # WARNING Not done.
        value = stream
        if _type == 0x0000:  # PtypUnspecified
            pass;
        elif _type == 0x0001:  # PtypNull
            if value != b'\x00\x00\x00\x00\x00\x00\x00\x00':
                print('Warning: Property type is PtypNull, but is not equal to 0.')
            value = None
        elif _type == 0x0002:  # PtypInteger16
            value = constants.STI16.unpack(value)[0]
        elif _type == 0x0003:  # PtypInteger32
            value = constants.STI32.unpack(value)[0]
        elif _type == 0x0004:  # PtypFloating32
            value = constants.STF32.unpack(value)[0]
        elif _type == 0x0005:  # PtypFloating64
            value = constants.STF64.unpack(value)[0]
        elif _type == 0x0006:  # PtypCurrency
            value = (constants.STI64.unpack(value))[0] / 10000.0
        elif _type == 0x0007:  # PtypFloatingTime
            value = constants.STF64.unpack(value)[0]
            # TODO parsing for this
            pass;
        elif _type == 0x000A:  # PtypErrorCode
            value = constants.STI32.unpack(value)[0]
            # TODO parsing for this
            pass;
        elif _type == 0x000B:  # PtypBoolean
            value = bool(constants.ST3.unpack(value)[0])
        elif _type == 0x000D:  # PtypObject/PtypEmbeddedTable
            # TODO parsing for this
            pass;
        elif _type == 0x0014:  # PtypInteger64
            value = constants.STI64.unpack(value)[0]
        elif _type == 0x001E:  # PtypString8
            # TODO parsing for this
            pass;
        elif _type == 0x001F:  # PtypString
            # TODO parsing for this
            pass;
        elif _type == 0x0040:  # PtypTime
            value = constants.ST3.unpack(value)[0]
        elif _type == 0x0048:  # PtypGuid
            # TODO parsing for this
            pass;
        elif _type == 0x00FB:  # PtypServerId
            # TODO parsing for this
            pass;
        elif _type == 0x00FD:  # PtypRestriction
            # TODO parsing for this
            pass;
        elif _type == 0x00FE:  # PtypRuleAction
            # TODO parsing for this
            pass;
        elif _type == 0x0102:  # PtypBinary
            # TODO parsing for this
            # Smh, how on earth am I going to code this???
            pass;
        elif _type & 0x1000 == 0x1000:  # PtypMultiple
            # TODO parsing for `multiple` types
            pass;
        return value;

    @property
    def flag_mandatory(self):
        """
        Boolean, is the "mandatory" flag set?
        """
        return self.__fm

    @property
    def flag_readable(self):
        """
        Boolean, is the "readable" flag set?
        """
        return self.__fr

    @property
    def flag_writable(self):
        """
        Boolean, is the "writable" flag set?
        """
        return self.__fw

    @property
    def flags(self):
        """
        Integer that contains property flags.
        """
        return self.__flags

    @property
    def name(self):
        """
        Property "name".
        """
        return self.__name

    @property
    def raw(self):
        """
        Raw binary string that defined the property.
        """
        return self.__raw

    @property
    def value(self):
        """
        Property value.
        """
        return self.__value
