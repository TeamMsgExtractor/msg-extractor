from __future__ import annotations


__all__ = [
    # Classes:
    'FixedLengthProp'
    'PropBase',
    'VariableLengthProp',

    # Functions:
    'createProp',
]


import datetime
import logging

from typing import Any

from .. import constants
from ..enums import ErrorCode, ErrorCodeType
from ..utils import filetimeToDatetime, properHex


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


def createProp(data : bytes) -> PropBase:
    temp = constants.st.ST2.unpack(data)[0]
    if temp in constants.FIXED_LENGTH_PROPS:
        return FixedLengthProp(data)
    else:
        if temp not in constants.VARIABLE_LENGTH_PROPS:
            # DEBUG.
            logger.warning(f'Unknown property type: {properHex(temp)}')
        return VariableLengthProp(data)


class PropBase:
    """
    Base class for Prop instances.
    """

    def __init__(self, data : bytes):
        self.__rawData = data
        self.__name = properHex(data[3::-1]).upper()
        self.__type, self.__flags = constants.st.ST2.unpack(data)
        self.__fm = self.__flags & 1 == 1
        self.__fr = self.__flags & 2 == 2
        self.__fw = self.__flags & 4 == 4

    @property
    def flagMandatory(self) -> bool:
        """
        Boolean, is the "mandatory" flag set?
        """
        return self.__fm

    @property
    def flagReadable(self) -> bool:
        """
        Boolean, is the "readable" flag set?
        """
        return self.__fr

    @property
    def flagWritable(self) -> bool:
        """
        Boolean, is the "writable" flag set?
        """
        return self.__fw

    @property
    def flags(self) -> int:
        """
        Integer that contains property flags.
        """
        return self.__flags

    @property
    def name(self) -> str:
        """
        Property "name".
        """
        return self.__name

    @property
    def rawData(self) -> bytes:
        """
        The raw bytes used to create this object.
        """
        return self.__rawData

    @property
    def type(self) -> int:
        """
        The type of property.
        """
        return self.__type



class FixedLengthProp(PropBase):
    """
    Class to contain the data for a single fixed length property.

    Currently a work in progress.
    """

    def __init__(self, data : bytes):
        super().__init__(data)
        self.__value = self.parseType(self.type, constants.st.STFIX.unpack(data)[0])

    def parseType(self, _type : int, stream : bytes) -> Any:
        """
        Converts the data in :param stream: to a much more accurate type,
        specified by :param _type:, if possible.
        :param stream: The data that the value is extracted from.

        WARNING: Not done.
        """
        # WARNING Not done.
        value = stream
        if _type == 0x0000: # PtypUnspecified
            pass
        elif _type == 0x0001: # PtypNull
            if value != b'\x00\x00\x00\x00\x00\x00\x00\x00':
                # DEBUG.
                logger.warning('Property type is PtypNull, but is not equal to 0.')
            value = None
        elif _type == 0x0002: # PtypInteger16
            value = constants.st.STI16.unpack(value)[0]
        elif _type == 0x0003: # PtypInteger32
            value = constants.st.STI32.unpack(value)[0]
        elif _type == 0x0004: # PtypFloating32
            value = constants.st.STF32.unpack(value)[0]
        elif _type == 0x0005: # PtypFloating64
            value = constants.st.STF64.unpack(value)[0]
        elif _type == 0x0006: # PtypCurrency
            value = (constants.st.STI64.unpack(value))[0] / 10000.0
        elif _type == 0x0007: # PtypFloatingTime
            value = constants.st.STF64.unpack(value)[0]
            return constants.PYTPFLOATINGTIME_START + datetime.timedelta(days = value)
        elif _type == 0x000A: # PtypErrorCode
            value = constants.st.STI32.unpack(value)[0]
            try:
                value = ErrorCodeType(value)
            except ValueError:
                logger.warning(f'Error type found that was not from Additional Error Codes. Value was {value}. You should report this to the developers.')
                # So here, the value should be from Additional Error Codes, but
                # it wasn't. So we are just returning the int. However, we want
                # to see if it is a normal error code.
                try:
                    logger.warning(f'REPORT TO DEVELOPERS: Error type of {ErrorCode(value)} was found.')
                except ValueError:
                    pass
        elif _type == 0x000B:  # PtypBoolean
            value = constants.st.ST3.unpack(value)[0] == 1
        elif _type == 0x0014:  # PtypInteger64
            value = constants.st.STI64.unpack(value)[0]
        elif _type == 0x0040:  # PtypTime
            rawTime = constants.st.ST3.unpack(value)[0]
            try:
                value = filetimeToDatetime(rawTime)
            except ValueError as e:
                logger.exception(e)
                logger.error(self.rawData)
        elif _type == 0x0048:  # PtypGuid
            # TODO parsing for this
            pass
        return value

    @property
    def value(self) -> Any:
        """
        Property value.
        """
        return self.__value



class VariableLengthProp(PropBase):
    """
    Class to contain the data for a single variable length property.
    """

    def __init__(self, data : bytes):
        super().__init__(data)
        self.__length, self.__reserved = constants.st.STVAR.unpack(data)
        if self.type == 0x001E:
            self.__realLength = self.__length - 1
        elif self.type == 0x001F:
            self.__realLength = self.__length - 2
        elif self.type in constants.MULTIPLE_2_BYTES_HEX:
            self.__realLength = self.__length // 2
        elif self.type in constants.MULTIPLE_4_BYTES_HEX:
            self.__realLength = self.__length // 4
        elif self.type in constants.MULTIPLE_8_BYTES_HEX:
            self.__realLength = self.__length // 8
        elif self.type in constants.MULTIPLE_16_BYTES_HEX:
            self.__realLength = self.__length // 16
        elif self.type == 0x000D:
            self.__realLength = None
        else:
            self.__realLength = self.__length

    @property
    def length(self) -> int:
        """
        The length field of the variable length property.
        """
        return self.__length

    @property
    def realLength(self) -> int:
        """
        The ACTUAL length of the stream that this property corresponds to.
        """
        return self.__realLength

    @property
    def reservedFlags(self) -> int:
        """
        The reserved flags field of the variable length property.
        """
        return self.__reserved
