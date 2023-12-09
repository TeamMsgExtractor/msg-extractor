from __future__ import annotations


__all__ = [
    # Classes:
    'FixedLengthProp',
    'PropBase',
    'VariableLengthProp',

    # Functions:
    'createProp',
    'createNewProp',
]


import abc
import datetime
import decimal
import logging

from typing import Any, Dict, Type

from .. import constants
from ..enums import ErrorCode, ErrorCodeType, PropertyFlags
from ..utils import filetimeToDatetime


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

# Define default values to use when creating each prop type. Only define a value
# if it would not be all null bytes.
_DEFAULT_PROP_VALS: Dict[str, bytes] = {
    '000D': b'\x00\x00\x00\x00\xFF\xFF\xFF\xFF\x00\x00\x00\x00',
    '001E': b'\x00\x00\x00\x00\x01\x00\x00\x00\x00\x00\x00\x00',
    '001F': b'\x00\x00\x00\x00\x02\x00\x00\x00\x00\x00\x00\x00',
    '0048': b'\x00\x00\x00\x00\x10\x00\x00\x00\x00\x00\x00\x00',
    '0000': b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
    '0000': b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
    '0000': b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
    '0000': b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
    '0000': b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
}


def createNewProp(name: str):
    """
    Creates a blank property using the specified name.

    :param name: An 8 character hex string containing the property ID and type.

    :raises TypeError: A type other than a str was given.
    :raises ValueError: The string was not 8 hex characters.
    :raises ValueError: An invalid property type was given.
    """
    if not isinstance(name, str):
        raise TypeError(':param name: MUST be a str.')
    if len(name) != 8:
        raise ValueError(':param name: MUST be 8 characters.')

    propVal = bytes.fromhex(name)[::-1]
    propVal += _DEFAULT_PROP_VALS.get(name[:4], b'\x00' * 12)
    return createProp(propVal)


def createProp(data: bytes) -> PropBase:
    """
    Creates an instance of PropBase from the specified bytes.

    If the prop type is not recognized, a VariableLengthProp will be created.
    """
    temp = constants.st.ST_PROP_BASE.unpack(data[:8])[0]
    if temp in constants.FIXED_LENGTH_PROPS:
        return FixedLengthProp(data)
    else:
        if temp not in constants.VARIABLE_LENGTH_PROPS:
            # DEBUG.
            logger.warning(f'Unknown property type: {temp:04X}')
        return VariableLengthProp(data)


class PropBase(abc.ABC):
    """
    Base class for Prop instances.
    """

    def __init__(self, data: bytes):
        self.__name = data[3::-1].hex().upper()
        self.__type, self.__pID, flags = constants.st.ST_PROP_BASE.unpack(data[:8])
        self.__flags = PropertyFlags(flags)

    def __bytes__(self) -> bytes:
        return self.toBytes()

    @abc.abstractmethod
    def toBytes(self) -> bytes:
        """
        Converts the property into a string of 16 bytes.
        """

    @property
    def flags(self) -> PropertyFlags:
        """
        Integer that contains property flags.
        """
        return self.__flags

    @flags.setter
    def flags(self, value: PropertyFlags):
        if not isinstance(value, PropertyFlags):
            raise TypeError(':property flags: MUST be an instance of PropertyFlags.')

        self.__flags = value

    @property
    def name(self) -> str:
        """
        Hexadecimal representation of the property ID followed by the type.
        """
        return self.__name

    @property
    def propertyID(self) -> int:
        """
        The property ID for this property.
        """
        return self.__pID

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

    def __init__(self, data: bytes):
        super().__init__(data)
        self.__value = self._parseType(self.type, data[8:], data)

    def _parseType(self, _type: int, stream: bytes, raw: bytes) -> Any:
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
            value = constants.st.ST_LE_UI16.unpack(value[:2])[0]
        elif _type == 0x0003: # PtypInteger32
            value = constants.st.ST_LE_UI32.unpack(value[:4])[0]
        elif _type == 0x0004: # PtypFloating32
            value = constants.st.ST_LE_F32.unpack(value[:4])[0]
        elif _type == 0x0005: # PtypFloating64
            value = constants.st.ST_LE_F64.unpack(value)[0]
        elif _type == 0x0006: # PtypCurrency
            value = decimal.Decimal((constants.st.ST_LE_I64.unpack(value))[0]) / 10000
        elif _type == 0x0007: # PtypFloatingTime
            value = constants.st.ST_LE_F64.unpack(value)[0]
            return constants.PYTPFLOATINGTIME_START + datetime.timedelta(days = value)
        elif _type == 0x000A: # PtypErrorCode
            value = constants.st.ST_LE_UI32.unpack(value[:4])[0]
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
            value = constants.st.ST_LE_UI16.unpack(value[:2])[0] != 0
        elif _type == 0x0014:  # PtypInteger64
            value = constants.st.ST_LE_UI64.unpack(value)[0]
        elif _type == 0x0040:  # PtypTime
            rawTime = constants.st.ST_LE_UI64.unpack(value)[0]
            try:
                value = filetimeToDatetime(rawTime)
            except ValueError:
                logger.exception(raw)
        return value

    def toBytes(self) -> bytes:
        """
        Converts the property into bytes.

        :raises ValueError: An issue occured where the value was not converted
            to bytes.
        """
        # First convert the value back to bytes in some way.
        value = self.value
        if self.type == 0x0001:
            value = b'\x00\x00\x00\x00\x00\x00\x00\x00'
        elif self.type == 0x0002:
            value = constants.st.ST_LE_UI16.pack(value) + b'\x00' * 6
        elif self.type == 0x0003 or self.type == 0x000A:
            value = constants.st.ST_LE_UI32.pack(value) + b'\x00' * 4
        elif self.type == 0x0004:
            value = constants.st.ST_LE_F32.pack(value) + b'\x00' * 4
        elif self.type == 0x0005:
            value = constants.st.ST_LE_F64.pack(value)
        elif self.type == 0x0006:
            value = constants.st.ST_LE_I64.pack(int(value * 10000))
        elif self.type == 0x0007:
            value = constants.st.ST_LE_F64.pack(
                (value - constants.PYTPFLOATINGTIME_START).total_seconds() / 86400
            )
        elif self.type == 0x000B:
            value = (b'\x01' + b'\x00' * 7) if value else (b'\x00' * 8)
        elif self.type == 0x0014:
            value = constants.st.ST_LE_UI64.pack(value)
        elif self.type == 0x0040:
            if hasattr(value, 'filetime') and value.filetime is not None:
                value = value.filetime
            elif isinstance(value, datetime.datetime):
                try:
                    value = value.timestamp()
                except OSError:
                    # Can't convert to a timestamp, so try to convert manually.
                    value = decimal.Decimal((value - datetime.datetime(1601, 1, 1)).total_seconds())
                else:
                    value = decimal.Decimal(value)
                    value += 11644473600

                # Here we now have a Decimal value. We want to convert it to an
                # int representing the 100 nanosecond intervals since the 0
                # date.
                value = int(value * 10000000)

            if isinstance(value, int):
                value = constants.st.ST_LE_UI64.pack(value)

        if not isinstance(value, bytes):
            raise ValueError(f'Failed to convert value to bytes (expected bytes at end, got {type(value)}). Please report to developer.')

        return constants.st.ST_PROP_BASE.pack(self.type, self.propertyID, self.flags) + value

    @property
    def signedValue(self) -> Any:
        """
        A signed representation of the value.

        Setting the value through this property will convert it if necessary
        before using the default value setter.

        :raises struct.error: The value was out of range when setting.
        """
        if self.type == 0x0002:
            return constants.st.ST_SBO_I16.unpack(constants.st.ST_SBO_UI16.pack(self.value))[0]
        if self.type == 0x0003:
            return constants.st.ST_SBO_I32.unpack(constants.st.ST_SBO_UI32.pack(self.value))[0]
        if self.type == 0x0014:
            return constants.st.ST_SBO_I64.unpack(constants.st.ST_SBO_UI64.pack(self.value))[0]

        # If not any of those types, return without modification.
        return self.value

    @signedValue.setter
    def signedValue(self, value: Any) -> None:
        if self.type == 0x0002:
            value = constants.st.ST_SBO_UI16.unpack(constants.st.ST_SBO_I16.pack(value))[0]
        if self.type == 0x0003:
            value = constants.st.ST_SBO_UI32.unpack(constants.st.ST_SBO_I32.pack(value))[0]
        if self.type == 0x0014:
            value = constants.st.ST_SBO_UI64.unpack(constants.st.ST_SBO_I64.pack(value))[0]

        self.value = value

    @property
    def value(self) -> Any:
        """
        Property value.
        """
        return self.__value

    @value.setter
    def value(self, value: Any) -> None:
        # Validate the value and perform necessary conversions.
        if self.type == 0x0000: # Unspecified.
            if not isinstance(value, bytes):
                raise TypeError(':property value: MUST be bytes when type is 0x0000.')
            if len(value) != 8:
                raise ValueError(':property value: MUST be 8 bytes when type is 0x0000.')
        elif self.type == 0x0001: # Null.
            raise TypeError(':property value: cannot be set when type is 0x0001.')
        elif self.type in (0x0002, 0x0003, 0x0014): # Ints.
            if not isinstance(value, int):
                raise TypeError(f':property value: MUST be an int when type is 0x{self.type:04X}')
            if value < 0:
                raise ValueError(f':property value: MUST be positive when type is 0x{self.type:04X}.')
            if self.type == 0x0002:
                if value > 0xFFFF:
                    raise ValueError(':property value: MUST be less than 0x10000 when type is 0x0002.')
            elif self.type == 0x0003:
                if value > 0xFFFFFFFF:
                    raise ValueError(':property value: MUST be less than 0x100000000 when type is 0x0003.')
            elif self.type == 0x0014:
                if value > 0xFFFFFFFFFFFFFFFF:
                    raise ValueError(':property value: MUST be less than 0x10000000000000000 when type is 0x0014.')
        elif self.type == 0x0004 or self.type == 0x0005: # Float/Double.
            if not isinstance(value, float):
                try:
                    value = float(value)
                except Exception:
                    raise TypeError(f':property value: MUST be float or convertable to float when type is 0x{self.type:04X}.')
        elif self.type == 0x0006:
            if not isinstance(value, decimal.Decimal):
                try:
                    value = decimal.Decimal(value)
                except decimal.InvalidOperation:
                    raise ValueError(':property value: MUST be Decimal or convertable to Decimal when type is 0x0006.')
                except TypeError:
                    raise TypeError(':property value: MUST be Decimal or convertable to Decimal when type is 0x0006.')
            if (value * 10000) > 0xFFFFFFFFFFFFFFFF:
                raise ValueError('Decimal value is too large (must be less than 0x10000000000000000 when multiplied by 10000).')
        elif self.type == 0x0007 or self.type == 0x0040:
            if not isinstance(value, datetime.datetime):
                raise TypeError(f':property value: MUST be an instance of datetime when type is 0x{self.type:04X}.')
        elif self.type == 0x000A:
            if not isinstance(value, int):
                raise TypeError(':property value: MUST be an int or instance of ErrorCodeType when type is 0x000A.')
            if not isinstance(value, ErrorCodeType):
                if value > 0xFFFFFFFF:
                    raise ValueError(':property value: MUST be less than 0x100000000 when type is 0x000A and the value is not an instance of ErrorCodeType.')
                # Convert only if possible.
                try:
                    value = ErrorCodeType(value)
                except ValueError:
                    # Ignore it if the conversion isn't possible.
                    pass
        elif self.type == 0x000B:
            if not isinstance(value, bool):
                raise TypeError(f':property value: MUST be bool when type is 0x{self.type:04X}.')

        self.__value = value



class VariableLengthProp(PropBase):
    """
    Class to contain the data for a single variable length property.
    """

    def __init__(self, data: bytes):
        super().__init__(data)
        self.__length, self.__reserved = constants.st.ST_PROP_VAR.unpack(data[8:])
        if self.type == 0x001E:
            # Check that the value is valid.
            if self.__length == 0:
                logger.warning('Property of type 0x001E found with length that was not at least 1. This will be corrected automatically.')
                self.__length = 1
            self.__realLength = self.__length - 1
        elif self.type == 0x001F:
            if self.__length == 0:
                logger.warning('Property of type 0x001F found with length that was not at least 2. This will be corrected automatically.')
                self.__length = 2
            if self.__length & 1:
                logger.warning('Property of type 0x001F found with length that is not a multiple of 2. This will not be corrected but is likely an error. This may cause issues with reading this property in other programs.')
            self.__realLength = self.__length // 2 - 1
        elif self.type in constants.MULTIPLE_2_BYTES_HEX:
            if self.__length & 1:
                logger.warning(f'Property of type {self.type} found with length that is not a multiple of 2. This will not be corrected but is likely an error. This may cause issues with reading this property in other programs.')
            self.__realLength = self.__length // 2
        elif self.type in constants.MULTIPLE_4_BYTES_HEX:
            if self.__length & 3:
                logger.warning(f'Property of type {self.type} found with length that is not a multiple of 4. This will not be corrected but is likely an error. This may cause issues with reading this property in other programs.')
            self.__realLength = self.__length // 4
        elif self.type in constants.MULTIPLE_8_BYTES_HEX:
            if self.__length & 7:
                logger.warning(f'Property of type {self.type} found with length that is not a multiple of 8. This will not be corrected but is likely an error. This may cause issues with reading this property in other programs.')
            self.__realLength = self.__length // 8
        elif self.type in constants.MULTIPLE_16_BYTES_HEX:
            if self.__length & 15:
                logger.warning(f'Property of type {self.type} found with length that is not a multiple of 16. This will not be corrected but is likely an error. This may cause issues with reading this property in other programs.')
            self.__realLength = self.__length // 16
        elif self.type == 0x000D:
            self.__realLength = 0xFFFFFFFF
            # Check the value and log a warning if it is bad before fixing it.
            if self.__length != 0xFFFFFFFF:
                logger.warning(f'Property of type 0x000D found with length that was not 0xFFFFFFFF (got {self.__length:08X}). This will be corrected automatically.')
                self.__length = 0xFFFFFFFF
        elif self.type == 0x0048:
            self.__realLength = 16
            # Check the value and log a warning if it is bad before fixing it.
            if self.__length != 16:
                logger.warning(f'Property of type 0x0048 found with length that was not 16 (got {self.__length}). This will be corrected automatically.')
                self.__length = 16
        elif self.type == 0x101E or self.type == 0x101F:
            if self.__length & 3:
                logger.warning(f'Property of type {self.type} found with length that is not a multiple of 4. This will not be corrected but is likely an error. This may cause issues with reading this property in other programs.')
            self.__realLength = self.__length // 4
        elif self.type == 0x1102:
            if self.__length & 7:
                logger.warning(f'Property of type {self.type} found with length that is not a multiple of 8. This will not be corrected but is likely an error. This may cause issues with reading this property in other programs.')
            self.__realLength = self.__length // 8
        else:
            if self.type in (0x00FB, 0x00FF, 0x00FE):
                # None of these are properly implemented, but we don't want to
                # actually raise an exception because of them. So just log the
                # issue.
                logger.error(f'Property type {self.type} has no documentation in [MS-OXMSG] but was found on this file. Please report this to the developer.')
            self.__realLength = self.__length

    def toBytes(self) -> bytes:
        ret = constants.st.ST_PROP_BASE.pack(self.type, self.propertyID, self.flags)
        ret += constants.st.ST_PROP_VAR.pack(self.__length, self.__reserved)
        return ret

    @property
    def size(self) -> int:
        """
        The size of the data the property corresponds to.

        For string streams, this is the number of characters contained. For
        multiple properties, this is the number of entries.

        When setting this, the underlying length field will be set which is a
        manipulated value. For multiple properties of a fixed length, this will
        be the size value multiplied by the length of the properties. For
        multiple strings, this will be 4 times the size value. For multiple
        binary, this will be 8 times the size value. For strings, this will be
        the number of characters plus 1 if non-unicode, otherwise the number of
        characters plus 1, all multiplied by 2. For binary, this will be the
        size with no modification.

        Size cannot be set for properties of type 0x000D and 0x0048.

        :raises TypeError: Tried to set the size for a property that cannot have
            the size changed.
        :raises ValueError: The translated value for the size is too large when
            setting.
        """
        return self.__realLength

    @size.setter
    def size(self, value: int) -> None:
        if not isinstance(value, int):
            raise TypeError(':property size: MUST be an int.')
        if value < 0:
            raise ValueError(':property size: MUST be positive.')
        if self.type == 0x000D or self.type == 0x0048:
            raise TypeError(f':property size: cannot be set for 0x{self.type:04X}.')

        # Convert the size to the actual length value.
        if self.type == 0x001E:
            length = value + 1
        elif self.type == 0x001F:
            length = (value + 1) * 2
        elif self.type in constants.MULTIPLE_2_BYTES_HEX:
            length = value * 2
        elif self.type in constants.MULTIPLE_4_BYTES_HEX:
            length = value * 4
        elif self.type in constants.MULTIPLE_8_BYTES_HEX:
            length = value * 8
        elif self.type in constants.MULTIPLE_16_BYTES_HEX:
            length = value * 16
        elif self.type == 0x101E or self.type == 0x101F:
            length = value * 4
        elif self.type == 0x1102:
            length = value * 8

        # Validate the range of the length value.
        if length > 0xFFFFFFFF:
            raise ValueError(f'Calculated length value is too large (max is 0xFFFFFFFF, got {length:08X}).')

        self.__length = length
        self.__realLength = value

    @property
    def reservedFlags(self) -> int:
        """
        The reserved flags field of the variable length property.
        """
        return self.__reserved

    @reservedFlags.setter
    def reservedFlags(self, value: int) -> None:
        if not isinstance(value, int):
            raise TypeError(':property reservedFlags: MUST be an int.')
        if value < 0:
            raise ValueError(':property reservedFlags: MUST be positive.')
        if value > 0xFFFFFFFF:
            raise ValueError(':property reservedFlags: MUST be less than 0x100000000.')

        self.__reserved = value
