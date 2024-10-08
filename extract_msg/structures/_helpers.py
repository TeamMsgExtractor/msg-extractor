"""
Module for helper classes to different data structures.
"""

__all__ = [
    'BytesReader',
]


import io
import struct

from typing import Any, Optional, Tuple, Type, TypeVar, Union

from .. import constants

_T = TypeVar('_T')


class BytesReader(io.BytesIO):
    """
    Extension of io.BytesIO that allows you to read specific data types from the
    stream.
    """

    def __init__(self, *args, littleEndian: bool = True, **kwargs):
        super().__init__(*args, **kwargs)
        self.__le = bool(littleEndian)
        if self.__le:
            self.__int8_t = constants.st.ST_LE_I8
            self.__int16_t = constants.st.ST_LE_I16
            self.__int32_t = constants.st.ST_LE_I32
            self.__int64_t = constants.st.ST_LE_I64
            self.__uint8_t = constants.st.ST_LE_UI8
            self.__uint16_t = constants.st.ST_LE_UI16
            self.__uint32_t = constants.st.ST_LE_UI32
            self.__uint64_t = constants.st.ST_LE_UI64
            self.__float_t = constants.st.ST_LE_F32
            self.__double_t = constants.st.ST_LE_F64
        else:
            self.__int8_t = constants.st.ST_BE_I8
            self.__int16_t = constants.st.ST_BE_I16
            self.__int32_t = constants.st.ST_BE_I32
            self.__int64_t = constants.st.ST_BE_I64
            self.__uint8_t = constants.st.ST_BE_UI8
            self.__uint16_t = constants.st.ST_BE_UI16
            self.__uint32_t = constants.st.ST_BE_UI32
            self.__uint64_t = constants.st.ST_BE_UI64
            self.__float_t = constants.st.ST_BE_F32
            self.__double_t = constants.st.ST_BE_F64

    def _readDecodedString(self, encoding: str, width: int = 1) -> str:
        """
        Reads a null terminated string with the specified character width
        decoded using the specified encoding. If it cannot be read or cannot be
        decoded then the position of the read pointer will not be changed.
        """
        position = self.tell()
        try:
            return self.readByteString(width).decode(encoding)
        except Exception:
            while self.tell() != position:
                self.seek(position)
            raise

    def assertNull(self, length: int, errorMsg: Optional[str] = None) -> bytes:
        """
        Reads the number of bytes specified and ensures they are all null.

        Ensures the reader returns back to the spot before attempting to read if
        there are not enough bytes to read.

        :param length: The amount of bytes to read.
        :param errorMsg: Optional, the error message to use if the bytes are not
            all null.

        :returns: The bytes read, if you need them.

        :raise IOError: Not enough bytes left to read.
        :raises ValueError: Assertion failed.
        """
        # Quick return for reading 0 bytes.
        if length == 0:
            return b''

        valueRead = self.tryReadBytes(length)
        if valueRead:
            if sum(valueRead) != 0:
                errorMsg = errorMsg or 'Bytes read were not all null.'
                raise ValueError(errorMsg)
        else:
            raise IOError('Not enough bytes left in buffer.')

        return valueRead

    def assertRead(self, value: bytes, errorMsg: Optional[str] = None) -> bytes:
        """
        Reads the number of bytes and compares them to the value provided. If it
        does not match, throws a value error.

        Ensures the reader returns back to the spot before attempting to read if
        there are not enough bytes to read.

        :param value: Value to compare read bytes to.
        :param errorMsg: Optional, an error message to emit on mismatch. Does
            not apply to the buffer being too small. Allows for a format string
            with the keyword values "expected" and "actual", representing the
            value given to the function and the actual value read, respectively.

        :returns: The bytes read, if you need them.

        :raises TypeError: The value given was not bytes.
        :raises ValueError: Assertion failed.
        """
        # Quick return for a value being empty.
        if len(value) == 0:
            return b''

        if not isinstance(value, bytes): # pyright: ignore
            raise TypeError(':param value: was not bytes.')

        valueRead = self.tryReadBytes(len(value))
        if valueRead:
            if valueRead != value:
                errorMsg = errorMsg or 'Value did not match (expected {expected}, got {actual}).'
                raise ValueError(errorMsg.format(expected = value, actual = valueRead))
        else:
            raise IOError('Not enough bytes left in buffer.')

        return valueRead

    def readAnsiString(self) -> str:
        """
        Reads a null-terminated string in ANSI format.
        """
        return self._readDecodedString('ansi')

    def readAsciiString(self) -> str:
        """
        Reads a null-terminated string in ASCII format.
        """
        return self._readDecodedString('ascii')

    def readByte(self) -> int:
        """
        Reads a signed byte from the stream.
        """
        value = self.tryReadBytes(1)
        if value:
            return self.__int8_t.unpack(value)[0]
        else:
            raise IOError('Not enough bytes left in buffer.')

    def readByteString(self, width: int = 1) -> bytes:
        """
        Reads a string of bytes until it finds the null character, returning
        everything before that and consuming the null. Unlike other string
        functions, this will not decode the data into a string.

        :param width: tells how big a character is (in bytes), so this function
        can be used for strings whose characters are multiple bytes.
        """
        if width < 1:
            raise ValueError('Character width must be at least 1.')

        position = self.tell()
        string = b''
        null = b'\x00' * width

        while True:
            nextChar = self.read(width)
            if nextChar == b'':
                # We reached the end of the buffer without finding the null. We
                # need to seek back to where we started and raise an exception.
                while self.tell() != position:
                    self.seek(position)
                raise IOError('Could not find null character.')
            elif nextChar == null:
                # If we find the null character, return what we have read.
                return string
            else:
                # Otherwise add the character to our string.
                string += nextChar

    def readClass(self, _class: Type[_T]) -> _T:
        """
        Takes anything with a __SIZE__ property and a call function that takes
        a single bytes argument and returns the result of that function.

        Generally, this is intended to take a fixed-size class and return an
        instance of the class created with that amount of bytes. However, there
        is little reason to truly limit it to only that.
        """
        if not hasattr(_class, '__SIZE__'):
            raise TypeError('Argument to readClass MUST have a __SIZE__ attribute.')
        value = self.tryReadBytes(_class.__SIZE__)
        if value:
            return _class(value)
        else:
            raise IOError('Not enough bytes left in buffer.')

    def readDouble(self) -> float:
        """
        Reads a double from the stream.
        """
        value = self.tryReadBytes(8)
        if value:
            return self.__double_t.unpack(value)[0]
        else:
            raise IOError('Not enough bytes left in buffer.')

    def readFloat(self) -> float:
        """
        Reads a float from the stream.
        """
        value = self.tryReadBytes(4)
        if value:
            return self.__float_t.unpack(value)[0]
        else:
            raise IOError('Not enough bytes left in buffer.')

    def readInt(self) -> int:
        """
        Reads a signed int from the stream.
        """
        value = self.tryReadBytes(4)
        if value:
            return self.__int32_t.unpack(value)[0]
        else:
            raise IOError('Not enough bytes left in buffer.')

    def readLong(self) -> int:
        """
        Reads a signed byte from the stream.
        """
        value = self.tryReadBytes(8)
        if value:
            return self.__int64_t.unpack(value)[0]
        else:
            raise IOError('Not enough bytes left in buffer.')

    def readShort(self) -> int:
        """
        Reads a signed short from the stream.
        """
        value = self.tryReadBytes(2)
        if value:
            return self.__int16_t.unpack(value)[0]
        else:
            raise IOError('Not enough bytes left in buffer.')

    def readStruct(self, _struct: Union[struct.Struct, Any]) -> Tuple[Any, ...]:
        """
        Read enough bytes for a struct and unpack it, returning the tuple of
        values.

        :param _struct: A struct or struct-like object using duck-typing. Only
            requires the object have an unpack method that takes a single
            argument and a size property to tell how many bytes to read.

        :raises IOError: If there are not enough bytes left to read.
        """
        value = self.tryReadBytes(_struct.size)
        if value:
            return _struct.unpack(value)
        else:
            raise IOError('Not enough bytes left in buffer.')

    def readUnsignedByte(self) -> int:
        """
        Reads an unsigned byte from the stream.
        """
        value = self.tryReadBytes(1)
        if value:
            return self.__uint8_t.unpack(value)[0]
        else:
            raise IOError('Not enough bytes left in buffer.')

    def readUnsignedInt(self) -> int:
        """
        Reads an unsigned int from the stream.
        """
        value = self.tryReadBytes(4)
        if value:
            return self.__uint32_t.unpack(value)[0]
        else:
            raise IOError('Not enough bytes left in buffer.')

    def readUnsignedLong(self) -> int:
        """
        Reads an unsigned long from the stream.
        """
        value = self.tryReadBytes(8)
        if value:
            return self.__uint64_t.unpack(value)[0]
        else:
            raise IOError('Not enough bytes left in buffer.')

    def readUnsignedShort(self) -> int:
        """
        Reads an unsigned short from the stream.
        """
        value = self.tryReadBytes(2)
        if value:
            return self.__uint16_t.unpack(value)[0]
        else:
            raise IOError('Not enough bytes left in buffer.')

    def readUtf8String(self) -> str:
        """
        Reads a null-terminated string in UTF-8 format.
        """
        return self._readDecodedString('utf-8')

    def readUtf16String(self) -> str:
        """
        Reads a null terminated string in UTF-16 format using the endienness of
        the reader to determine which one to use.
        """
        return self._readDecodedString('utf-16-le' if self.__le else 'utf-16-be', 2)

    def readUtf32String(self) -> str:
        """
        Reads a null terminated string in UTF-32 format using the endienness of
        the reader to determine which one to use.
        """
        return self._readDecodedString('utf-32-le' if self.__le else 'utf-32-be', 4)

    def tryReadBytes(self, size: int) -> bytes:
        """
        Tries to read the specified number of bytes, returning b'' if not
        possible. Will only change the position of the read pointer if reading
        was possible.
        """
        if size < 1:
            raise ValueError(':param size: must be at least 1.')
        position = self.tell()
        value = self.read(size)
        if len(value) == size:
            return value

        # Ensure that we seek back to where we started if we could not read.
        while self.tell() != position:
            self.seek(position)
        return b''
