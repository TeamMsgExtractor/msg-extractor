"""
Module for helper classes to different data structures.
"""

import io

from .. import constants


class BytesReader(io.BytesIO):
    """
    Extension of io.BytesIO that allows you to read specific data types from the
    stream.
    """
    def __init__(self, *args, littleEndian = True, **kwargs):
        super().__init__(*args, **kwargs)
        self.__le = bool(littleEndian)
        if self.__le:
            self.__int8_t = constants.ST_LE_I8
            self.__int16_t = constants.ST_LE_I16
            self.__int32_t = constants.ST_LE_I32
            self.__int64_t = constants.ST_LE_I64
            self.__uint8_t = constants.ST_LE_UI8
            self.__uint16_t = constants.ST_LE_UI16
            self.__uint32_t = constants.ST_LE_UI32
            self.__uint64_t = constants.ST_LE_UI64
            self.__float_t = constants.ST_LE_F32
            self.__double_t = constants.ST_LE_F64
        else:
            self.__int8_t = constants.ST_BE_I8
            self.__int16_t = constants.ST_BE_I16
            self.__int32_t = constants.ST_BE_I32
            self.__int64_t = constants.ST_BE_I64
            self.__uint8_t = constants.ST_BE_UI8
            self.__uint16_t = constants.ST_BE_UI16
            self.__uint32_t = constants.ST_BE_UI32
            self.__uint64_t = constants.ST_BE_UI64
            self.__float_t = constants.ST_BE_F32
            self.__double_t = constants.ST_BE_F64

    def _readDecodedString(self, encoding, width : int = 1):
        """
        Reads a null terminated string with the specified character width
        decoded using the specified encoding. If it cannot be read or cannot be
        decoded then the position of the read pointer will not be changed.
        """
        position = self.tell()
        try:
            return self.readByteString(width).decode(encoding)
        except:
            while self.tell() != position:
                self.seek(position)
            raise

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

    def readByteString(self, width : int = 1) -> bytes:
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
        endFound = False;
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

    def tryReadBytes(self, size : int) -> bytes:
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
