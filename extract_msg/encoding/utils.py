"""
Internal utilities for extract_msg.encoding.
"""

__all__ = [
    'variableByteDecode',
    'variableByteEncode',
]


import codecs

from typing import Dict, Tuple


def createVBEncoding(codecName: str, decodingTable: Dict[int, str]) -> codecs.CodecInfo:
    """
    Creates the classes for a variable byte encoding, returning the CodecInfo
    instance associated with it.

    Currently only supports encodings with up to 2 bytes per character.

    :param codecName: The name of the codec being created.
    :param decodingTable: The table to use for decoding data. Will be used to
        create the encodingTable.
    """
    # Create the encoding table. If there are duplicate keys, they will use the
    # last one introduced. Iterating in reverse means the lowest possible value
    # for a character will be the one used.
    encodingTable = {
        value: bytes((key,)) if key < 256 else bytes((key >> 8, key & 0xFF))
        for key, value in reversed(decodingTable.items()) if value is not None
    }

    # Create the classes.
    class Codec(codecs.Codec):
        def encode(self, text, errors='strict'):
            return variableByteEncode(codecName, text, errors, encodingTable)

        def decode(self, data, errors='strict'):
            return variableByteDecode(codecName, data, errors, decodingTable)

    class IncrementalEncoder(codecs.IncrementalEncoder):
        def encode(self, text, final=False):
            return variableByteEncode(codecName, text, self.errors, encodingTable)[0]

    class IncrementalDecoder(codecs.IncrementalDecoder):
        def decode(self, data, final=False):
            return variableByteDecode(codecName, data, self.errors, decodingTable)[0]

    class StreamWriter(Codec, codecs.StreamWriter):
        pass

    class StreamReader(Codec, codecs.StreamReader):
        pass

    # Return the CodecInfo instance.
    return codecs.CodecInfo(
        name=codecName,
        encode=Codec().encode,
        decode=Codec().decode,
        incrementalencoder=IncrementalEncoder,
        incrementaldecoder=IncrementalDecoder,
        streamwriter=StreamWriter,
        streamreader=StreamReader,
    )


def createSBEncoding(codecName: str, decodingTable: Dict[int, str]) -> codecs.CodecInfo:
    """
    Creates the classes for a single byte encoding, returning the CodecInfo
    instance associated with it.

    Currently only supports encodings with up to 2 bytes per character.

    :param codecName: The name of the codec being created.
    :param decodingTable: The table to use for decoding data. Will be used to
        create the encodingTable.
    """
    # Create the encoding table. If there are duplicate keys, they will use the
    # last one introduced. Iterating in reverse means the lowest possible value
    # for a character will be the one used.
    encodingTable = {
        value: bytes((key,))
        for key, value in reversed(decodingTable.items()) if value is not None
    }

    # Create the classes.
    class Codec(codecs.Codec):
        def encode(self, text, errors='strict'):
            return singleByteEncode(codecName, text, errors, encodingTable)

        def decode(self, data, errors='strict'):
            return singleByteDecode(codecName, data, errors, decodingTable)

    class IncrementalEncoder(codecs.IncrementalEncoder):
        def encode(self, text, final=False):
            return singleByteEncode(codecName, text, self.errors, encodingTable)[0]

    class IncrementalDecoder(codecs.IncrementalDecoder):
        def decode(self, data, final=False):
            return singleByteDecode(codecName, data, self.errors, decodingTable)[0]

    class StreamWriter(Codec, codecs.StreamWriter):
        pass

    class StreamReader(Codec, codecs.StreamReader):
        pass

    # Return the CodecInfo instance.
    return codecs.CodecInfo(
        name=codecName,
        encode=Codec().encode,
        decode=Codec().decode,
        incrementalencoder=IncrementalEncoder,
        incrementaldecoder=IncrementalDecoder,
        streamwriter=StreamWriter,
        streamreader=StreamReader,
    )


def singleByteDecode(codecName: str, data, errors: str, decodeTable: Dict[int, str]) -> Tuple[str, int]:
    """
    Function for decoding single-byte codecs.

    :param codecName: The name of the codec, used for error messages.
    :param data: A bytes-like object to decode.
    :param errors: The error behavior to use.
    :param decodeTable: The mapping of values to use. Continuation bytes MUST be
        defined in the table, but SHOULD be set to None. This allows for the
        function to detect what bytes are valid for continuation.
    """
    if len(data) == 0:
        return ('', 0)

    errorHandler = codecs.lookup_error(errors)
    output = ''

    iterator = enumerate(data)
    start = 0

    for start, byte in iterator:
        if byte in decodeTable:
            output += decodeTable[byte]
        else:
            err = UnicodeDecodeError(codecName,
                                     data,
                                     start,
                                     start + 1,
                                     'character maps to <undefined>',
                                     )
            rep = errorHandler(err)
            output += rep[0]
            for _ in range(rep[1] - start - 1):
                try:
                    next(iterator)
                except StopIteration:
                    break

    return (output, len(output))


def singleByteEncode(codecName: str, data, errors: str, encodeTable: Dict[str, bytes]) -> Tuple[bytes, int]:
    """
    Function for encoding variable-byte codecs that use one or two bytes per
    character.

    :param codecName: The name of the codec, used for error messages.
    :param data: A bytes-like object to decode.
    :param errors: The error behavior to use.
    :param encodeTable: The mapping of values to use.
    """
    if len(data) == 0:
        return

    errorHandler = codecs.lookup_error(errors)
    output = b''
    iterator = enumerate(data)
    start = 0
    for start, char in iterator:
        if char in encodeTable:
            output += encodeTable[char]
        else:
            err = UnicodeEncodeError(codecName,
                                     data,
                                     start,
                                     start + 1,
                                     'illegal multibyte sequence',
                                     )
            rep = errorHandler(err)
            output += rep[0]
            # Skip the specified number of characters.
            for _ in range(rep[1] - start - 1):
                try:
                    next(iterator)
                except StopIteration:
                    break



def variableByteDecode(codecName: str, data, errors: str, decodeTable: Dict[int, str]) -> Tuple[str, int]:
    """
    Function for decoding variable-byte codecs that use one or two bytes per character.

    Checks if a character is less than ``0x80``, mapping it directly if so.
    Otherwise, it reads the next byte and combines the two before looking up the
    new value.

    :param codecName: The name of the codec, used for error messages.
    :param data: A bytes-like object to decode.
    :param errors: The error behavior to use.
    :param decodeTable: The mapping of values to use. Continuation bytes MUST be
        defined in the table, but SHOULD be set to None. This allows for the
        function to detect what bytes are valid for continuation.
    """
    if len(data) == 0:
        return ('', 0)

    errorHandler = codecs.lookup_error(errors)
    output = ''

    iterator = enumerate(data)
    start = 0

    for start, byte in iterator:
        # Variable byte should be an integer here.
        if byte < 0x80:
            if byte in decodeTable:
                output += decodeTable[byte]
            else:
                err = UnicodeDecodeError(codecName,
                                         data,
                                         start,
                                         start + 1,
                                         'character maps to <undefined>',
                                        )
                rep = errorHandler(err)
                output += rep[0]
                # Skip the specified number of characters.
                for _ in range(rep[1] - start - 1):
                    try:
                        next(iterator)
                    except StopIteration:
                        break
        elif byte not in decodeTable:
            err = UnicodeDecodeError(codecName,
                                     data,
                                     start,
                                     start + 1,
                                     'invalid start byte',
                                    )
            rep = errorHandler(err)
            output += rep[0]
            # Skip the specified number of characters.
            for _ in range(rep[1] - start - 1):
                try:
                    next(iterator)
                except StopIteration:
                    break

        else:
            try:
                byte = (byte << 8) | next(iterator)[1]
                if byte in decodeTable:
                    output += decodeTable[byte]
                else:
                    err = UnicodeDecodeError(codecName,
                                             data,
                                             start,
                                             start + 2,
                                             'character maps to <undefined>',
                                             )
                    rep = errorHandler(err)
                    output += rep[0]
                    # Skip the specified number of characters.
                    for _ in range(rep[1] - start - 1):
                        next(iterator)
            except StopIteration:
                err = UnicodeDecodeError(codecName,
                                         data,
                                         start,
                                         start + 1,
                                         'unexpected end of data',
                                         )
                rep = errorHandler(err)
                output += rep[0]
                break
                # No more data, so that's all that needs to happen.
    return (output, start)


def variableByteEncode(codecName: str, data, errors: str, encodeTable: Dict[str, bytes]) -> Tuple[bytes, int]:
    """
    Function for encoding variable-byte codecs that use one or two bytes per
    character.

    :param codecName: The name of the codec, used for error messages.
    :param data: A bytes-like object to decode.
    :param errors: The error behavior to use.
    :param encodeTable: The mapping of values to use.
    """
    if len(data) == 0:
        return

    errorHandler = codecs.lookup_error(errors)
    output = b''
    iterator = enumerate(data)
    start = 0

    for start, char in iterator:
        if char in encodeTable:
            data += encodeTable[char]
        else:
            err = UnicodeEncodeError(codecName,
                                     data,
                                     start,
                                     start + 1,
                                     'illegal multibyte sequence',
                                     )
            rep = errorHandler(err)
            output += rep[0]
            # Skip the specified number of characters.
            for _ in range(rep[1] - start - 1):
                try:
                    next(iterator)
                except StopIteration:
                    break

    return (output, start + 1)
