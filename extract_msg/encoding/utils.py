"""
Internal utilities for extract_msg.encoding.
"""

__all__ = [
    'variableByteDecode',
    'variableByteEncode',
]


import codecs

from typing import Dict, Tuple


def variableByteDecode(codecName : str, data, errors : str, decodeTable : Dict[int, str]) -> Tuple[str, int]:
    """
    Function for decoding variable-byte codecs.

    Checks if a character is less than 0x80, mapping it directly if so.
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
                                         'character maps to <undefined>'
                                        )
                rep = errorHandler(err)
                output += rep[0]
                # Skip the specified number of characters.
                for _ in range(rep[1] - start - 1):
                    iterator.__next__()
        elif byte not in decodeTable:
            err = UnicodeDecodeError(codecName,
                                     data,
                                     start,
                                     start + 1,
                                     'invalid start byte'
                                    )
            rep = errorHandler(err)
            output += rep[0]
            # Skip the specified number of characters.
            for _ in range(rep[1] - start - 1):
                iterator.__next__()

        else:
            try:
                byte = (byte << 8) | iterator.__next__()[1]
                if byte in decodeTable:
                    output += decodeTable[byte]
                else:
                    err = UnicodeDecodeError(codecName,
                                             data,
                                             start,
                                             start + 2,
                                             'character maps to <undefined>'
                                             )
                    rep = errorHandler(err)
                    output += rep[0]
                    # Skip the specified number of characters.
                    for _ in range(rep[1] - start - 1):
                        iterator.__next__()
            except StopIteration:
                err = UnicodeDecodeError(codecName,
                                         data,
                                         start,
                                         start + 1,
                                         'unexpected end of data'
                                         )
                rep = errorHandler(err)
                output += rep[0]
                break
                # No more data, so that's all that needs to happen.
    return (output, start)


def variableByteEncode(codecName : str, data, errors : str, encodeTable : Dict[str, int]) -> Tuple[bytes, int]:
    """
    Function for decoding variable-byte codecs.

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
    for start, char in iterator:
        if char not in encodeTable:
            err = UnicodeEncodeError(codecName,
                                     data,
                                     start,
                                     start + 1,
                                     'illegal multibyte sequence')
            rep = errorHandler(err)
            output += rep[0]
            # Skip the specified number of characters.
            for _ in range(rep[1] - start - 1):
                iterator.__next__()
        else:
            data += encodeTable[char]

    return output
