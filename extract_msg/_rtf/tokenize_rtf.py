__all__ = [
    'tokenizeRTF',
]


import io

from typing import List, Optional, Tuple

from .token import Token, TokenType


_KNOWN_DESTINATIONS = (
    b'aftncn',
    b'aftnsep',
    b'aftnsepc',
    b'annotation',
    b'author',
    b'buptim',
    b'category',
    b'colortbl',
    b'comment',
    b'company',
    b'creatim',
    b'doccomm',
    b'dptxbxtext',
    b'factoidname',
    b'fonttbl',
    b'footer',
    b'footerf',
    b'footerl',
    b'footerr',
    b'ftncn',
    b'ftnsep',
    b'ftnsepc',
    b'header',
    b'headerf',
    b'headerl',
    b'headerr',
    b'hlinkbase',
    b'keywords',
    b'manager',
    b'operator',
    b'pict',
    b'printim',
    b'private',
    b'revtim',
    b'stylesheet',
    b'subject',
    b'title',
)


def _finishTag(startText: bytes, reader: io.BytesIO) -> Tuple[bytes, Optional[bytes], Optional[int], bytes]:
    """
    Finishes reading a tag, returning the needed parameters to make it a
    token.

    The return is a 4 tuple of the raw token bytes, the name field, the
    parameter field (as an int), and the next character after the tag.
    """
    # Very simple rules here. Anything other than a letter and we change
    # state. If the next character is a hypen, check if the character after
    # is a digit, otherwise return. If it is a digit or that previously
    # mentioned next character was a digit, read digits until anything else
    # is detected, then return.
    name = startText[-1:]
    param = b''

    while (nextChar := reader.read(1)) != b'' and nextChar.isalpha():
        # Read until not alpha.
        startText += nextChar
        name += nextChar

    # Check what the next character is to decide what to do with it.
    if nextChar == b'-':
        # We do this as a separate check.
        nextNext = reader.read(1)
        if nextNext == b'':
            raise ValueError('Unexpected end of data.')
        elif nextNext.isdigit():
            startText += nextChar
            nextChar = nextNext

    if nextChar.isdigit():
        startText += nextChar
        param += nextChar
        while (nextChar := reader.read(1)) != b'' and nextChar.isdigit():
            startText += nextChar
            param += nextChar

        param = int(param)
    else:
        param = None

    # Finally, check if the next char is a space, and if it is, read one
    # more char to replace it.
    if nextChar == b' ':
        nextChar = reader.read(1)

    return startText, name, param, nextChar


def _readControl(startChar: bytes, reader: io.BytesIO) -> Tuple[Tuple[Token, ...], bytes]:
    """
    Attempts to read the next data as a control, returning as many tokens
    as necessary.
    """
    # First, read the next character, as it decides how to handle
    # everything.
    nextChar = reader.read(1)
    if nextChar == b'':
        raise ValueError('Unexpected end of data.')
    elif nextChar.isalpha():
        # If is an alphabetical character, so start the handling of a tag.
        text, name, param, nextChar = _finishTag(startChar + nextChar, reader)
        # Important, check if the name is "bin". If it is, handle that
        # specially before returning.
        if name == b'bin':
            if nextChar == b'':
                raise ValueError('Unexpected end of data.')
            binText = nextChar + reader.read(param - 1)
            if len(binText) != param:
                raise ValueError('Unexpected end of data.')
            return (Token(text, TokenType.CONTROL, name, param), Token(binText, TokenType.BINARY)), nextChar
        elif name in _KNOWN_DESTINATIONS:
            return (Token(text, TokenType.DESTINATION, name, param),), nextChar

        return (Token(text, TokenType.CONTROL, name, param),), nextChar
    else:
        # Most control symbols would return immediately, but there are two
        # exceptions.
        startChar += nextChar
        if nextChar == b'*':
            # This is going to be a custom destination. First, validation.
            if len(nextChar := reader.read(1)) != 1:
                raise ValueError('Unexpected end of data.')
            elif nextChar != b'\\':
                raise ValueError(f'Bad custom destination (expected a backslash, got {nextChar}).')

            startChar += nextChar

            # Check the the next char is alpha.
            if not (nextChar := reader.read(1)).isalpha():
                raise ValueError(f'Expected alpha character for destination, got {nextChar}.')

            startChar += nextChar

            # Call the function to read until a clear end of tag.
            text, name, param, nextChar = _finishTag(startChar, reader)
            return (Token(text, TokenType.IGNORABLE_DESTINATION, name, param),), nextChar
        elif nextChar == b'\'':
            # This is a hex character, so immediately read 2 more bytes.
            hexChars = reader.read(2)
            if len(hexChars) != 2:
                raise ValueError('Unexpected end of data.')
            try:
                param = int(hexChars, 16)
            except ValueError as e:
                context = e.__cause__ or e.__context__
                raise ValueError(f'Hex data was not hexidecimal (got {hexChars}).') from context
            return (Token(startChar + hexChars, TokenType.SYMBOL, None, param),), reader.read(1)
        else:
            # If it is a control symbol, immediately return.
            return (Token(startChar, TokenType.SYMBOL),), reader.read(1)


def _readText(startChar: bytes, reader: io.BytesIO) -> Tuple[Tuple[Token, ...], bytes]:
    """
    Attempts to read the next data as text.
    """
    chars = [startChar]
    # Text is actually the easiest to read, as we just read until end of
    # stream or until a special character. However, a few characters are
    # simply dropped during reading.
    while (nextChar := reader.read(1)) != b'' and nextChar not in (b'{', b'}', b'\\'):
        # Certain characters are simply dropped.
        if nextChar not in (b'\r', b'\n'):
            chars.append(nextChar)

    # Now, we actually are reading the text as *individual tokens*, so we
    # need to

    return tuple(Token(x, TokenType.TEXT) for x in chars), nextChar


def tokenizeRTF(data: bytes, validateStart: bool = True) -> List[Token]:
    """
    Reads in the bytes and sets the tokens list to the contents after
    tokenizing.

    If tokenizing fails, the current tokens list will not be changed.

    :param validateStart: If ``False``, does not check the first few tags.
        Useful when tokenizing a snippet rather than a document.

    :raises TypeError: The data is not recognized as RTF.
    :raises ValueError: An issue with basic parsing occured.
    """
    reader = io.BytesIO(data)
    if validateStart:
        # This tokenizer *only* breaks things up. It does *not* care about
        # groups and stuff, as that is for a parser to deal with. All we do is
        # track the current backslash state and token state. We also simply
        # check that the first token is "\rtf1" preceeded by a group start, and
        # that is it.
        start = reader.read(6)
        if start != b'{\\rtf1':
            raise TypeError('Data does not start with "{\\rtf1".')

        tokens = [
            Token(b'{', TokenType.GROUP_START),
            Token(b'\\rtf1', TokenType.CONTROL, b'rtf', 1),
        ]
        nextChar = reader.read(1)

        # If the next character is a space, ignore it.
        if nextChar == b' ':
            nextChar = reader.read(1)
    else:
        tokens = []
        nextChar = reader.read(1)

    # At every iteration, so long as there is more data, nextChar should be
    # set. As such, use it to determine what kind of data to try to read,
    # using the delimeter of that type of data to know what to do next.
    while nextChar != b'':
        # We should have exactly one character, the start of the next
        # section. Use it to determine what to do.
        if nextChar in (b'\r', b'\n'):
            # Just read the next character and start the loop over.
            nextChar = reader.read(1)
            continue

        if nextChar == b'\\':
            newTokens, nextChar = _readControl(nextChar, reader)
        elif nextChar == b'{':
            # This will always be a group start, which has nothing left to
            # read.
            nextChar = reader.read(1)
            newTokens = (Token(b'{', TokenType.GROUP_START),)
        elif nextChar == b'}':
            # This will always be a group end, which has nothing left to
            # read.
            nextChar = reader.read(1)
            newTokens = (Token(b'}', TokenType.GROUP_END),)
        else:
            # Otherwise, it's just text.
            newTokens, nextChar = _readText(nextChar, reader)

        tokens.extend(newTokens)

    return tokens
