import copy
import enum
import io

from typing import List, NamedTuple, Tuple


class TokenType(enum.Enum):
    GROUP_START = 0
    GROUP_END = 1
    CONTROL = 2
    SYMBOL = 3
    TEXT = 4
    DESTINATION = 5
    IGNORABLE_DESTSINATION = 6
    # This one is special, used for handling the binary data.
    BINARY = 7



class Token(NamedTuple):
    # The raw bytes for the token, used to recreate the document.
    raw : bytes
    # The type of the token.
    type : TokenType
    ## The following are optional as they only apply for certain types of tokens.
    # The name of the token, if it is a control or destination.
    name : Optional[bytes] = None
    # The parameter of the token, if it has one. If the token is a `\'hh` token,
    # this will be the decimal equivelent of the hex value.
    parameter : Optional[int] = None
    # The symbol the token represents, if it is a symbol.
    symbol : Optional[str] = None



class RTFTokenizer:
    """
    Class designed to take in RTF bytes and split it into individual tokens with
    minimal validation of the contents.

    Tokens can be iterated by iteraating over the instance directly or can be
    accessed directly from the `tokens` instance variable.
    """

    def __init__(self, data : bytes = None):
        self.tokens : List[Token] = []
        # Feed the data to our parser if provided. Feeding will clear all of the
        # existing data.
        if data:
            self.feed(data)

    def __iter__(self):
        return self.tokens.__iter__()

    def __handleTag(self, tag : bytes, param : bytes) -> Token:
        """
        Handles converting an RTF tag into a Token.
        """
        if tag[1] == '*':
            # Handle custom destination. We also need to handle bad destination
            # data. TODO
            pass
        elif tag[1] in self.__knownSymbols:
            pass

        # TODO
        pass

    def __readControl(self, startChar : bytes, reader : io.BytesIO) -> Tuple[Tuple[Token], bytes]:
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

                pass


            elif nextChar == b'\'':
                # This is a hex character, so immediately read 2 more bytes.
                hexChars = reader.read(2)
                if len(hexChars) != 2:
                    raise ValueError('Unexpected end of data.')
                try:
                    param = int(hexChars, 16)
                except ValueError:
                    context = e.__cause__ or e.__context__
                    raise ValueError(f'Hex data was not hexidecimal (got {hexChars}).') from context
                return (self.__handleTag(startChar + hexChars, param),), reader.read(1)
            else:
                # If it is a control symbol, immediately return.
                return (self.__handleTag(startChar, b''),), reader.read(1)

        else:
            # If is an alphabetical character, so start the handling of a tag.
            pass

        # Handling \binN is going to be the hardest to do, so just give it
        # to it's entire own function.

    def __readText(self, startChar : bytes, reader : io.BytesIO) -> Tuple[Tuple[Token], bytes]:
        """
        Attempts to read the next data as text.
        """
        # Text is actually the easiest to read, as we just read until end of
        # stream or until a special character. However, a few characters are
        # simply dropped during reading.
        while (nextChar := reader.read(1)) not in (b'{', b'}', b'\\'):
            # Certain characters are simply dropped.
            if nextChar not in (b'\r', b'\n'):
                startChar += nextChar

        return (Token(startChar, TokenType.Text),), nextChar

    def feed(self, data : bytes) -> None:
        """
        Reads in the bytes and sets the tokens list to the contents after
        tokenizing. If tokenizing fails, the current tokens list will not be
        changed.

        Direct references to the previous tokens list will only point to the
        previous and not to the current one.

        :raises TypeError: The data is not recognized as RTF.
        :raises ValueError: An issue with basic parsing occured.
        """
        reader = io.BytesIO(data)
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
            Token(b'\rtf1', TokenType.CONTROL, b'rtf', 1),
        ]
        nextChar = reader.read(1)

        # If the next character is a space, ignore it.
        if nextChar == ' ':
            nextChar = reader.read()

        newToken = None

        # At every iteration, so long as there is more data, nextChar should be
        # set. As such, use it to determine what kind of data to try to read,
        # using the delimeter of that type of data to know what to do next.
        while nextChar != b'':
            # We should hav exactly one character, the start of the next
            # section. Use it to determine what to do.
            if nextChar == b'\\':
                newTokens, nextChar = self.__readTag(nextChar, reader)
            elif nextChar == b'{':
                # This will always be a group start, which has nothing left to
                # read.
                nextChar = reader.read()
                newTokens = (Token(b'{', TokenType.GROUP_START),)
            elif nextChar == b'}':
                # This will always be a group end, which has nothing left to
                # read.
                nextChar = reader.read()
                newTokens = (Token(b'}', TokenType.GROUP_END),)
            else:
                # Otherwise, it's just text.
                newTokens, nextChar = self.__readText(nextChar, reader)
            tokens.extend(newTokens)

        self.tokens = tokens
