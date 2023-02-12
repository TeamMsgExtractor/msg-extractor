import copy
import enum
import io

from typing import List, NamedTuple


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
        if nextChar != ' ':
            reader.seek(reader.tell() - 1)

        # Setup the loop variables.
        lastCharacter = b''
        inTag = False
        param = b''
        raw = b''
        # Tracking for if we are handling a `\'HH` tag.
        isHex = False

        while (currentChar := reader.read()) != b'':
            if currentChar == b' ' and inTag:
                # End the tag and drop the space.
                tokens.append(self.__handleTag(raw, param))
                raw = b''
                param = b''
                inTag = False
            elif currentChar in (b'{', b'}'):
                # Brackets are second fastest to parse, so do them next.
                if inTag:
                    if raw == b'\\':
                        # If we only have the backslash, this is a symbol.
                        tokens.append(self.__handleTag(b'\\' + currentChar, b''))
                        raw = b''
                        param = b''
                        inTag = False
                    else:
                        # We already have a currentTag, so we need to push it
                        # then push a group.
                        tokens.append(self.__handleTag(raw, param))
                        raw = b''
                        param = b''
                        inTag = False
                        if currentChar = b'{':
                            tokens.append(Token(currentChar, TokenType.GROUP_START))
                        else:
                            tokens.append(Token(currentChar, TokenType.GROUP_END))
                else:
                    if raw:
                        tokens.append(Token(raw, TokenType.TEXT))
                        raw = b''
                    if currentChar = b'{':
                        tokens.append(Token(currentChar, TokenType.GROUP_START))
                    else:
                        tokens.append(Token(currentChar, TokenType.GROUP_END))
            elif currentChar == b'*' and inTag:
                if len(raw) == 1:
                    raw += b'*'
                else:
                    # End the current tag and start new text.
                    tokens.append(self.__handleTag(raw, param))
                    raw = b'*'
                    param = b''
                    inTag = False
            elif currentChar == b'\\':
                if inTag:
                    if lastCharacter == b'\\':
                        # If the current character is a backslash and we are in
                        # a tag, then it is a backslash symbol and we need to
                        # push it to the list.
                        tokens.append(self.__handleTag(raw + b'\\', param))
                        raw = b''
                        param = b''
                    elif lastCharacter == b'*':
                        # This is a custom destination, but we aren't handling
                        # that in this section, so just add to the current tag.
                        raw += currentChar
                    else:
                        # We are starting a new tag.
                        tokens.append(self.__handleTag(raw, param))
                        raw = b'\\'
                        param = b''
                else:
                    # If we were not already in a tag, check if we have text to
                    # push, and if we do we push it. After that, start a tag.
                    if raw:
                        tokens.append()
                    raw += b'\\'
                    inTag = True
            elif currentChar == b'\'' and inTag:
                # If the current character is an apostrophe and we are in a tag,
                # check if it is the first character after the backslash. If it
                # is, then we are handling hex character, otherwise we are

            elif currentChar.isalpha():
                pass


        lastCharacter = currentChar

    # Since we are done, set the tokens list to the new one.
    self.tokens = tokens



class TokenType(enum.Enum):
    GROUP_START = 0
    GROUP_END = 1
    CONTROL = 2
    SYMBOL = 3
    DESTINATION = 4
    TEXT = 5



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
