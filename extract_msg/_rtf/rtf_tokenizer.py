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

    def feed(self, data : bytes) -> None:
        """
        Reads in the bytes and sets the tokens list to the contents after
        tokenizing. If tokenizing fails, the current tokens list will not be
        changed.

        Direct references to the previous tokens list will only point to the
        previous and not to the current one.

        :raises TypeError: The data is not recognized as RTF.
        """
        reader = io.BytesIO(data)
        # This tokenizer *only* breaks things up. It does *not* care about
        # groups and stuff, as that is for a parser to deal with. All we do is
        # track the current backslash state and token state. We also simply
        # check that the first token is "\rtf1" preceeded by a group start, and
        # that is it.
        start = reader.read(6)
        if start != b'{\\rtf1':
            raise TypeError('Data')

        tokens = [Token(b'{'), Token(b'\rtf')]



class TokenType(enum.Enum):
    GROUP_START = 0
    GROUP_END = 1
    CONTROL = 2
    SYMBOL = 3
    DESTINATION = 4



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
