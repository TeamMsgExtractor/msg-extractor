__all__ = [
    'Token',
    'TokenType',
]


import enum

from typing import NamedTuple, Optional


class TokenType(enum.Enum):
    GROUP_START = 0
    GROUP_END = 1
    CONTROL = 2
    SYMBOL = 3
    TEXT = 4
    DESTINATION = 5
    IGNORABLE_DESTINATION = 6
    # This one is special, used for handling the binary data.
    BINARY = 7



class Token(NamedTuple):
    # The raw bytes for the token, used to recreate the document.
    raw: bytes
    # The type of the token.
    type: TokenType
    ## The following are optional as they only apply for certain types of tokens.
    # The name of the token, if it is a control or destination.
    name: Optional[bytes] = None
    # The parameter of the token, if it has one. If the token is a `\'hh` token,
    # this will be the decimal equivelent of the hex value.
    parameter: Optional[int] = None
