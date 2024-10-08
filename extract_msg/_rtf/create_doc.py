__all__ = [
    'createDocument',
]


from typing import Iterable

from .token import Token, TokenType



def createDocument(tokens: Iterable[Token]) -> bytes:
    """
    Combines the tokenized data into bytes and returns the document.
    """
    document = b''

    # Recombining follows a few very basic rules that are based solely on the
    # token type. Since every token has the raw bytes, this is pretty easy. In
    # fact, control words are the only place where we put a space, as a space
    # anywhere else would be literal, and omitting a space could cause issues on
    # some control words.
    for token in tokens:
        if token.type in (TokenType.CONTROL, TokenType.DESTINATION, TokenType.IGNORABLE_DESTINATION):
            document += token.raw + b' '
        else:
            document += token.raw

    return document
