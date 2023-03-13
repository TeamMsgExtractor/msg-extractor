import copy

from .token import Token, TokenType
from .tokenize_rtf import tokenizeRTF

from typing import List, Optional, Union


# A tuple of destinations (including custom ones) used in the header.
_HEADER_DESTINATIONS = (
    'fonttbl'
)

# A tuple of control words that are part of the header and that we simply skip.
_HEADER_SKIPPABLE = (
    # Tag used to specify something.
    b'fbidis',
    # Character set tags.
    b'ansi',
    b'mac',
    b'pc',
    b'pca',
    b'ansicpg',
    # From.
    b'fromtext',
    b'fromhtml',
    # Def font.
    b'deff',
    b'adeff',
    b'stshfdbch',
    b'stshfloch',
    b'stshfhich',
    b'stshfbi',
    # Def lang.
    b'deflang',
    b'deflangfe',
    b'adeflang',

)


def injectStartRTF(document : bytes, injectTokens : Union[bytes, List[Token]]) -> List[Token]:
    """
    Injects the specified tokens into the document, returning a new copy of the
    document as a list of Tokens. Injects the data just before the first
    rendered character.

    :param document: The bytes representing the RTF document.
    :param injectTokens: The tokens to inject into the document. Can either be
        a list of Tokens or bytes to be tokenized.

    :raises TypeError: The data is not recognized as RTF.
    :raises ValueError: An issue with basic parsing occured.
    """
    return injectStartRTFTokenized(tokenizeRTF(document), injectTokens)


def injectStartRTFTokenized(document : List[Token], injectTokens : Union[bytes, List[Token]]) -> List[Token]:
    """
    Like :function injectStartRTF:, injects the specified tokens into the
    document, returning a reference to the document, except that it accepts a
    document in the form of a list of tokens. Injects the data just before the
    first rendered character.

    :param document: The list of tokens representing the RTF document. Will only
        be modified if the function is successful.
    :param injectTokens: The tokens to inject into the document. Can either be
        a list of Tokens or bytes to be tokenized.

    :raises TypeError: The data is not recognized as RTF.
    :raises ValueError: An issue with basic parsing occured.
    """
    # Get to a list of tokens to inject instead of
    if isinstance(injectTokens, bytes):
        injectTokens = tokenizeRTF(injectTokens, False)

    # Find the location to insert into. THis is annoyingly complicated, and we
    # do this by looking for the parts of the header (if they exist) as we go
    # token by token. The moment we confirm we are no longer in the header (and
    # we are not in a custom destination that we can simply ignore), we use the
    # last recorded spot as the insert point. We don't move that recorded spot
    # until we know that what we checked was part of the header.

    currentLocation = 0

    # First confirm the first two tokens are what we expect.
    if len(document < 3):
        raise ValueError('RTF documents cannot be less than 3 tokens.')
    if document[0].type != TokenType.GROUP_START or :


    # We have verified the minimal amount. Now,
