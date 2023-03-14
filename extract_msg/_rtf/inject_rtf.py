import copy

from .token import Token, TokenType
from .tokenize_rtf import tokenizeRTF

from typing import List, Iterable, Optional, Union


# A tuple of destinations used in the header. All ignorable ones are skipped
# anyways, so we don't need to list those here.
_HEADER_DESTINATIONS = (
    b'fonttbl',
    b'',
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


def _listInsertMult(dest : List, source : Iterable, index : int = -1):
    """
    Inserts into :param dest: all the items in :param source: at the index
    specified. :param dest: can be any mutable sequence with :method insert:,
    :method __len__:, and :method extend:.

    If :param index: is not specified, the default position is the end of the
    list. This is also where things will be inserted if index is greater than or
    equal to the size of the list.
    """
    if index == -1 or index >= len(dest):
        dest.extend(source)
    else:
        for offset, item in enumerate(source):
            dest.insert(index + offset, item)


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


def injectStartRTFTokenized(document : List[Token], injectTokens : Union[bytes, Iterable[Token]]) -> List[Token]:
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
    if document[0].type != TokenType.GROUP_START or document[1].raw != b'\\rtf1':
        raise TypeError('RTF document *must* start with "{\\rtf1".')

    # Confirm that all start groups have an end group somewhere.
    if sum(x.type == TokenType.GROUP_START for x in document) != sum(x.type == TokenType.GROUP_END for x in document):
        raise ValueError('Number of group opens did not match number of group closes.')

    # If the length is exactly 3, insert right before the end and return.
    if len(document) == 3:
        _listInsertMult(document, injectTokens, 2)
        return document

    # We have verified the minimal amount. Now, iterate through the rest to find
    # the injection point.
