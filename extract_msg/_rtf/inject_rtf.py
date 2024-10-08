__all__ = [
    'injectStartRTF',
    'injectStartRTFTokenized',
]


from .token import Token, TokenType
from .tokenize_rtf import tokenizeRTF

from typing import List, Iterable, TypeVar, Union


_T = TypeVar('_T')

# A tuple of destinations used in the header. All ignorable ones are skipped
# anyways, so we don't need to list those here.
_HEADER_DESTINATIONS = (
    b'fonttbl',
    b'colortbl',
    b'stylesheet',
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


def _listInsertMult(dest: List[_T], source: Iterable[_T], index: int = -1):
    """
    Inserts into :param dest: all the items in :param source: at the index
    specified.

    :param dest: Any mutable sequence with methods ``insert``, ``__len__``, and
        ``extend``.

    If :param index: is not specified, the default position is the end of the
    list. This is also where things will be inserted if index is greater than or
    equal to the size of the list.
    """
    if index == -1 or index >= len(dest):
        dest.extend(source)
    else:
        for offset, item in enumerate(source):
            dest.insert(index + offset, item)


def injectStartRTF(document: bytes, injectTokens: Union[bytes, List[Token]]) -> List[Token]:
    """
    Injects the specified tokens into the document, returning a new copy of the
    document as a list of Tokens.

    Injects the data just before the first rendered character.

    :param document: The bytes representing the RTF document.
    :param injectTokens: The tokens to inject into the document. Can either be
        a list of ``Token``\\s or ``bytes`` to be tokenized.

    :raises TypeError: The data is not recognized as RTF.
    :raises ValueError: An issue with basic parsing occured.
    """
    return injectStartRTFTokenized(tokenizeRTF(document), injectTokens)


def injectStartRTFTokenized(document: List[Token], injectTokens: Union[bytes, Iterable[Token]]) -> List[Token]:
    """
    Like :function injectStartRTF:, injects the specified tokens into the
    document, returning a reference to the document, except that it accepts a
    document in the form of a list of tokens.

    Injects the data just before the first rendered character.

    :param document: The list of tokens representing the RTF document. Will only
        be modified if the function is successful.
    :param injectTokens: The tokens to inject into the document. Can either be
        a list of ``Token``\\s or ``bytes`` to be tokenized.

    :raises TypeError: The data is not recognized as RTF.
    :raises ValueError: An issue with basic parsing occured.
    """
    # Get to a list of tokens to inject instead of
    if isinstance(injectTokens, bytes):
        injectTokens = tokenizeRTF(injectTokens, False)

    # Find the location to insert into. This is annoyingly complicated, and we
    # do this by looking for the parts of the header (if they exist) as we go
    # token by token. The moment we confirm we are no longer in the header (and
    # we are not in a custom destination that we can simply ignore), we use the
    # last recorded spot as the insert point. We don't move that recorded spot
    # until we know that what we checked was part of the header.

    # First confirm the first two tokens are what we expect.
    if len(document) < 3:
        raise ValueError('RTF documents cannot be less than 3 tokens.')
    if document[0].type is not TokenType.GROUP_START or document[1].raw != b'\\rtf1':
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
    currentInsertPos = 2
    # Current number of open groups.
    groupCount = 1
    # Set to True when looking for if the group is a destination.
    checkingDest = False

    for item in document[2:]:
        if groupCount == 1:
            if item.type is TokenType.GROUP_END:
                break
            elif item.type is TokenType.GROUP_START:
                groupCount += 1
                checkingDest = True
            elif item.type is TokenType.CONTROL and item.name in _HEADER_SKIPPABLE:
                # If the control is one we know about in the header, skip it.
                currentInsertPos += 1
            else:
                # Anything else means we are out of the header.
                break
        elif checkingDest:
            if item.type is TokenType.DESTINATION:
                # If it is *not* a header destination, just break, otherwise add
                # 2 to the insert location and skip the destination.
                if item.name in _HEADER_DESTINATIONS:
                    currentInsertPos += 2
                else:
                    break
            elif item.type is TokenType.IGNORABLE_DESTINATION:
                # Add 2 to insert location and skip.
                currentInsertPos += 2
            else:
                # If it is not an ignorible destination, we are now out of the
                # header, so break.
                break
            checkingDest = False
        else:
            # Skip the current token, keeping track of groups.
            if item.type is TokenType.GROUP_START:
                groupCount += 1
            if item.type is TokenType.GROUP_END:
                groupCount -= 1
            currentInsertPos += 1

    _listInsertMult(document, injectTokens, currentInsertPos)
    return document
