from .token import Token, TokenType
from .tokenize_rtf import tokenizeRTF

from typing import List, Optional, Union


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
