"""
Utilities for extract-msg that are more specialized for the custom_attachments
submodule than for the main module.
"""

import bs4

from typing import List


_WHITESPACE_BREAKERS = (
    '<img',
    '<span',
)

_WHITESPACE_TAGS = (
    '<br',
)


def _isWhitespaceBreaker(token : str) -> bool:
    """
    Helper function to indicate that a tag breaks a chain of whitespace.
    """
    for x in _WHITESPACE_BREAKERS:
        if token.startswith(x) and len(token) > len(x) and token[len(x)] in ('>', ' ', '/'):
            return True

    return False


def _isWhitespaceToken(token : str) -> bool:
    if token[0] == '<':
        for x in _WHITESPACE_TAGS:
            if token.startswith(x) and len(token) > len(x) and token[len(x)] in ('>', ' ', '/'):
                return True
    elif token in ('&#32;', '&Tab;', '&NewLine;', '&#9;', '&#10;'):
        return True
    else:
        return token.isspace()

    return False


def htmlSplitRendered(html : bytes) -> List[str]:
    """
    Takes html bytes and returns a list of the rendered characters, with data
    that is not being rendered being attached to the next rendered character.
    """
    # Unfortunately bs4 didn't seem particularly great for tokenizing, so I did
    # my own function that works well enough. First, let's tokenize the html.
    tokens = tokenizeHtml(bs4.BeautifulSoup(html, features = 'html.parser').decode())

    # Next, let's break things down further.
    breakDown = []
    for token in tokens:
        # We tell what we are looking at by checking the first character of the
        # token. If it's a <, then it is an HTML tag. If it is a & then it is an
        # escape. Otherwise, it is plain text. For both tags and escapes, just
        # dump them into the list.
        if token[0] in ('<', '&'):
            breakDown.append(token)
        else:
            # If we are looking at plain text, add it by extending the list.
            breakDown.extend(token)

    # Now that we have broken things down further, let's go through and join our
    # pieces togethered into rendered tokens. Here is were we actually need to
    # know what an html tag is. If it's an escape or just a non-whitespace
    # character, we can just shove it onto what we currently have.
    current = ''
    renderedCharacters = []
    lastWhitespace = None
    for item in breakDown:
        if item[0] == '&':
            if _isWhitespaceToken(item):
                if lastWhitespace is None:
                    lastWhitespace == item
            else:
                if lastWhitespace is not None and lastWhitespace[0] != '<':
                    current += lastWhitespace
                    renderedCharacters.append(current)
                    current = ''
                    lastWhitespace = None
                current += item
                renderedCharacters.append(current)
                current = ''
        elif item[0] == '<':
            if _isWhitespaceToken(item):
                # If we are here, add it to current, push current, and set this
                # tag as the last whitespace.
                current += item
                renderedCharacters.append(current)
                current = ''
                lastWhitespace = item
            else:

                # Some tags will break whitespace chains.
                if _isWhitespaceBreaker(item):
                    if lastWhitespace is not None and lastWhitespace[0] != '<':
                        current += lastWhitespace
                        renderedCharacters.append(current)
                        current = ''
                        lastWhitespace = None

                current += item
        else:
            # Here is where we handle text, which is not particularly fun.
            # Basically if it is whitespace and lastWhitespace is not none, we
            # set the whitespace.
            if _isWhitespaceToken(item):
                if lastWhitespace is None:
                    lastWhitespace = item
            else:
                if lastWhitespace is not None and lastWhitespace[0] != '<':
                    current += lastWhitespace
                    renderedCharacters.append(current)
                    current = ''
                lastWhitespace = None
                current += item
                renderedCharacters.append(current)
                current = ''

    if current:
        renderedCharacters.append(current)

    return renderedCharacters


def tokenizeHtml(html : str) -> List[str]:
    # Setup a few variables for state tracking.
    inTag = False
    # Used for tracking escapes starting with &. If your escape ends up at 100
    # characters because it is missing the semicolon, we are going to throw an
    # error.
    inEscape = False
    inString = False
    # Only used when in string. Last character was a backslash.
    isBackslash = False
    # Tells which quote type we are in.
    isDoubleQuote = False

    tokens = []
    currentToken = ''

    # Finally, let's start breaking things up. Our rules are that if we start a
    # quote while in a tag, then we acknowledge it, otherwise it is treated as
    # plain text.
    for character in html:
        # First we need to know our state, as our state determines what how we
        # process a character.
        if inTag:
            currentToken += character
            if inString:
                if character == '"' and isDoubleQuote:
                    # If isBackslash then we stay in the quote, otherwise...
                    isQuote = isBackslash
                elif character == "'" and not isDoubleQuote:
                    # If isBackslash then we stay in the quote, otherwise...
                    isQuote = isBackslash
                elif character == '\\':
                    isBackslash = not isBackslash
                if character != '\\':
                    isBackslash = False
            else:
                if character == '>':
                    inTag = False
                    tokens.append(currentToken)
                    currentToken = ''
        elif inEscape:
            currentToken += character
            if len(currentToken) > 99:
                raise ValueError('Found escape that was too long (is a ; missing?)')
            if character == ';':
                tokens.append(currentToken)
                currentToken = ''
                inEscape = False
        elif inString:
            # This is an error. We should *never* be in a quote if we are not in
            # a tag.
            raise ValueError('Found to be inQuote when not in tag.')
        else:
            # We are currently processing plain text, so let's just handle.
            if character == '&':
                if currentToken:
                    tokens.append(currentToken)
                currentToken = character
                inEscape = True
            elif character == '<':
                if currentToken:
                    tokens.append(currentToken)
                currentToken = character
                inTag = True
                inString = False
            else:
                currentToken += character

    if currentToken:
        tokens.append(currentToken)

    return tokens
