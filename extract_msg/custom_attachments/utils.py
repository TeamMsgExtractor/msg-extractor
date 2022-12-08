"""
Utilities for extract-msg that are more specialized for the custom_attachments
submodule than for the main module.
"""

import bs4

from typing import List


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


def htmlSplitRendered(html : bytes) -> List[str]:
    """
    Takes html bytes and returns a list of the rendered characters, with data
    that is not being rendered being attached to the next rendered character.
    """
    # We use bs4 to convert the bytes to a string as accurately as possible. We
    # would also use it for tokenizing, but it doesn't allow for a quick and
    # easy way to do that and might just be faster to do ourselves.
    tokens = tokenizeHtml(bs4.BeautifulSoup(html).decode())
