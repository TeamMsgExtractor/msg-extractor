"""
Utilities for extract-msg that are more specialized for the custom_attachments
submodule than for the main module.
"""

import bs4


def tokenizeHtml(html : str) -> List[str]:
    # Setup a few variables for state tracking.
    inTag = False
    inEscape = False
    inString = False
    # Only used when in string. Last character was a backslash.
    isBackslash = False

def htmlSplitRendered(html : bytes) -> List[str]:
    """
    Takes html bytes and returns a list of the rendered characters, with data
    that is not being rendered being attached to the next rendered character.
    """
    # We use bs4 to convert the bytes to a string as accurately as possible. We
    # would also use it for tokenizing, but it doesn't allow for a quick and
    # easy way to do that and might just be faster to do ourselves.
    tokens = tokenizeHtml(bs4.BeautifulSoup(html).decode())
