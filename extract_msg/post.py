from . import constants
from .message import Message


# Post is effectively identical to Message, just with a different header.
class Post(Message):
    """
    Class for parsing Post messages.
    """

    def __init__(self, path, **kwargs):
        super().__init__(path, **kwargs)
        # We don't need any special data, just need to override the save
        # functions for getting the body. Past that, we just let Message do all
        # the work.

    @property
    def headerFormatProperties(self) -> Tuple[Tuple[str, str], ...]:
        """
        Returns a tuple of tuples of two strings, the first string being a
        name and the second being one of the properties of this class. These are
        used for controlling how data is inserted into the headers when saving
        an MSG file. The names used for the first string in the tuple will
        correspond to the name used for the format string.

        If you need to override the default behavior, override this in your
        class.
        """
        return (
            ('sender', 'sender'),
            ('subject', 'subject'),
            ('convo', 'conversation'),
            ('date', 'date'),
        )

    @property
    def htmlInjectableHeader(self):
        """
        The header that can be formatted and injected into the html body.
        """
        return constants.HTML_INJECTABLE_HEADERS['Post']

    @property
    def rtfEncapInjectableHeader(self):
        """
        The header that can be formatted and injected into the plain RTF body.
        """
        return constants.RTF_ENC_INJECTABLE_HEADER['Post']

    @property
    def rtfPlainInjectableHeader(self):
        """
        The header that can be formatted and injected into the encapsulated RTF
        body.
        """
        return constants.RTF_PLAIN_INJECTABLE_HEADER['Post']
