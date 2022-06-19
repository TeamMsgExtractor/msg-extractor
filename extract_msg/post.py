from typing import Tuple

from . import constants
from .message_base import MessageBase
from .utils import inputToBytes, inputToString


class Post(MessageBase):
    """
    Class for parsing Post messages.
    """

    def __init__(self, path, **kwargs):
        super().__init__(path, **kwargs)
        # We don't need any special data, just need to override the save
        # functions for getting the body. Past that, we just let Message do all
        # the work.

    def getJson(self) -> str:
        """
        Returns the JSON representation of the Message.
        """
        return json.dumps({
            'from': inputToString(self.sender, 'utf-8'),
            'bcc': inputToString(self.bcc, 'utf-8'),
            'subject': inputToString(self.subject, 'utf-8'),
            'date': inputToString(self.date, 'utf-8'),
            'body': decode_utf7(self.body),
        })

    def getSaveBody(self, **kwargs) -> bytes:
        """
        Returns the plain text body that will be used in saving based on the
        arguments.

        :param kwargs: Used to allow kwargs expansion in the save function.
            Arguments absorbed by this are simply ignored.
        """
        # Get the type of line endings.
        crlf = inputToBytes(self.crlf, 'utf-8')

        outputBytes = b'From: ' + inputToBytes(self.sender, 'utf-8') + crlf
        outputBytes += b'Posted At: ' + inputToBytes(self.date, 'utf-8') + crlf
        outputBytes += b'Conversation: ' + inputToBytes(self.conversation, 'utf-8') + crlf
        outputBytes += b'Subject: ' + inputToBytes(self.subject, 'utf-8') + crlf
        outputBytes += b'-----------------' + crlf + crlf
        outputBytes += inputToBytes(self.body, 'utf-8')

        return outputBytes

    @property
    def conversation(self) -> str:
        """
        The name of the conversation being posted to.
        """
        return self._ensureSet('_convo', '__substg1.0_0070')

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
    def htmlInjectableHeader(self) -> str:
        """
        The header that can be formatted and injected into the html body.
        """
        return constants.HTML_INJECTABLE_HEADERS['Post']

    @property
    def rtfEncapInjectableHeader(self) -> str:
        """
        The header that can be formatted and injected into the plain RTF body.
        """
        return constants.RTF_ENC_INJECTABLE_HEADERS['Post']

    @property
    def rtfPlainInjectableHeader(self) -> str:
        """
        The header that can be formatted and injected into the encapsulated RTF
        body.
        """
        return constants.RTF_PLAIN_INJECTABLE_HEADERS['Post']
