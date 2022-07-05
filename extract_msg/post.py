from . import constants
from .message_base import MessageBase
from .utils import inputToBytes, inputToString


class Post(MessageBase):
    """
    Class for parsing Post messages.
    """

    def getJson(self) -> str:
        """
        Returns the JSON representation of the Post.
        """
        return json.dumps({
            'from': inputToString(self.sender, self.stringEncoding),
            'subject': inputToString(self.subject, self.stringEncoding),
            'date': inputToString(self.date, self.stringEncoding),
            'conversation': inputTostring(self.conversation, self.stringEncoding),
            'body': decode_utf7(self.body),
        })

    @property
    def conversation(self) -> str:
        """
        The name of the conversation being posted to.
        """
        return self._ensureSet('_convo', '__substg1.0_0070')

    @property
    def headerFormatProperties(self) -> constants.HEADER_FORMAT_TYPE:
        """
        Returns a dictionary of properties, in order, to be formatted into the
        header. Keys are the names to use in the header while the values are one
        of the following:
        None: Signifies no data was found for the property and it should be
            omitted from the header.
        str: A string to be formatted into the header using the string encoding.
        Tuple[Union[str, None], bool]: A string should be formatted into the
            header. If the bool is True, then place an empty string if the value
            is None, otherwise follow the same behavior as regular None.

        Additional note: If the value is an empty string, it will be dropped as
        well by default.
        """
        return {
            '-main details-': {
                'From': self.sender,
                'Posted At': self.date,
                'Conversation': self.conversation,
            },
            '-subject-': {
                'Subject': self.subject,
            },
            '-importance-': {
                'Importance': self.importanceString,
            },
        }
