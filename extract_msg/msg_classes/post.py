__all__ = [
    'Post',
]


import functools
import json

from typing import Optional

from .. import constants
from .message_base import MessageBase
from ..utils import inputToString

from imapclient.imapclient import decode_utf7


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
            'conversation': inputToString(self.conversation, self.stringEncoding),
            'body': decode_utf7(self.body),
        })

    @functools.cached_property
    def conversation(self) -> Optional[str]:
        """
        The name of the conversation being posted to.
        """
        return self._getStringStream('__substg1.0_0070')

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

        Additionally you can group members of a header together by placing them
        in an embedded dictionary. Groups will be spaced out using a second
        instance of the join string. If any member of a group is being printed,
        it will be spaced apart from the next group/item.

        If you class should not do *any* header injection, return None from this
        property.
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
