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
            'date': inputToString(self.date.__format__(self.datetimeFormat), self.stringEncoding),
            'conversation': inputToString(self.conversation, self.stringEncoding),
            'body': decode_utf7(self.body),
        })

    @functools.cached_property
    def conversation(self) -> Optional[str]:
        """
        The name of the conversation being posted to.
        """
        return self.getStringStream('__substg1.0_0070')

    @property
    def headerFormatProperties(self) -> constants.HEADER_FORMAT_TYPE:
        return {
            '-main details-': {
                'From': self.sender,
                'Posted At': self.date.__format__(self.datetimeFormat),
                'Conversation': self.conversation,
            },
            '-subject-': {
                'Subject': self.subject,
            },
            '-importance-': {
                'Importance': self.importanceString,
            },
        }
