__all__ = [
    'Post',
]


import functools
import json

from typing import Optional

from .. import constants
from .message_base import MessageBase


class Post(MessageBase):
    """
    Class for parsing Post messages.
    """

    def getJson(self) -> str:
        """
        Returns the JSON representation of the Post.
        """
        return json.dumps({
            'from': self.sender,
            'subject': self.subject,
            'date': self.date.__format__(self.datetimeFormat) if self.date else None,
            'conversation': self.conversation,
            'body': self.body,
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
                'Posted At': self.date.__format__(self.datetimeFormat) if self.date else None,
                'Conversation': self.conversation,
            },
            '-subject-': {
                'Subject': self.subject,
            },
            '-importance-': {
                'Importance': self.importanceString,
            },
        }
