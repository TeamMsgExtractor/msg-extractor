__all__ = [
    'MessageSigned',
]


from typing import TypeVar

from .message_signed_base import MessageSignedBase


_T = TypeVar('_T')


class MessageSigned(MessageSignedBase[_T]):
    """
    Parser for Signed Microsoft Outlook message files.
    """
