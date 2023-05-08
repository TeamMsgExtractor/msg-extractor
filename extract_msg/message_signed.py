import logging

from .message_signed_base import MessageSignedBase


__all__ = [
    'MessageSigned',
]


class MessageSigned(MessageSignedBase):
    """
    Parser for Signed Microsoft Outlook message files.
    """
