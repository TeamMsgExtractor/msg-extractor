__all__ = [
    'MessageSigned',
]


from .message_signed_base import MessageSignedBase


class MessageSigned(MessageSignedBase):
    """
    Parser for Signed Microsoft Outlook message files.
    """
