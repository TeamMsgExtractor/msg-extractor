__all__ = [
    'Message',
]


from .message_base import MessageBase


class Message(MessageBase):
    """
    Parser for Microsoft Outlook message files.
    """
