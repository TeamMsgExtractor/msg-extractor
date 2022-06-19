import logging

from .message_base import MessageBase


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


# Due to changes to how saving works, Message is now identical to MessageBase,
# just with a different name. This is for backwards compatability, type
# identification, and possible future specializations to it.
class Message(MessageBase):
    """
    Parser for Microsoft Outlook message files.
    """

    def __init__(self, path, **kwargs):
        super().__init__(path, **kwargs)
