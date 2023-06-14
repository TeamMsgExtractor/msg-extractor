"""
Submodule for attachment classes.
"""

__all__ = [
    # Modules.
    'custom_attachments',

    # Classes.
    'Attachment',
    'AttachmentBase',
    'CustomAttachmentHandler',
    'SignedAttachment',

    # Functions.
    'registerHandler',
]


from . import custom_attachments
from .attachment import Attachment
from .attachment_base import AttachmentBase
from .custom_attachments import CustomAttachmentHandler, registerHandler
from .signed_attachment import SignedAttachment