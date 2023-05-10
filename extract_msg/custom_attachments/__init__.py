from __future__ import annotations


"""
Submodule designed to help with saving and using custom attachments. Custom
attachments are those follow standards not defined in the MSG documentation. Use
the function `getHandler` to get an instance of a subclass of
CustomAttachmentHandler.

CustomAttachmentHandler subclasses will all define the following methods:
    injectHtml: A method which takes HTML and inserts the

It should hopefully be completely unnecessary for your code to know what type of
handler it is using, as the abstract base class should give all of the functions
you would typically want.

If you would like to add your own handler, simply subclass
CustomAttachmentHandler and add it using the `registerHandler` function.
"""

__all__ = [
    # Classes.
    'CustomAttachmentHandler',
    'OutlookImage',

    # Functions.
    'getHandler',
    'registerHandler',
]


from typing import List, TYPE_CHECKING

from .custom_handler import CustomAttachmentHandler


# Create a way to register handlers.
_knownHandlers : List[CustomAttachmentHandler] = []
# This line is cheating a little bit, but is more efficient than wrapping it in
# a function.
registerHandler = _knownHandlers.append


# Import built-in handler modules. They will all automatically register their
# respecive handler(s).
from .outlook_image import OutlookImage


if TYPE_CHECKING:
    from ..attachment import Attachment


# Function designed to route to the correct handler.
def getHandler(attachment : Attachment) -> CustomAttachmentHandler:
    """
    Takes an attachment and uses it to find the correct handler. Returns an
    instance created using the specified attachment.

    :raises NotImplementedError: No handler could be found.
    :raises ValueError: A handler was found, but something was wrong with the
        attachment data.
    """
    for handler in _knownHandlers:
        if handler.isCorrectHandler(attachment):
            return handler(attachment)

    raise NotImplementedError('No valid handler could be found for the attachment. Contact the developers for help.')
