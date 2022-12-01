"""
Submodule designed to help with saving and using custom attachments. Custom
attachments are those follow standards not defined in the MSG documentation. Use
the function `getHandler` to get an instance of a subclass of
CustomAttachmentHandler.

CustomAttachmentHandler subclasses will all define the following methods:
    injectHtml: A method which takes HTML and inserts the
"""

from custom_handler import CustomAttachmentHandler
from outlook_signature import OutlookSignature


# Function designed to route to the correct handler.
def getHandler(attachment : 'Attachment'):
    """
    Takes an attachment and uses it to find the correct hanlder.
    """
