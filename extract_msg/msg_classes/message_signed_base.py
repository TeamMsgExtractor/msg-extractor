__all__ = [
    'MessageSignedBase',
]


import functools
import html
import logging
import re

from typing import Generic, List, Optional, Type, TypeVar

from ..attachments import AttachmentBase, SignedAttachment
from ..enums import DeencapType, ErrorBehavior
from ..exceptions import StandardViolationError
from .message_base import MessageBase
from ..utils import inputToBytes, inputToString, unwrapMultipart


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

_T = TypeVar('_T')


class MessageSignedBase(MessageBase, Generic[_T]):
    """
    Base class for Message like msg files.
    """

    def __init__(self, path, signedAttachmentClass : Type[_T] = SignedAttachment, **kwargs):
        """
        Supports all of the options from :method MessageBase.__init__: with some
        additional ones.

        :param signedAttachmentClass: optional, the class the object will use
            for signed attachments.
        """
        self.__sAttCls = signedAttachmentClass
        super().__init__(path, **kwargs)

    @functools.cached_property
    def attachments(self) -> List[_T]:
        """
        Returns a list of all attachments.

        :raises StandardViolationError: The standard for signed messages was
            blatantly violated.
        """
        atts = self._rawAttachments

        if len(atts) != 1:
            if ErrorBehavior.STANDARDS_VIOLATION in self.errorBehavior:
                if len(atts) == 0:
                    logger.error('Signed message has no attachments, a violation of the standard.')
                    self._sAttachments = []
                    self._signedBody = None
                    self._signedHtmlBody = None
                    return []
                # If there is at least one attachment, just try to use the
                # first.
            else:
                raise StandardViolationError('Signed messages without exactly 1 (regular) attachment constitute a violation of the standard.')

        # We need to unwrap the multipart stream.
        unwrapped = unwrapMultipart(atts[0].data)

        # Now store everything where it needs to be and make the
        # attachments.
        self._signedBody = unwrapped['plain_body']
        self._signedHtmlBody = inputToBytes(unwrapped['html_body'], 'utf-8')

        return [self.__sAttCls(self, **att) for att in unwrapped['attachments']]

    @functools.cached_property
    def body(self) -> Optional[str]:
        """
        Returns the message body, if it exists.
        """
        if (body := self.getStringStream('__substg1.0_1000')) is not None:
            pass
        elif self.signedBody:
            body = self.signedBody
        elif self.rtfBody:
            # If the body doesn't exist, see if we can get it from the RTF
            # body.
            body = self.deencapsulateBody(self.rtfBody, DeencapType.PLAIN)

        if body:
            body = inputToString(body, 'utf-8')
            if re.search('\n', body):
                if re.search('\r\n', body):
                    self._crlf = '\r\n'

        return body

    @functools.cached_property
    def htmlBody(self) -> Optional[bytes]:
        """
        Returns the html body, if it exists.
        """
        if (htmlBody := self.getStream('__substg1.0_10130102')) is not None:
            pass
        elif self.signedHtmlBody:
            htmlBody = self.signedHtmlBody
        elif self.rtfBody:
            logger.info('HTML body was not found, attempting to generate from RTF.')
            htmlBody = self.deencapsulateBody(self.rtfBody, DeencapType.HTML)
        elif self.body:
            # Convert the plain text body to html.
            logger.info('HTML body was not found, attempting to generate from plain text body.')
            correctedBody = html.escpae(self.body).replace('\r', '').replace('\n', '<br />')
            htmlBody = f'<html><body>{correctedBody}</body></head>'.encode('utf-8')
        else:
            logger.info('HTML body could not be found nor generated.')

        return htmlBody

    @functools.cached_property
    def _rawAttachments(self) -> List[AttachmentBase]:
        """
        A property to allow access to the non-signed attachments.
        """
        return super().attachments

    @property
    def signedAttachmentClass(self) -> Type[_T]:
        """
        The attachment class used for signed attachments.
        """
        return self.__sAttCls

    @functools.cached_property
    def signedBody(self) -> Optional[str]:
        """
        Returns the body from the signed message if it exists.
        """
        self.attachments
        return self._signedBody

    @functools.cached_property
    def signedHtmlBody(self) -> Optional[bytes]:
        """
        Returns the HTML body from the signed message if it exists.
        """
        self.attachments
        return self._signedHtmlBody
