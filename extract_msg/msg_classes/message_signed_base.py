__all__ = [
    'MessageSignedBase',
]


import functools
import html
import logging
import re

from typing import List, Optional

from ..enums import ErrorBehavior
from ..exceptions import StandardViolationError
from .message_base import MessageBase
from ..attachments import SignedAttachment
from ..utils import inputToBytes, inputToString, unwrapMultipart


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class MessageSignedBase(MessageBase):
    """
    Base class for Message like msg files.
    """

    def __init__(self, path, **kwargs):
        """
        Supports all of the options from :method MessageBase.__init__: with some
        additional ones.

        :param signedAttachmentClass: optional, the class the object will use
            for signed attachments.
        """
        self.__signedAttachmentClass = kwargs.get('signedAttachmentClass', SignedAttachment)
        super().__init__(path, **kwargs)

    @property
    def attachments(self) -> List:
        """
        Returns a list of all attachments.

        :raises StandardViolationError: The standard for signed messages was
            blatantly violated.
        """
        try:
            return self._sAttachments
        except AttributeError:
            atts = super().attachments

            if len(atts) != 1:
                if ErrorBehavior.STANDARDS_VIOLATION in self.errorBehavior:
                    if len(atts) == 0:
                        logger.error('Signed message has no attachments, a violation of the standard.')
                        self._sAttachments = []
                        self._signedBody = None
                        self._signedHtmlBody = None
                        return
                    # If there is at least one attachment, just try to use the
                    # first.
                else:
                    raise StandardViolationError('Signed messages without exactly 1 (regular) attachment constitue a violation of the standard.')

            # We need to unwrap the multipart stream.
            unwrapped = unwrapMultipart(atts[0].data)

            # Now store everything where it needs to be and make the
            # attachments.
            self._sAttachments = [self.__signedAttachmentClass(self, **att) for att in unwrapped['attachments']]
            self._signedBody = unwrapped['plain_body']
            self._signedHtmlBody = inputToBytes(unwrapped['html_body'], 'utf-8')

            return self._sAttachments

    @property
    def body(self) -> Optional[str]:
        """
        Returns the message body, if it exists.
        """
        try:
            return self._body
        except AttributeError:
            if self._ensureSet('_body', '__substg1.0_1000'):
                pass
            elif self.signedBody:
                self._body = self.signedBody
            else:
                # If the body doesn't exist, see if we can get it from the RTF
                # body.
                if self.deencapsulatedRtf and self.deencapsulatedRtf.content_type == 'text':
                    self._body = self.deencapsulatedRtf.text

            if self._body:
                self._body = inputToString(self._body, 'utf-8')
                a = re.search('\n', self._body)
                if a is not None:
                    if re.search('\r\n', self._body) is not None:
                        self.__crlf = '\r\n'
            return self._body

    @property
    def htmlBody(self) -> Optional[bytes]:
        """
        Returns the html body, if it exists.
        """
        try:
            return self._htmlBody
        except AttributeError:
            if self._ensureSet('_htmlBody', '__substg1.0_10130102', False):
                # Reducing line repetition.
                pass
            elif self.signedHtmlBody:
                self._htmlBody = self.signedHtmlBody
            elif self.rtfBody:
                logger.info('HTML body was not found, attempting to generate from RTF.')
                if self.deencapsulatedRtf and self.deencapsulatedRtf.content_type == 'html':
                    self._htmlBody = self.deencapsulatedRtf.html.encode('utf-8')
                else:
                    logger.info('Could not deencapsulate HTML from RTF body.')
            elif self.body:
                # Convert the plain text body to html.
                logger.info('HTML body was not found, attempting to generate from plain text body.')
                correctedBody = html.escpae(self.body).replace('\r', '').replace('\n', '<br />')
                self._htmlBody = f'<html><body>{correctedBody}</body></head>'.encode('utf-8')
            else:
                logger.info('HTML body could not be found nor generated.')

            return self._htmlBody

    @functools.cached_property
    def _rawAttachments(self) -> List:
        """
        A property to allow access to the non-signed attachments.
        """
        return super().attachments

    @property
    def signedAttachmentClass(self):
        """
        The attachment class used for signed attachments.
        """
        return self.__signedAttachmentClass

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
