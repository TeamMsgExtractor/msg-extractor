__all__ = [
    'MessageSignedBase',
]


import html
import logging
import re

from typing import List, Optional

from .enums import ErrorBehavior
from .exceptions import StandardViolationError
from .message_base import MessageBase
from .signed_attachment import SignedAttachment
from .utils import inputToBytes, inputToString, unwrapMultipart


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class MessageSignedBase(MessageBase):
    """
    Base class for Message like msg files.
    """

    def __init__(self, path, **kwargs):
        """
        :param path: path to the msg file in the system or is the raw msg file.
        :param prefix: used for extracting embeded msg files
            inside the main one. Do not set manually unless
            you know what you are doing.
        :param attachmentClass: optional, the class the Message object
            will use for attachments. You probably should
            not change this value unless you know what you
            are doing.
        :param signedAttachmentClass: optional, the class the object will use
            for signed attachments.
        :param filename: optional, the filename to be used by default when
            saving.
        :param delayAttachments: optional, delays the initialization of
            attachments until the user attempts to retrieve them. Allows MSG
            files with bad attachments to be initialized so the other data can
            be retrieved.
        :param overrideEncoding: optional, an encoding to use instead of the one
            specified by the msg file. Do not report encoding errors caused by
            this.
        :param attachmentErrorBehavior: Optional, the behavior to use in the
            event of an error when parsing the attachments.
        :param recipientSeparator: Optional, Separator string to use between
            recipients.
        """
        self.__recipientSeparator = kwargs.get('recipientSeparator', ';')
        self.__signedAttachmentClass = kwargs.get('signedAttachmentClass', SignedAttachment)
        super().__init__(path, **kwargs)
        # Initialize properties in the order that is least likely to cause bugs.
        # TODO have each function check for initialization of needed data so these
        # lines will be unnecessary.
        if not kwargs.get('delayAttachments', False):
            self.attachments

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
                if self.errorBehavior & ErrorBehavior.STANDARDS_VIOLATION:
                    if len(atts) == 0:
                        logger.error('Signed message has no attachments, a violation of the standard.')
                        self._sAttachments = None
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

    @property
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

    @property
    def signedBody(self) -> Optional[str]:
        """
        Returns the body from the signed message if it exists.
        """
        try:
            return self._signedBody
        except AttributeError:
            self.attachments
            return self._signedBody

    @property
    def signedHtmlBody(self) -> Optional[bytes]:
        """
        Returns the HTML body from the signed message if it exists.
        """
        try:
            return self._signedHtmlBody
        except AttributeError:
            self.attachments
            return self._signedHtmlBody
