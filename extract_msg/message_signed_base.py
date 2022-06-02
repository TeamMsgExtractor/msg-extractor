import base64
import email.utils
import logging
import os
import re

import bs4
import compressed_rtf
import mailbits
import RTFDE

from . import constants
from .exceptions import UnrecognizedMSGTypeError
from .message_base import MessageBase
from .recipient import Recipient
from .signed_attachment import SignedAttachment
from .utils import addNumToDir, inputToBytes, inputToString, prepareFilename

from email.parser import Parser as EmailParser

from imapclient.imapclient import decode_utf7

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
        :param attachmentErrorBehavior: Optional, the behaviour to use in the
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

    def headerInit(self) -> bool:
        """
        Checks whether the header has been initialized.
        """
        try:
            self._header
            return True
        except AttributeError:
            return False

    @property
    def attachments(self) -> list:
        """
        Returns a list of all attachments.
        """
        try:
            return self._sAttachments
        except AttributeError:
            atts = super().attachments
            self._sAttachments = []
            self._signedBody = None
            self._signedHtmlBody = None
            if len(atts) > 1:
                logger.warn('')
            elif len(atts) == 0:
                logger.warn('Failed to access any attachments from the signed message.')
                return self._sAttachments

            try:
                mainAttachment = next(att for att in atts if hasattr(att, 'getFilename') and att.getFilename() == 'smime.p7m')
            except StopIteration:
                logger.warn('Failed to find signed attachment.')
                return self._sAttachments

            # If we are here, we should have the attachment. So now we need to
            # try to parse and unwrap the data.
            toParse = [mailbits.email2dict(email.message_from_bytes(mainAttachment.data))]
            output = []

            while len(toParse) != 0:
                parsing = toParse.pop(0)
                for part in parsing['content']:
                    # If it is multipart, push it to the toParse list, otherwise add it
                    # to the output.
                    if part['headers']['content-type']['content_type'].startswith('multipart'):
                        toParse.append(part)
                    else:
                        output.append(part)

            # At this point, `output` has out parts.
            for part in output:
                # Get the mime type.
                mime = part['headers']['content-type']['content_type']

                # Now try to grab a name. If it doesn't exist, we make one.
                try:
                    name = part['headers']['content-type']['params']['name']
                except KeyError:
                    if mime == 'text/plain':
                        self._signedBody = part['content']
                        continue
                    elif mime == 'text/html':
                        self._signedHtmlBody = part['content']
                        continue
                    else:
                        name = 'unknown.bin'
                self._sAttachments.append(self.__signedAttachmentClass(self, part['content'], name, mime))

            return self._sAttachments


    @property
    def bcc(self):
        """
        Returns the bcc field, if it exists.
        """
        return self._genRecipient('bcc', 3)

    @property
    def body(self):
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
    def htmlBody(self) -> bytes:
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
                correctedBody = self.body.encode('utf-8').replace('\r', '').replace('\n', '</br>')
                self._htmlBody = f'<html><body>{correctedBody}</body></head>'
            else:
                logger.info('HTML body could not be found nor generated.')

            return self._htmlBody

    @property
    def htmlBodyPrepared(self) -> bytes:
        """
        Returns the HTML body that has (where possible) the embedded attachments
        inserted into the body.
        """
        # If we can't get an HTML body then we have nothing to do.
        if not self.htmlBody:
            return self.htmlBody

        # Create the BeautifulSoup instance to use.
        soup = bs4.BeautifulSoup(self.htmlBody, 'html.parser')

        # Get a list of image tags to see if we can inject into. If the source
        # of an image starts with "cid:" that means it is one of the attachments
        # and is using the content id of that attachment.
        tags = (tag for tag in soup.findAll('img') if tag.get('src') and tag.get('src').startswith('cid:'))

        for tag in tags:
            # Iterate through the attachments until we get the right one.
            cid = tag['src'][4:]
            data = next((attachment.data for attachment in self.attachments if attachment.cid == cid), None)
            # If we found anything, inject it.
            if data:
                tag['src'] = (b'data:image;base64,' + base64.b64encode(data)).decode('utf-8')

        return soup.prettify('utf-8')

    @property
    def inReplyTo(self) -> str:
        """
        Returns the message id that this message is in reply to.
        """
        return self._ensureSet('_in_reply_to', '__substg1.0_1042')

    @property
    def isRead(self) -> bool:
        """
        Returns if this email has been marked as read.
        """
        return bool(self.mainProperties['0E070003'].value & 1)

    @property
    def messageId(self):
        try:
            return self._messageId
        except AttributeError:
            headerResult = None
            if self.headerInit():
                headerResult = self._header['message-id']
            if headerResult is not None:
                self._messageId = headerResult
            else:
                if self.headerInit():
                    logger.info('Header found, but "Message-Id" is not included. Will be generated from other streams.')
                self._messageId = self._getStringStream('__substg1.0_1035')
            return self._messageId

    @property
    def parsedDate(self):
        return email.utils.parsedate(self.date)

    @property
    def _rawAttachments(self):
        """
        A property to allow access to the non-signed attachments.
        """
        return super().attachments

    @property
    def recipientSeparator(self) -> str:
        return self.__recipientSeparator

    @property
    def recipients(self) -> list:
        """
        Returns a list of all recipients.
        """
        try:
            return self._recipients
        except AttributeError:
            # Get the recipients
            recipientDirs = []
            prefixLen = self.prefixLen
            for dir_ in self.listDir():
                if dir_[prefixLen].startswith('__recip') and\
                        dir_[prefixLen] not in recipientDirs:
                    recipientDirs.append(dir_[prefixLen])

            self._recipients = []

            for recipientDir in recipientDirs:
                self._recipients.append(Recipient(recipientDir, self))

            return self._recipients

    @property
    def rtfBody(self) -> bytes:
        """
        Returns the decompressed Rtf body from the message.
        """
        try:
            return self._rtfBody
        except AttributeError:
            self._rtfBody = compressed_rtf.decompress(self.compressedRtf) if self.compressedRtf else None
            return self._rtfBody

    @property
    def sender(self) -> str:
        """
        Returns the message sender, if it exists.
        """
        try:
            return self._sender
        except AttributeError:
            # Check header first
            if self.headerInit():
                headerResult = self.header['from']
                if headerResult is not None:
                    self._sender = headerResult
                    return headerResult
                logger.info('Header found, but "sender" is not included. Will be generated from other streams.')
            # Extract from other fields
            text = self._getStringStream('__substg1.0_0C1A')
            email = self._getStringStream('__substg1.0_5D01')
            # Will not give an email address sometimes. Seems to exclude the email address if YOU are the sender.
            result = None
            if text is None:
                result = email
            else:
                result = text
                if email is not None:
                    result += ' <' + email + '>'

            self._sender = result
            return result

    @property
    def signedAttachmentClass(self):
        """
        The attachment class used for signed attachments.
        """
        return self.__signedAttachmentClass

    @property
    def signedBody(self):
        """
        Returns the body from the signed message if it exists.
        """
        try:
            return self._signedBody
        except AttributeError:
            self.attachments
            return self._signedBody

    @property
    def signedHtmlBody(self):
        """
        Returns the HTML body from the signed message if it exists.
        """
        try:
            return self._signedHtmlBody
        except AttributeError:
            self.attachments
            return self._signedHtmlBody

    @property
    def subject(self):
        """
        Returns the message subject, if it exists.
        """
        return self._ensureSet('_subject', '__substg1.0_0037')

    @property
    def to(self):
        """
        Returns the to field, if it exists.
        """
        return self._genRecipient('to', 1)
