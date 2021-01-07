import email.utils
import json
import logging
import re

import compressed_rtf
from imapclient.imapclient import decode_utf7

from email.parser import Parser as EmailParser
from extract_msg import constants
from extract_msg.attachment import Attachment, BrokenAttachment, UnsupportedAttachment
from extract_msg.compat import os_ as os
from extract_msg.msg import MSGFile
from extract_msg.recipient import Recipient
from extract_msg.utils import addNumToDir, inputToBytes, inputToString



logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

class MessageBase(MSGFile):
    """
    Base class for Message like msg files.
    """

    def __init__(self, path, prefix = '', attachmentClass = Attachment, filename = None,
                 delayAttachments = False, overrideEncoding = None,
                 attachmentErrorBehavior = constants.ATTACHMENT_ERROR_THROW):
        """
        :param path: path to the msg file in the system or is the raw msg file.
        :param prefix: used for extracting embeded msg files
            inside the main one. Do not set manually unless
            you know what you are doing.
        :param attachmentClass: optional, the class the Message object
            will use for attachments. You probably should
            not change this value unless you know what you
            are doing.
        :param filename: optional, the filename to be used by default when saving.
        :param delayAttachments: optional, delays the initialization of attachments
            until the user attempts to retrieve them. Allows MSG files with bad
            attachments to be initialized so the other data can be retrieved.
        :param overrideEncoding: optional, an encoding to use instead of the one
            specified by the msg file. Do not report encoding errors caused by this.
        """
        MSGFile.__init__(self, path, prefix, attachmentClass, filename, overrideEncoding, attachmentErrorBehavior)
        self.__attachmentsDelayed = delayAttachments
        self.__attachmentsReady = False
        # Initialize properties in the order that is least likely to cause bugs.
        # TODO have each function check for initialization of needed data so these
        # lines will be unnecessary.
        self.mainProperties
        self.header
        self.recipients
        if not delayAttachments:
            self.attachments
        self.to
        self.cc
        self.sender
        self.date
        self.__crlf = '\n'  # This variable keeps track of what the new line character should be
        self.body
        self.named

    def _genRecipient(self, recipientType, recipientInt):
        """
        Returns the specified recipient field
        """
        private = '_' + recipientType
        try:
            return getattr(self, private)
        except AttributeError:
            # Check header first
            value = None
            if self.headerInit():
                value = self.header[recipientType]
            if value is None:
                if self.headerInit():
                    logger.info('Header found, but "{}" is not included. Will be generated from other streams.'.format(recipientType))
                f = []
                for x in self.recipients:
                    if x.type & 0x0000000f == recipientInt:
                        f.append(x.formatted)
                if len(f) > 0:
                    st = f[0]
                    if len(f) > 1:
                        for x in range(1, len(f)):
                            st += ', {0}'.format(f[x])
                    value = st
            if value is not None:
                value = value.replace(' \r\n\t', ' ').replace('\r\n\t ', ' ').replace('\r\n\t', ' ')
                value = value.replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
                while value.find('  ') != -1:
                    value = value.replace('  ', ' ')
            setattr(self, private, value)
            return value

    def _registerNamedProperty(self, entry, _type, name = None):
        if self.attachmentsDelayed and not self.attachmentsReady:
            try:
                self.__waitingProperties
            except AttributeError:
                self.__waitingProperties = []
            self.__waitingProperties.append((entry, _type, name))
        else:
            for attachment in self.attachments:
                attachment._registerNamedProperty(entry, _type, name)

    def close(self):
        try:
            self._attachments
            for attachment in self.attachments:
                if attachment.type == 'msg':
                    attachment.data.close()
        except AttributeError:
            pass
        MSGFile.close(self)

    def headerInit(self):
        """
        Checks whether the header has been initialized.
        """
        try:
            self._header
            return True
        except AttributeError:
            return False

    def save_attachments(self, contentId = False, json = False, useFileName = False, raw = False, customPath = None):
        """
        Saves only attachments in the same folder.
        """
        for attachment in self.attachments:
            attachment.save(contentId, json, useFileName, raw, customPath)

    @property
    def attachments(self):
        """
        Returns a list of all attachments.
        """
        try:
            return self._attachments
        except AttributeError:
            # Get the attachments
            attachmentDirs = []

            for dir_ in self.listDir(False, True):
                if dir_[len(self.prefixList)].startswith('__attach') and\
                        dir_[len(self.prefixList)] not in attachmentDirs:
                    attachmentDirs.append(dir_[len(self.prefixList)])

            self._attachments = []

            for attachmentDir in attachmentDirs:
                try:
                    self._attachments.append(self.attachmentClass(self, attachmentDir))
                except NotImplementedError as e:
                    if self.attachmentErrorBehavior > constants.ATTACHMENT_ERROR_THROW:
                        logger.error('Error processing attachment at {}'.format(attachmentDir))
                        logger.exception(e)
                        self._attachments.append(UnsupportedAttachment(self, attachmentDir))
                    else:
                        raise
                except Exception as e:
                    if self.attachmentErrorBehavior == constants.ATTACHMENT_ERROR_BROKEN:
                        logger.error('Error processing attachment at {}'.format(attachmentDir))
                        logger.exception(e)
                        self._attachments.append(BrokenAttachment(self, attachmentDir))
                    else:
                        raise

            self.__attachmentsReady = True
            try:
                self.__waitingProperties
                if self.__attachmentsDelayed:
                    for attachment in self._attachments:
                        for prop in self.__waitingProperties:
                            attachment._registerNamedProperty(*prop)
            except:
                pass

            return self._attachments

    @property
    def attachmentsDelayed(self):
        """
        Returns True if the attachment initialization was delayed.
        """
        return self.__attachmentsDelayed

    @property
    def attachmentsReady(self):
        """
        Returns True if the attachments are ready to be used.
        """
        return self.__attachmentsReady

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
            self._body = self._getStringStream('__substg1.0_1000')
            if self._body:
                self._body = inputToString(self._body, 'utf-8')
                a = re.search('\n', self._body)
                if a is not None:
                    if re.search('\r\n', self._body) is not None:
                        self.__crlf = '\r\n'
            return self._body

    @property
    def cc(self):
        """
        Returns the cc field, if it exists.
        """
        return self._genRecipient('cc', 2)

    @property
    def compressedRtf(self):
        """
        Returns the compressed RTF stream, if it exists.
        """
        return self._ensureSet('_compressedRtf', '__substg1.0_10090102', False)

    @property
    def crlf(self):
        """
        Returns the value of self.__crlf, should you need it for whatever reason.
        """
        self.body
        return self.__crlf

    @property
    def date(self):
        """
        Returns the send date, if it exists.
        """
        try:
            return self._date
        except AttributeError:
            self._date = self._prop.date
            return self._date

    @property
    def header(self):
        """
        Returns the message header, if it exists. Otherwise it will generate one.
        """
        try:
            return self._header
        except AttributeError:
            headerText = self._getStringStream('__substg1.0_007D')
            if headerText is not None:
                self._header = EmailParser().parsestr(headerText)
                self._header['date'] = self.date
            else:
                logger.info('Header is empty or was not found. Header will be generated from other streams.')
                header = EmailParser().parsestr('')
                header.add_header('Date', self.date)
                header.add_header('From', self.sender)
                header.add_header('To', self.to)
                header.add_header('Cc', self.cc)
                header.add_header('Message-Id', self.messageId)
                # TODO find authentication results outside of header
                header.add_header('Authentication-Results', None)
                self._header = header
            return self._header

    @property
    def headerDict(self):
        """
        Returns a dictionary of the entries in the header
        """
        try:
            return self._headerDict
        except AttributeError:
            self._headerDict = dict(self.header._headers)
            try:
                self._headerDict.pop('Received')
            except KeyError:
                pass
            return self._headerDict

    @property
    def htmlBody(self):
        """
        Returns the html body, if it exists.
        """
        return self._ensureSet('_htmlBody', '__substg1.0_10130102', False)

    @property
    def inReplyTo(self):
        """
        Returns the message id that this message is in reply to.
        """
        return self._ensureSet('_in_reply_to', '__substg1.0_1042')

    @property
    def isRead(self):
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
    def recipients(self):
        """
        Returns a list of all recipients.
        """
        try:
            return self._recipients
        except AttributeError:
            # Get the recipients
            recipientDirs = []

            for dir_ in self.listDir():
                if dir_[len(self.prefixList)].startswith('__recip') and\
                        dir_[len(self.prefixList)] not in recipientDirs:
                    recipientDirs.append(dir_[len(self.prefixList)])

            self._recipients = []

            for recipientDir in recipientDirs:
                self._recipients.append(Recipient(recipientDir, self))

            return self._recipients

    @property
    def rtfBody(self):
        """
        Returns the decompressed Rtf body from the message.
        """
        return compressed_rtf.decompress(self.compressedRtf)

    @property
    def sender(self):
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
