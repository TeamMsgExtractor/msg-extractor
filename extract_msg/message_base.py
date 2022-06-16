import base64
import email.utils
import html
import logging
import re

import bs4
import chardet
import compressed_rtf
import RTFDE

from . import constants
from .enums import DeencapType, RecipientType
from .exceptions import DeencapMalformedData, DeencapNotEncapsulated
from .msg import MSGFile
from .recipient import Recipient
from .utils import inputToString, prepareFilename
from email.parser import Parser as EmailParser


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class MessageBase(MSGFile):
    """
    Base class for Message like msg files.
    """

    def __init__(self, path, **kwargs):
        """
        :param path: path to the msg file in the system or is the raw msg file.
        :param prefix: used for extracting embeded msg files
            inside the main one. Do not set manually unless
            you know what you are doing.
        :param parentMsg: Used for syncronizing named properties instances. Do
            not set this unless you know what you are doing.
        :param attachmentClass: Optional, the class the Message object
            will use for attachments. You probably should
            not change this value unless you know what you
            are doing.
        :param filename: Optional, the filename to be used by default when
            saving.
        :param delayAttachments: Optional, delays the initialization of
            attachments until the user attempts to retrieve them. Allows MSG
            files with bad attachments to be initialized so the other data can
            be retrieved.
        :param overrideEncoding: Optional, an encoding to use instead of the one
            specified by the msg file. Do not report encoding errors caused by
            this.
        :param attachmentErrorBehavior: Optional, the behavior to use in the
            event of an error when parsing the attachments.
        :param recipientSeparator: Optional, separator string to use between
            recipients.
        :param ignoreRtfDeErrors: Optional, specifies that any errors that occur
            from the usage of RTFDE should be ignored (default: False).
        :param deencapsulationFunc: Optional, if specified must be a callable
            that will override the way that HTML/text is deencapsulated from the
            RTF body. This function must take exactly 2 arguments, the first
            being the RTF body from the message and the second being an instance
            of the enum DeencapType that will tell the function what type of
            body is desired. The function should return a string for plain text
            and bytes for HTML. If any problems occur, the function *must*
            either return None or raise one of the appropriate functions from
            extract_msg.exceptions. All other functions must be handled
            internally or they will continue. The original deencapsulation
            method will not run if this is set.
        """
        super().__init__(path, **kwargs)
        recipientSeparator = ';'
        self.__recipientSeparator = kwargs.get('recipientSeparator', ';')
        self.__ignoreRtfDeErrors = kwargs.get('ignoreRtfDeErrors', False)
        self.__deencap = kwargs.get('deencapsulationFunc')
        # Initialize properties in the order that is least likely to cause bugs.
        # TODO have each function check for initialization of needed data so these
        # lines will be unnecessary.
        self.mainProperties
        self.header
        self.recipients

        self.to
        self.cc
        self.sender
        self.date
        # This variable keeps track of what the new line character should be.
        self.__crlf = '\n'
        try:
            self.body
        except Exception as e:
            # Prevent an error in the body from preventing opening.
            logger.exception('Critical error accessing the body. File opened but accessing the body will throw an exception.')
        self.named
        self.namedProperties

    def _genRecipient(self, recipientType, recipientInt : RecipientType):
        """
        Returns the specified recipient field.
        """
        private = '_' + recipientType
        recipientInt = RecipientType(recipientInt)
        try:
            return getattr(self, private)
        except AttributeError:
            value = None
            # Check header first.
            if self.headerInit():
                value = self.header[recipientType]
                if value:
                    value = value.replace(',', self.__recipientSeparator)

            # If the header had a blank field or didn't have the field, generate it manually.
            if not value:
                # Check if the header has initialized.
                if self.headerInit():
                    logger.info(f'Header found, but "{recipientType}" is not included. Will be generated from other streams.')

                # Get a list of the recipients of the specified type.
                foundRecipients = tuple(recipient.formatted for recipient in self.recipients if recipient.type == recipientInt)

                # If we found recipients, join them with the recipient separator and a space.
                if len(foundRecipients) > 0:
                    value = (self.__recipientSeparator + ' ').join(foundRecipients)

            # Code to fix the formatting so it's all a single line. This allows the user to format it themself if they want.
            # This should probably be redone to use re or something, but I can do that later. This shouldn't be a huge problem for now.
            if value:
                value = value.replace(' \r\n\t', ' ').replace('\r\n\t ', ' ').replace('\r\n\t', ' ')
                value = value.replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
                while value.find('  ') != -1:
                    value = value.replace('  ', ' ')

            # Set the field in the class.
            setattr(self, private, value)

            return value

    def deencapsulateBody(self, rtfBody : bytes, bodyType : DeencapType):
        """
        A function to deencapsulate the specified body from the rtfBody. Returns
        a string for plain text and bytes for HTML. If specified, uses the
        deencapsulation override function. Returns None if nothing could be
        deencapsulated.

        If you want to change the deencapsulation behaviour in a base class,
        simply override this function.
        """
        if rtfBody:
            bodyType = DeencapType(bodyType)
            if bodyType == DeencapType.PLAIN:
                if self.__deencap:
                    try:
                        return self.__deencap(rtfBody, DeencapType.PLAIN)
                    except DeencapMalformedData:
                        logger.exception('Custom deencapsulation function reported encapsulated data was malformed.')
                    except DeencapNotEncapsulated:
                        logger.exception('Custom deencapsulation function reported data is not encapsulated.')
                else:
                    if self.deencapsulatedRtf and self.deencapsulatedRtf.content_type == 'text':
                        return self.deencapsulatedRtf.text
            else:
                if self.__deencap:
                    try:
                        return self.__deencap(rtfBody, DeencapType.HTML)
                    except DeencapMalformedData:
                        logger.exception('Custom deencapsulation function reported encapsulated data was malformed.')
                    except DeencapNotEncapsulated:
                        logger.exception('Custom deencapsulation function reported data is not encapsulated.')
                else:
                    if self.deencapsulatedRtf and self.deencapsulatedRtf.content_type == 'html':
                        return self.deencapsulatedRtf.html.encode('utf-8')

            if bodyType == DeencapType.PLAIN:
                logger.info('Could not deencapsulate plain text from RTF body.')
            else:
                logger.info('Could not deencapsulate HTML from RTF body.')
        else:
            logger.info('No RTF body to deencapsulate from.')
        return None

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
    def bcc(self):
        """
        Returns the bcc field, if it exists.
        """
        return self._genRecipient('bcc', RecipientType.BCC)

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
            else:
                # If the body doesn't exist, see if we can get it from the RTF
                # body.
                if self.rtfBody:
                    self._body = self.deencapsulateBody(self.rtfBody, DeencapType.PLAIN)

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
        return self._genRecipient('cc', RecipientType.CC)

    @property
    def compressedRtf(self):
        """
        Returns the compressed RTF stream, if it exists.
        """
        return self._ensureSet('_compressedRtf', '__substg1.0_10090102', False)

    @property
    def crlf(self):
        """
        Returns the value of self.__crlf, should you need it for whatever
        reason.
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
    def deencapsulatedRtf(self) -> RTFDE.DeEncapsulator:
        """
        Returns the instance of the deencapsulated RTF body. If there is no RTF
        body or the body is not encasulated, returns None.
        """
        try:
            return self._deencapsultor
        except AttributeError:
            if self.rtfBody:
                # If there is an RTF body, we try to deencapsulate it.
                body = self.rtfBody
                # Sometimes you get MSG files whose RTF body has stuff
                # *after* the body, and RTFDE can't handle that. Here is
                # how we compensate.
                while body and body[-1] != 125:
                    body = body[:-1]

                try:
                    try:
                        self._deencapsultor = RTFDE.DeEncapsulator(body)
                    except UnicodeDecodeError:
                        # There is a known issue that bytes are not well decoded
                        # by RTFDE right now, so let's see if we can't manually
                        # decode it and see if that will work.
                        #
                        # There is also the fact that it is decoded *at all*
                        # before binary data is stripped out. This data should
                        # almost certainly be stripped out, so let's log it and
                        # then log if we removed any of them before trying this.
                        logger.warn(f'RTFDE failed to decode rtfBody for message with subject "{self.subject}". Attempting to cut out unnecessary data and override decoding.')

                        match = constants.RE_BIN.search(body)
                        # Because we are going to be actively removing things,
                        # we want to search the entire thing over again.
                        while match:
                            logger.info(f'Found match to bin data starting at location {match.start()}. Replacing with nothing.')
                            length = int(match.group(1))
                            # Extract the entire binary section and replace it.
                            body = body.replace(body[match.start():match.end() + length], b'', 1)
                            match = constants.RE_BIN.search(body)

                        self._deencapsultor = RTFDE.DeEncapsulator(body.decode(chardet.detect(body)['encoding']))
                    self._deencapsultor.deencapsulate()
                except RTFDE.exceptions.NotEncapsulatedRtf as e:
                    logger.debug('RTF body is not encapsulated.')
                    self._deencapsultor = None
                except RTFDE.exceptions.MalformedEncapsulatedRtf as _e:
                    logger.info('RTF body contains malformed encapsulated content.')
                    self._deencapsultor = None
                except Exception:
                    # If we are just ignoring the errors, log it then set to
                    # None. Otherwise, continue the exception.
                    if not self.__ignoreRtfDeErrors:
                        raise
                    logger.exception('Unhandled error happened while using RTFDE. You have choosen to ignore these errors.')
                    self._deencapsultor = None
            else:
                self._deencapsultor = None
            return self._deencapsultor

    @property
    def defaultFolderName(self) -> str:
        """
        Generates the default name of the save folder.
        """
        try:
            return self._defaultFolderName
        except AttributeError:
            d = self.parsedDate

            dirName = '{0:02d}-{1:02d}-{2:02d}_{3:02d}{4:02d}'.format(*d) if d else 'UnknownDate'
            dirName += ' ' + (prepareFilename(self.subject) if self.subject else '[No subject]')

            self._defaultFolderName = dirName
            return dirName

    @property
    def header(self):
        """
        Returns the message header, if it exists. Otherwise it will generate
        one.
        """
        try:
            return self._header
        except AttributeError:
            headerText = self._getStringStream('__substg1.0_007D')
            if headerText:
                self._header = EmailParser().parsestr(headerText)
                self._header['date'] = self.date
            else:
                logger.info('Header is empty or was not found. Header will be generated from other streams.')
                header = EmailParser().parsestr('')
                header.add_header('Date', self.date)
                header.add_header('From', self.sender)
                header.add_header('To', self.to)
                header.add_header('Cc', self.cc)
                header.add_header('Bcc', self.bcc)
                header.add_header('Message-Id', self.messageId)
                # TODO find authentication results outside of header
                header.add_header('Authentication-Results', None)
                self._header = header
            return self._header

    @property
    def headerDict(self) -> dict:
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
            elif self.rtfBody:
                logger.info('HTML body was not found, attempting to generate from RTF.')
                self._htmlBody = self.deencapsulateBody(self.rtfBody, DeencapType.HTML)
            # This is it's own if statement so we can ensure it will generate
            # even if there is an rtfBody, in the event it doesn't have HTML.
            if not self._htmlBody and self.body:
                # Convert the plain text body to html.
                logger.info('HTML body was not found, attempting to generate from plain text body.')
                correctedBody = html.escape(self.body).replace('\r', '').replace('\n', '</br>')
                self._htmlBody = f'<html><body>{correctedBody}</body></head>'.encode('utf-8')

            if not self._htmlBody:
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
        return self._genRecipient('to', RecipientType.TO)
