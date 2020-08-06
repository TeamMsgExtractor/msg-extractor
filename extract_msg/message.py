import email.utils
import json
import logging
import re

import compressed_rtf
from imapclient.imapclient import decode_utf7

from email.parser import Parser as EmailParser
from extract_msg import constants
from extract_msg.attachment import Attachment
from extract_msg.compat import os_ as os
from extract_msg.msg import MSGFile
from extract_msg.recipient import Recipient
from extract_msg.utils import addNumToDir, inputToBytes, inputToString



logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

class Message(MSGFile):
    """
    Parser for Microsoft Outlook message files.
    """

    def __init__(self, path, prefix = '', attachmentClass = Attachment, filename = None, delayAttachments = False):
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
        """
        MSGFile.__init__(self, path, prefix, attachmentClass, filename)
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
            headerResult = None
            if self.headerInit():
                headerResult = self.header[recipientType]
            if headerResult is not None:
                setattr(self, private, headerResult)
            else:
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
                    setattr(self, private, st)
                else:
                    setattr(self, private, None)
            return getattr(self, private)

    def _registerNamedProperty(self, entry, _type, name = None):
        if self.attachmentsDelayed and not self.attachmentsReady:
            try:
                self.__waitingProperties
            except KeyError:
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

    def dump(self):
        """
        Prints out a summary of the message
        """
        print('Message')
        print('Subject:', self.subject)
        print('Date:', self.date)
        print('Body:')
        print(self.body)

    def headerInit(self):
        """
        Checks whether the header has been initialized.
        """
        try:
            self._header
            return True
        except AttributeError:
            return False

    def save(self, toJson = False, useFileName = False, raw = False, ContentId = False, customPath = None, customFilename = None, html = False, rtf = False):
        """
        Saves the message body and attachments found in the message. Setting toJson
        to true will output the message body as JSON-formatted text. The body and
        attachments are stored in a folder. Setting useFileName to true will mean that
        the filename is used as the name of the folder; otherwise, the message's date
        and subject are used as the folder name.
        Here is the absolute order of prioity for the name of the folder:
            1. customFilename
            2. self.filename if useFileName
            3. {date} {subject}
        """
        count = 0
        count += 1 if toJson else 0
        count += 1 if html else 0
        count += 1 if rtf else 0
        count += 1 if raw else 0

        if count > 1:
            raise IncompatibleOptionsException('Only one of the following options may be used at a time: toJSon, raw, html, rtf')

        crlf = inputToBytes(self.__crlf, 'utf-8')

        if customFilename != None and customFilename != '':
            dirName = customFilename
        else:
            if useFileName:
                # strip out the extension
                if self.filename is not None:
                    dirName = self.filename.split('/').pop().split('.')[0]
                else:
                    ValueError(
                        'Filename must be specified, or path must have been an actual path, to save using filename')
            else:
                # Create a directory based on the date and subject of the message
                d = self.parsedDate
                if d is not None:
                    dirName = '{0:02d}-{1:02d}-{2:02d}_{3:02d}{4:02d}'.format(*d)
                else:
                    dirName = 'UnknownDate'

                if self.subject is None:
                    subject = '[No subject]'
                else:
                    subject = ''.join(i for i in self.subject if i not in r'\/:*?"<>|')

                dirName = dirName + ' ' + subject

        if customPath != None and customPath != '':
            if customPath[-1] != '/' or customPath[-1] != '\\':
                customPath += '/'
            dirName = customPath + dirName
        try:
            os.makedirs(dirName)
        except Exception:
            newDirName = addNumToDir(dirName)
            if newDirName is not None:
                dirName = newDirName
            else:
                raise Exception(
                    "Failed to create directory '%s'. Does it already exist?" %
                    dirName
                )

        oldDir = os.getcwdu()
        try:
            os.chdir(dirName)
            attachmentNames = []
            # Save the attachments
            for attachment in self.attachments:
                attachmentNames.append(attachment.save(ContentId, toJson, useFileName, raw, html = html, rtf = rtf))

            # Save the message body
            fext = 'json' if toJson else 'txt'

            useHtml = False
            useRtf = False
            #if html:
            #    if self.htmlBody is not None:
            #        useHtml = True
            #        fext = 'html'
            #elif rtf:
            #    if self.htmlBody is not None:
            #        useRtf = True
            #        fext = 'rtf'

            with open('message.' + fext, 'wb') as f:
                if toJson:
                    emailObj = {'from': inputToString(self.sender, 'utf-8'),
                                'to': inputToString(self.to, 'utf-8'),
                                'cc': inputToString(self.cc, 'utf-8'),
                                'subject': inputToString(self.subject, 'utf-8'),
                                'date': inputToString(self.date, 'utf-8'),
                                'attachments': attachmentNames,
                                'body': decode_utf7(self.body)}

                    f.write(inputToBytes(json.dumps(emailObj), 'utf-8'))
                else:
                    if useHtml:
                        # Do stuff
                        pass
                    elif useRtf:
                        # Do stuff
                        pass
                    else:
                        f.write(b'From: ' + inputToBytes(self.sender, 'utf-8') + crlf)
                        f.write(b'To: ' + inputToBytes(self.to, 'utf-8') + crlf)
                        f.write(b'CC: ' + inputToBytes(self.cc, 'utf-8') + crlf)
                        f.write(b'Subject: ' + inputToBytes(self.subject, 'utf-8') + crlf)
                        f.write(b'Date: ' + inputToBytes(self.date, 'utf-8') + crlf)
                        f.write(b'-----------------' + crlf + crlf)
                        f.write(inputToBytes(self.body, 'utf-8'))

        except Exception as e:
            self.saveRaw()
            raise

        finally:
            # Return to previous directory
            os.chdir(oldDir)

    def save_attachments(self, contentId = False, json = False, useFileName = False, raw = False, customPath = None):
        """
        Saves only attachments in the same folder.
        """
        for attachment in self.attachments:
            attachment.save(contentId, json, useFileName, raw, customPath)

    def saveRaw(self):
        # Create a 'raw' folder
        oldDir = os.getcwdu()
        try:
            rawDir = 'raw'
            os.makedirs(rawDir)
            os.chdir(rawDir)
            sysRawDir = os.getcwdu()

            # Loop through all the directories
            for dir_ in self.listdir():
                sysdir = '/'.join(dir_)
                code = dir_[-1][-8:]
                if code in constants.PROPERTIES:
                    sysdir = sysdir + ' - ' + constants.PROPERTIES[code]
                os.makedirs(sysdir)
                os.chdir(sysdir)

                # Generate appropriate filename
                if dir_[-1].endswith('001E'):
                    filename = 'contents.txt'
                else:
                    filename = 'contents'

                # Save contents of directory
                with open(filename, 'wb') as f:
                    f.write(self._getStream(dir_))

                # Return to base directory
                os.chdir(sysRawDir)

        finally:
            os.chdir(oldDir)

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
                self._attachments.append(self.attachmentClass(self, attachmentDir))
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
