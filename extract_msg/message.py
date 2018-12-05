import copy
import email.utils
import json
import olefile
import os
import re
from email.parser import Parser as EmailParser
from extract_msg import constants
from extract_msg.attachment import Attachment
from extract_msg.debug import debug
from extract_msg.properties import Properties
from extract_msg.recipient import Recipient
from extract_msg.utils import addNumToDir, encode, has_len, stri, windowsUnicode, xstr
from imapclient.imapclient import decode_utf7



class Message(olefile.OleFileIO):
    """
    Parser for Microsoft Outlook message files.
    """

    def __init__(self, path, prefix='', attachmentClass=Attachment, filename=None):
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
        """
        # WARNING DO NOT MANUALLY MODIFY PREFIX. Let the program set it.
        if debug:
            # DEBUG
            print('DEBUG: prefix: {}'.format(prefix))
        self.__path = path
        self.__attachmentClass = attachmentClass
        olefile.OleFileIO.__init__(self, path)
        prefixl = []
        if prefix != '':
            if not isinstance(prefix, stri):
                try:
                    prefix = '/'.join(prefix)
                except:
                    raise TypeError('Invalid prefix type: ' + str(type(prefix)) +
                        '\n(This was probably caused by you setting it manually).')
            prefix = prefix.replace('\\', '/')
            g = prefix.split("/")
            if g[-1] == '':
                g.pop()
            prefixl = g
            if prefix[-1] != '/':
                prefix += '/'
            filename = self._getStringStream(prefixl[:-1] + ['__substg1.0_3001'], prefix=False)
        self.__prefix = prefix
        self.__prefixList = prefixl
        if filename != None:
            self.filename = filename
        elif has_len(path):
            if len(path) < 1536:
                self.filename = path
            else:
                self.filename = None
        else:
            self.filename = None

        # Initialize properties in the order that is least likely to cause bugs.
        # TODO have each function check for initialization of needed data so these
        # lines will be unnecessary.
        self.mainProperties
        self.header
        self.recipients
        self.attachments
        self.to
        self.cc
        self.sender
        self.date
        self.__crlf = '\n'  # This variable keeps track of what the new line character should be
        self.body

    def listDir(self, streams=True, storages=False):
        """
        Replacement for OleFileIO.listdir that runs at the current prefix directory.
        """
        temp = self.listdir(streams, storages)
        if self.__prefix == '':
            return temp
        prefix = self.__prefix.split('/')
        if prefix[-1] == '':
            prefix.pop()
        out = []
        for x in temp:
            good = True
            if len(x) <= len(prefix):
                good = False
            if good:
                for y in range(len(prefix)):
                    if x[y] != prefix[y]:
                        good = False
            if good:
                out.append(x)
        return out

    def Exists(self, inp):
        """
        Checks if :param inp: exists in the msg file.
        """
        if isinstance(inp, list):
            inp = self.__prefixList + inp
        else:
            inp = self.__prefix + inp
        return self.exists(inp)

    def _getStream(self, filename, prefix=True):
        if isinstance(filename, list):
            filename = '/'.join(filename)
        if prefix:
            filename = self.__prefix + filename
        if self.exists(filename):
            stream = self.openstream(filename)
            return stream.read()
        else:
            return None

    def _getStringStream(self, filename, prefer='unicode', prefix=True):
        """
        Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then :param prefer: specifies which will be
        returned.
        """

        if isinstance(filename, list):
            # Join with slashes to make it easier to append the type
            filename = '/'.join(filename)

        asciiVersion = self._getStream(filename + '001E', prefix)
        unicodeVersion = windowsUnicode(self._getStream(filename + '001F', prefix))
        if debug:
            # DEBUG
            print('DEBUG: _getStringSteam called for {}. Ascii version found: {}. Unicode version found: {}.'.format(
                filename, asciiVersion != None, unicodeVersion != None))
        if asciiVersion is None:
            return unicodeVersion
        elif unicodeVersion is None:
            return asciiVersion
        else:
            if prefer == 'unicode':
                return unicodeVersion
            else:
                return asciiVersion

    @property
    def path(self):
        """
        Returns the message path if generated from a file,
        otherwise returns the data used to generate the
        Message instance.
        """
        return self.__path

    @property
    def prefix(self):
        """
        Returns the prefix of the Message instance.
        Intended for developer use.
        """
        return self.__prefix

    @property
    def prefixList(self):
        """
        Returns the prefix list of the Message instance.
        Intended for developer use.
        """
        return copy.deepcopy(self.__prefixList)

    @property
    def subject(self):
        """
        Returns the message subject, if it exists.
        """
        try:
            return self._subject
        except AttributeError:
            self._subject = encode(self._getStringStream('__substg1.0_0037'))
            return self._subject

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
                header = EmailParser().parsestr('')
                header.add_header('Date', self.date)
                header.add_header('From', self.sender)
                header.add_header('To', self.to)
                header.add_header('Cc', self.cc)
                header.add_header('Message-Id', self.message_id)
                #TODO find authentication results outside of header
                header.add_header('Authenitcation-Results', None)

                self._header = header
            return self._header

    def headerInit(self):
        """
        Checks whether the header has been initialized.
        """
        try:
            self._header
            return True
        except AttributeError:
            return False

    @property
    def mainProperties(self):
        """
        Returns the Properties instance used by the Message instance.
        """
        try:
            return self._prop
        except AttributeError:
            self._prop = Properties(self._getStream('__properties_version1.0'),
                                    constants.TYPE_MESSAGE if self.__prefix == '' else constants.TYPE_MESSAGE_EMBED)
            return self._prop

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
    def parsedDate(self):
        return email.utils.parsedate(self.date)

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
    def to(self):
        """
        Returns the to field, if it exists.
        """
        try:
            return self._to
        except AttributeError:
            # Check header first
            headerResult = None
            if self.headerInit():
                headerResult = self.header['to']
            if headerResult is not None:
                self._to = headerResult
            else:
                f = []
                for x in self.recipients:
                    if x.type & 0x0000000f == 1:
                        f.append(x.formatted)
                if len(f) > 0:
                    st = f[0]
                    if len(f) > 1:
                        for x in range(1, len(f)):
                            st += '; {0}'.format(f[x])
                    self._to = st
                else:
                    self._to = None
            return self._to

    @property
    def compressedRtf(self):
        """
        Returns the compressed RTF stream, if it exists.
        """
        try:
            return self._compressedRtf
        except AttributeError:
            self._compressedRtf = self._getStream('__substg1.0_10090102')
            return self._compressedRtf

    @property
    def htmlBody(self):
        """
        Returns the html body, if it exists.
        """
        try:
            return self._htmlBody
        except AttributeError:
            self._htmlBody = self._getStream('__substg1.0_10130102')
            return self._htmlBody

    @property
    def cc(self):
        """
        Returns the cc field, if it exists.
        """
        try:
            return self._cc
        except AttributeError:
            # Check header first
            headerResult = None
            if self.headerInit():
                headerResult = self.header['cc']
            if headerResult is not None:
                self._cc = headerResult
            else:
                f = []
                for x in self.recipients:
                    if x.type & 0x0000000f == 2:
                        f.append(x.formatted)
                if len(f) > 0:
                    st = f[0]
                    if len(f) > 1:
                        for x in range(1, len(f)):
                            st += '; {0}'.format(f[x])
                    self._cc = st
                else:
                    self._cc = None
            return self._cc

    @property
    def message_id(self):
        try:
            return self._message_id
        except AttributeError:
            if self.headerInit():
                self._message_idself._header['message-id']
            else:
                self._message_id = self._getStringStream('__substg1.0_1035')
            return self._message_id

    @property
    def reply_to(self):
        try:
            return self._reply_to
        except AttributeError:
            self._reply_to = self._getStringStream('__substg1.0_1042')
            return self._reply_to

    @property
    def body(self):
        """
        Returns the message body, if it exists.
        """
        try:
            return self._body
        except AttributeError:
            self._body = encode(self._getStringStream('__substg1.0_1000'))
            a = re.search('\n', self._body)
            if a != None:
                if re.search('\r\n', self._body) != None:
                    self.__crlf = '\r\n'
            return self._body

    @property
    def crlf(self):
        """
        Returns the value of self.__crlf, should you need it for whatever reason.
        """
        self.body
        return self.__crlf

    @property
    def attachmentClass(self):
        """
        Returns the Attachment class being used, should you need to use it externally for whatever reason.
        """
        return self.__attachmentClass

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

            for dir_ in self.listDir():
                if dir_[len(self.__prefixList)].startswith('__attach') and dir_[
                    len(self.__prefixList)] not in attachmentDirs:
                    attachmentDirs.append(dir_[len(self.__prefixList)])

            self._attachments = []

            for attachmentDir in attachmentDirs:
                self._attachments.append(self.__attachmentClass(self, attachmentDir))

            return self._attachments

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
                if dir_[len(self.__prefixList)].startswith('__recip') and dir_[
                    len(self.__prefixList)] not in recipientDirs:
                    recipientDirs.append(dir_[len(self.__prefixList)])

            self._recipients = []

            for recipientDir in recipientDirs:
                self._recipients.append(Recipient(recipientDir.split('#')[-1], self))

            return self._recipients

    def save(self, toJson=False, useFileName=False, raw=False, ContentId=False, customPath=None, customFilename=None):
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
        if customFilename != None and customFilename != '':
            dirName = customFilename
        else:
            if useFileName:
                # strip out the extension
                if self.filename != None:
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

        oldDir = os.getcwd()
        try:
            os.chdir(dirName)

            # Save the message body
            fext = 'json' if toJson else 'text'
            f = open('message.' + fext, 'w')
            # From, to , cc, subject, date

            attachmentNames = []
            # Save the attachments
            for attachment in self.attachments:
                attachmentNames.append(attachment.save(ContentId))

            if toJson:

                emailObj = {'from': xstr(self.sender),
                            'to': xstr(self.to),
                            'cc': xstr(self.cc),
                            'subject': xstr(self.subject),
                            'date': xstr(self.date),
                            'attachments': attachmentNames,
                            'body': decode_utf7(self.body)}

                f.write(json.dumps(emailObj, ensure_ascii=True))
            else:
                f.write('From: ' + xstr(self.sender) + self.__crlf)
                f.write('To: ' + xstr(self.to) + self.__crlf)
                f.write('CC: ' + xstr(self.cc) + self.__crlf)
                f.write('Subject: ' + xstr(self.subject) + self.__crlf)
                f.write('Date: ' + xstr(self.date) + self.__crlf)
                f.write('-----------------' + self.__crlf + self.__crlf)
                f.write(self.body)

            f.close()

        except Exception as e:
            self.saveRaw()
            raise

        finally:
            # Return to previous directory
            os.chdir(oldDir)

    def saveRaw(self):
        # Create a 'raw' folder
        oldDir = os.getcwd()
        try:
            rawDir = 'raw'
            os.makedirs(rawDir)
            os.chdir(rawDir)
            sysRawDir = os.getcwd()

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

    def dump(self):
        """
        Prints out a summary of the message
        """
        print('Message')
        print('Subject:', self.subject)
        print('Date:', self.date)
        print('Body:')
        print(self.body)

    def debug(self):
        for dir_ in self.listDir():
            if dir_[-1].endswith('001E') or dir_[-1].endswith('001F'):
                print('Directory: ' + str(dir_[:-1]))
                print('Contents: {}'.format(self._getStream(dir_)))

    def save_attachments(self, contentId=False, json=False, useFileName=False, raw=False, customPath=None):
        """
        Saves only attachments in the same folder.
        """
        for attachment in self.attachments:
            attachment.save(contentId, json, useFileName, raw, customPath)
