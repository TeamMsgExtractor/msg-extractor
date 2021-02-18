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
from extract_msg.exceptions import DataNotFoundError, IncompatibleOptionsError
from extract_msg.message_base import MessageBase
from extract_msg.recipient import Recipient
from extract_msg.utils import addNumToDir, inputToBytes, inputToString, prepareFilename



logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

class Message(MessageBase):
    """
    Parser for Microsoft Outlook message files.
    """

    def __init__(self, path, prefix = '', attachmentClass = Attachment, filename = None, delayAttachments = False, overrideEncoding = None, attachmentErrorBehavior = constants.ATTACHMENT_ERROR_THROW):
        MessageBase.__init__(self, path, prefix, attachmentClass, filename, delayAttachments, overrideEncoding, attachmentErrorBehavior)

    def dump(self):
        """
        Prints out a summary of the message
        """
        print('Message')
        print('Subject:', self.subject)
        print('Date:', self.date)
        print('Body:')
        print(self.body)

    def save(self, toJson = False, useFileName = False, raw = False, ContentId = False, customPath = None, customFilename = None):#, html = False, rtf = False, allowFallback = False):
        """
        Saves the message body and attachments found in the message. The body and
        attachments are stored in a folder. Setting useFileName to true will mean that
        the filename is used as the name of the folder; otherwise, the message's date
        and subject are used as the folder name.
        Here is the absolute order of prioity for the name of the folder:
            1. customFilename
            2. self.filename if useFileName
            3. {date} {subject}
        """
        #There are several parameters used to determine how the message will be saved.
        #By default, the message will be saved as plain text. Setting one of the
        #following parameters to True will change that:
        #    * :param html: will try to output the message in HTML format.
        #    * :param json: will output the message in JSON format.
        #    * :param raw: will output the message in a raw format.
        #    * :param rtf: will output the message in RTF format.
        #
        #Usage of more than one formatting parameter will raise an exception.
        #
        #Using HTML or RTF will raise an exception if they could not be retrieved
        #unless you have :param allowFallback: set to True. Fallback will go in this
        #order, starting at the top most format that is set:
        #    * HTML
        #    * RTF
        #    * Plain text
        #"""
        count = 1 if toJson else 0
        #count += 1 if html else 0
        #count += 1 if rtf else 0
        count += 1 if raw else 0

        if count > 1:
            raise IncompatibleOptionsError('Only one of the following options may be used at a time: toJson, raw, html, rtf')

        crlf = inputToBytes(self.crlf, 'utf-8')

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
                    subject = prepareFilename(self.subject)

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
                attachmentNames.append(attachment.save(ContentId, toJson, useFileName, raw))#, html = html, rtf = rtf, allowFallback = allowFallback))

            # Save the message body
            fext = 'json' if toJson else 'txt'

            useHtml = False
            useRtf = False
            #if html:
            #    if self.htmlBody is not None:
            #        useHtml = True
            #        fext = 'html'
            #elif not allowFallback:
            #    raise DataNotFoundError('Could not find the htmlBody')

            #if rtf or (html and not useHtml):
            #    if self.rtfBody is not None:
            #        useRtf = True
            #        fext = 'rtf'
            #elif not allowFallback:
            #    raise DataNotFoundError('Could not find the rtfBody')

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

        # Return the instance so that functions can easily be chained.
        return self

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
