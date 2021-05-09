import json
import logging

from imapclient.imapclient import decode_utf7

from extract_msg import constants
from extract_msg.attachment import Attachment
from extract_msg.compat import os_ as os
from extract_msg.exceptions import DataNotFoundError, IncompatibleOptionsError
from extract_msg.message_base import MessageBase
from extract_msg.utils import addNumToDir, inputToBytes, inputToString


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

    def getJson(self):
        """
        Returns the JSON representation of the Message.
        """


    def save(self, **kwargs):
        """
        Saves the message body and attachments found in the message.

        The body and attachments are stored in a folder in the current running
        directory unless :param customPath: has been specified. The name of the
        folder will be determined by 3 factors.
           * If :param customFilename: has been set, the value provided for that
             will be used.
           * If :param useMsgFilename: has been set, the name of the file used
             to create the Message instance will be used.
           * If the file name has not been provided or :param useMsgFilename:
             has not been set, the name of the folder will be created using the
             `defaultFolderName` property.

        :param useMsgFilename: is passed recursively to all save functions, but
        is used exclusively by embedded msg files.

        There are several parameters used to determine how the message will be
        saved. By default, the message will be saved as plain text. Setting one
        of the following parameters to True will change that:
           * :param html: will try to output the message in HTML format.
           * :param json: will output the message in JSON format.
           * :param raw: will output the message in a raw format.
           * :param rtf: will output the message in RTF format.

        Usage of more than one formatting parameter will raise an exception.

        Using HTML or RTF will raise an exception if they could not be retrieved
        unless you have :param allowFallback: set to True. Fallback will go in
        this order, starting at the top most format that is set:
           * HTML
           * RTF
           * Plain text
        """

        # Move keyword arguments into variables.
        _json = kwargs.get('json', False)
        html = kwargs.get('html', False)
        rtf = kwargs.get('rtf', False)
        raw = kwargs.get('raw', False)

        # Variables involved in the save location.
        path = os.path.abspath(kwargs.get('customPath', os.getcwdu())).replace('\\', '/')
        customFilename = kwargs.get('customFilename')
        useMsgFilename = kwargs.get('useMsgFilename', False)

        # Reset this for sub save calls.
        kwargs['customFilename'] = None

        # Check if incompatible options have been provided in any way.
        if _json + html + rtf + raw > 1:
            raise IncompatibleOptionsError('Only one of the following options may be used at a time: toJson, raw, html, rtf')

        # Get the type of line endings.
        crlf = inputToBytes(self.crlf, 'utf-8')

        # Prepare the path.
        path += '/' if path[-1] != '/' else ''

        if customFilename:
            # First we need to validate it. If there are invalid characters, this will detect it.
            if constants.RE_INVALID_PATH_CHARACTERS.search(customFilename):
                raise ValueError('Invalid character found in customFilename. Must not contain any of the following characters: \\/:*?"<>|')
            path += customFilename
        elif useMsgFilename:
            if not self.filename:
                raise ValueError(':param useMsgFilename: is only available if you are using an msg file on the disk or have provided a filename.')
            path += self.filename
        else:
            path += self.defaultFolderName

        # Prepare the path one last time.
        path += '/' if path[-1] != '/' else ''

        # Update the kwargs.
        kwargs['customPath'] = path

        # Create the folders.
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

        if raw:
            self.saveRaw(path)
            return self


        try:
            # Save the attachments
            attachmentNames = [attachment.save(**kwargs) for attachment in self.attachments]

            # Save the message body
            fext = 'json' if _json else 'txt'

            useHtml = False
            useRtf = False
            if html:
               if self.htmlBody:
                   useHtml = True
                   fext = 'html'
            elif not allowFallback:
               raise DataNotFoundError('Could not find the htmlBody')

            if rtf or (html and not useHtml):
               if self.rtfBody:
                   useRtf = True
                   fext = 'rtf'
            elif not allowFallback:
               raise DataNotFoundError('Could not find the rtfBody')
            with open(path + 'message.' + fext, 'wb') as f:
                if _json:
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

        except Exception:
            self.saveRaw(path)
            raise

        finally:
            # Return to previous directory
            os.chdir(oldDir)

        # Return the instance so that functions can easily be chained.
        return self
