import json
import logging
import zipfile

from imapclient.imapclient import decode_utf7

from . import constants
from .attachment import Attachment
from .compat import os_ as os
from .exceptions import DataNotFoundError, IncompatibleOptionsError
from .message_base import MessageBase
from .utils import addNumToDir, inputToBytes, inputToString, makeDirs, prepareFilename


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

class Message(MessageBase):
    """
    Parser for Microsoft Outlook message files.
    """

    def __init__(self, path, prefix = '', attachmentClass = Attachment, filename = None, delayAttachments = False, overrideEncoding = None, attachmentErrorBehavior = constants.ATTACHMENT_ERROR_THROW, recipientSeparator = ';'):
        MessageBase.__init__(self, path, prefix, attachmentClass, filename, delayAttachments, overrideEncoding, attachmentErrorBehavior, recipientSeparator)

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
        return json.dumps({
            'from': inputToString(self.sender, 'utf-8'),
            'to': inputToString(self.to, 'utf-8'),
            'cc': inputToString(self.cc, 'utf-8'),
            'subject': inputToString(self.subject, 'utf-8'),
            'date': inputToString(self.date, 'utf-8'),
            'body': decode_utf7(self.body)
        })

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
           * :param maxNameLength: will force all file names to be shortened
             to fit in the space (with the extension included in the length). If
             a number is added to the directory that will not be included in the
             length, so it is recommended to plan for up to 5 characters extra
             to be a part of the name.

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

        If you want to save the contents into a ZipFile or similar object,
        either pass a path to where you want to create one or pass an instance
        to :param zip:. If :param zip: is an instance, :param customPath: will
        refer to a location inside the zip file.

        If you want to save the header, should it be found, set
        :param saveHeader: to true.
        """

        # Move keyword arguments into variables.
        _json = kwargs.get('json', False)
        html = kwargs.get('html', False)
        rtf = kwargs.get('rtf', False)
        raw = kwargs.get('raw', False)
        allowFallback = kwargs.get('allowFallback', False)
        zip = kwargs.get('zip')
        maxNameLength = kwargs.get('maxNameLength')

        # Variables involved in the save location.
        customFilename = kwargs.get('customFilename')
        useMsgFilename = kwargs.get('useMsgFilename', False)
        #maxPathLength = kwargs.get('maxPathLength', 255)

        # ZipFile handling.
        if zip:
            # `raw` and `zip` are incompatible.
            if raw:
                raise IncompatibleOptionsError('The options `raw` and `zip` are incompatible.')
            # If we are doing a zip file, first check that we have been given a path.
            if isinstance(zip, constants.STRING):
                # If we have a path then we use the zip file.
                zip = zipfile.ZipFile(zip, 'a', zipfile.ZIP_DEFLATED)
                kwargs['zip'] = zip
                createdZip = True
            else:
                createdZip = False
            # Path needs to be done in a special way if we are in a zip file.
            path = kwargs.get('customPath', '').replace('\\', '/')
            path += '/' if path and path[-1] != '/' else ''
            # Set the open command to be that of the zip file.
            _open = zip.open
            # Zip files use w for writing in binary.
            mode = 'w'
        else:
            path = os.path.abspath(kwargs.get('customPath', os.getcwdu())).replace('\\', '/')
            # Prepare the path.
            path += '/' if path[-1] != '/' else ''
            mode = 'wb'
            _open = open

        # Reset this for sub save calls.
        kwargs['customFilename'] = None

        # Check if incompatible options have been provided in any way.
        if _json + html + rtf + raw > 1:
            raise IncompatibleOptionsError('Only one of the following options may be used at a time: toJson, raw, html, rtf')

        # Get the type of line endings.
        crlf = inputToBytes(self.crlf, 'utf-8')

        # TODO: insert code here that will handle checking all of the msg files to see if the path with overflow.

        if customFilename:
            # First we need to validate it. If there are invalid characters, this will detect it.
            if constants.RE_INVALID_FILENAME_CHARACTERS.search(customFilename):
                raise ValueError('Invalid character found in customFilename. Must not contain any of the following characters: \\/:*?"<>|')
            path += customFilename
        elif useMsgFilename:
            if not self.filename:
                raise ValueError(':param useMsgFilename: is only available if you are using an msg file on the disk or have provided a filename.')
            # Get the actual name of the file.
            filename = os.path.split(self.filename)[1]
            # Remove the extensions.
            filename = os.path.splitext(filename)[0]
            # Prepare the filename by removing any special characters.
            filename = prepareFilename(filename)
            # Shorted the filename.
            filename = filename[:maxNameLength]
            # Check to make sure we actually have a filename to use.
            if not filename:
                raise ValueError('Invalid filename found in self.filename: "{}"'.format(self.filename))

            # Add the file name to the path.
            path += filename
        else:
            path += self.defaultFolderName[:maxNameLength]

        # Create the folders.
        if not zip:
            try:
                makeDirs(path)
            except Exception:
                newDirName = addNumToDir(path)
                if newDirName:
                    path = newDirName
                else:
                    raise Exception(
                        "Failed to create directory '%s'. Does it already exist?" %
                        path
                    )

        # Prepare the path one last time.
        path += '/' if path[-1] != '/' else ''

        # Update the kwargs.
        kwargs['customPath'] = path

        if raw:
            self.saveRaw(path)
            return self

        try:
            # Save the attachments.
            attachmentNames = [attachment.save(**kwargs) for attachment in self.attachments]

            # Determine the extension to use for the body.
            fext = 'json' if _json else 'txt'

            # Check whether we should be using HTML or RTF.
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

            with _open(path + 'message.' + fext, mode) as f:
                if _json:
                    emailObj = json.loads(self.getJson())
                    emailObj['attachments'] = attachmentNames

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
            if not zip:
                self.saveRaw(path)
            raise
        finally:
            # Close the ZipFile if this function created it.
            if zip and createdZip:
                zip.close()

        # Return the instance so that functions can easily be chained.
        return self
