import json
import logging
import os
import pathlib
import subprocess
import zipfile

import bs4

from imapclient.imapclient import decode_utf7

from . import constants
from .exceptions import DataNotFoundError, IncompatibleOptionsError, WKError
from .message_base import MessageBase
from .utils import addNumToDir, addNumToZipDir, createZipOpen, findWk, injectHtmlHeader, injectRtfHeader, inputToBytes, inputToString, prepareFilename


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class Message(MessageBase):
    """
    Parser for Microsoft Outlook message files.
    """

    def __init__(self, path, **kwargs):
        super().__init__(path, **kwargs)

    def dump(self) -> None:
        """
        Prints out a summary of the message
        """
        print('Message')
        print('Subject:', self.subject)
        print('Date:', self.date)
        print('Body:')
        print(self.body)

    def getJson(self) -> str:
        """
        Returns the JSON representation of the Message.
        """
        return json.dumps({
            'from': inputToString(self.sender, 'utf-8'),
            'to': inputToString(self.to, 'utf-8'),
            'cc': inputToString(self.cc, 'utf-8'),
            'bcc': inputToString(self.bcc, 'utf-8'),
            'subject': inputToString(self.subject, 'utf-8'),
            'date': inputToString(self.date, 'utf-8'),
            'body': decode_utf7(self.body),
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
             to be a part of the name. Default is 256.

        It should be noted that regardless of the value for maxNameLength, the
        name of the file containing the body will always have the name 'message'
        followed by the full extension.

        There are several parameters used to determine how the message will be
        saved. By default, the message will be saved as plain text. Setting one
        of the following parameters to True will change that:
           * :param html: will output the message in HTML format.
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
        to :param zip:. If :param zip: is set, :param customPath: will refer to
        a location inside the zip file.

        :param attachmentsOnly: Turns off saving the body and only saves the
            attachments when set.
        :param skipAttachments: Turns off saving attachments.
        :param charset: If the html is being prepared, the charset to use for
            the Content-Type meta tag to insert. This exists to ensure that
            something parsing the html can properly determine the encoding (as
            not having this tag can cause errors in some programs). Set this to
            `None` or an empty string to not insert the tag (Default: 'utf-8').
        :param kwargs: Used to allow kwargs expansion in the save function.
        :param preparedHtml: When set, prepares the HTML body for standalone
            usage, doing things like adding tags, injecting attachments, etc.
            This is useful for things like trying to convert the HTML body
            directly to PDF.
        :param pdf: Used to enable saving the body as a PDF file.
        :param wkPath: Used to manually specify the path of the wkhtmltopdf
            executable. If not specified, the function will try to find it.
            Useful if wkhtmltopdf is not on the path. If :param pdf: is False,
            this argument is ignored.
        :param wkOptions: Used to specify additional options to wkhtmltopdf.
            this must be a list or list-like object composed of strings and
            bytes.
        """
        # Move keyword arguments into variables.
        _json = kwargs.get('json', False)
        html = kwargs.get('html', False)
        rtf = kwargs.get('rtf', False)
        raw = kwargs.get('raw', False)
        pdf = kwargs.get('pdf', False)
        allowFallback = kwargs.get('allowFallback', False)
        _zip = kwargs.get('zip')
        maxNameLength = kwargs.get('maxNameLength', 256)

        # Variables involved in the save location.
        customFilename = kwargs.get('customFilename')
        useMsgFilename = kwargs.get('useMsgFilename', False)
        #maxPathLength = kwargs.get('maxPathLength', 255)

        # Track if we are only saving the attachments.
        attachOnly = kwargs.get('attachmentsOnly', False)
        # Track if we are skipping attachments.
        skipAttachments = kwargs.get('skipAttachments', False)

        if pdf:
            kwargs['preparedHtml'] = True

        # ZipFile handling.
        if _zip:
            # `raw` and `zip` are incompatible.
            if raw:
                raise IncompatibleOptionsError('The options `raw` and `zip` are incompatible.')
            # If we are doing a zip file, first check that we have been given a path.
            if isinstance(_zip, (str, pathlib.Path)):
                # If we have a path then we use the zip file.
                _zip = zipfile.ZipFile(_zip, 'a', zipfile.ZIP_DEFLATED)
                kwargs['zip'] = _zip
                createdZip = True
            else:
                createdZip = False
            # Path needs to be done in a special way if we are in a zip file.
            path = pathlib.Path(kwargs.get('customPath', ''))
            # Set the open command to be that of the zip file.
            _open = createZipOpen(_zip.open)
            # Zip files use w for writing in binary.
            mode = 'w'
        else:
            path = pathlib.Path(kwargs.get('customPath', '.')).absolute()
            mode = 'wb'
            _open = open

        # Reset this for sub save calls.
        kwargs['customFilename'] = None

        # Check if incompatible options have been provided in any way.
        if _json + html + rtf + raw + attachOnly + pdf > 1:
            raise IncompatibleOptionsError('Only one of the following options may be used at a time: json, raw, html, rtf, attachmentsOnly, pdf.')

        # TODO: insert code here that will handle checking all of the msg files to see if the path with overflow.

        if customFilename:
            # First we need to validate it. If there are invalid characters, this will detect it.
            if constants.RE_INVALID_FILENAME_CHARACTERS.search(customFilename):
                raise ValueError('Invalid character found in customFilename. Must not contain any of the following characters: \\/:*?"<>|')
            path /= customFilename[:maxNameLength]
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
                raise ValueError(f'Invalid filename found in self.filename: "{self.filename}"')

            # Add the file name to the path.
            path /= filename[:maxNameLength]
        else:
            path /= self.defaultFolderName[:maxNameLength]

        # Create the folders.
        if not _zip:
            try:
                os.makedirs(path)
            except Exception:
                newDirName = addNumToDir(path)
                if newDirName:
                    path = newDirName
                else:
                    raise Exception(f'Failed to create directory "{path}". Does it already exist?')
        else:
            # In my testing I ended up with multiple files in a zip at the same
            # location so let's try to handle that.
            pathCompare = str(path).rstrip('/') + '/'
            if any(x.startswith(pathCompare) for x in _zip.namelist()):
                newDirName = addNumToZipDir(path, _zip)
                if newDirName:
                    path = newDirName
                else:
                    raise Exception(f'Failed to create directory "{path}". Does it already exist?')

        # Update the kwargs.
        kwargs['customPath'] = path

        if raw:
            self.saveRaw(path)
            return self

        # If the user has requested the headers for this file, save it now.
        if kwargs.get('saveHeader', False):
            headerText = self._getStringStream('__substg1.0_007D')
            if not headerText:
                headerText = constants.HEADER_FORMAT.format(subject = self.subject, **self.header)

            with _open(str(path / 'header.txt'), mode) as f:
                f.write(headerText.encode('utf-8'))

        try:
            if not attachOnly:
                # Check what to save the body with.
                fext = 'json' if _json else 'txt'

                useHtml = False
                usePdf = False
                useRtf = False
                if html:
                    if self.htmlBody:
                        useHtml = True
                        fext = 'html'
                    elif not allowFallback:
                        raise DataNotFoundError('Could not find the htmlBody.')

                if pdf:
                    if self.htmlBody:
                        usePdf = True
                        fext = 'pdf'
                    elif not allowFallback:
                        raise DataNotFoundError('Count not find the htmlBody to convert to pdf.')

                if rtf or (html and not useHtml) or (pdf and not usePdf):
                    if self.rtfBody:
                        useRtf = True
                        fext = 'rtf'
                    elif not allowFallback:
                        raise DataNotFoundError('Could not find the rtfBody.')

            if not skipAttachments:
                # Save the attachments.
                attachmentNames = [attachment.save(**kwargs) for attachment in self.attachments]

            if not attachOnly:
                with _open(str(path / ('message.' + fext)), mode) as f:
                    if _json:
                        emailObj = json.loads(self.getJson())
                        if not skipAttachments:
                            emailObj['attachments'] = attachmentNames

                        f.write(inputToBytes(json.dumps(emailObj), 'utf-8'))
                    elif useHtml:
                        f.write(self.getSaveHtmlBody(**kwargs))
                    elif usePdf:
                        f.write(self.getSavePdfBody(**kwargs))
                    elif useRtf:
                        f.write(self.getSaveRtfBody(**kwargs))
                    else:
                        f.write(self.getSaveBody(**kwargs))

        except Exception:
            if not _zip:
                self.saveRaw(path)
            raise
        finally:
            # Close the ZipFile if this function created it.
            if _zip and createdZip:
                _zip.close()

        # Return the instance so that functions can easily be chained.
        return self

    def getSaveBody(self, **kwargs) -> bytes:
        """
        Returns the plain text body that will be used in saving based on the
        arguments.

        :param **kwargs: Used to allow kwargs expansion in the save function.
            Arguments absorbed by this are simply ignored.
        """
        # Get the type of line endings.
        crlf = inputToBytes(self.crlf, 'utf-8')

        outputBytes = b'From: ' + inputToBytes(self.sender, 'utf-8') + crlf
        outputBytes += b'To: ' + inputToBytes(self.to, 'utf-8') + crlf
        outputBytes += b'Cc: ' + inputToBytes(self.cc, 'utf-8') + crlf
        outputBytes += b'Bcc: ' + inputToBytes(self.bcc, 'utf-8') + crlf
        outputBytes += b'Subject: ' + inputToBytes(self.subject, 'utf-8') + crlf
        outputBytes += b'Date: ' + inputToBytes(self.date, 'utf-8') + crlf
        outputBytes += b'-----------------' + crlf + crlf
        outputBytes += inputToBytes(self.body, 'utf-8')

        return outputBytes

    def getSaveHtmlBody(self, preparedHtml : bool = False, charset : str = 'utf-8', **kwargs) -> bytes:
        """
        Returns the HTML body that will be used in saving based on the
        arguments.

        :param preparedHtml: Whether or not the HTML should be prepared for
            standalone use (add tags, inject images, etc.).
        :param charset: If the html is being prepared, the charset to use for
            the Content-Type meta tag to insert. This exists to ensure that
            something parsing the html can properly determine the encoding (as
            not having this tag can cause errors in some programs). Set this to
            `None` or an empty string to not insert the tag (Default: 'utf-8').
        :param kwargs: Used to allow kwargs expansion in the save function.
            Arguments absorbed by this are simply ignored.

        :raises BadHtmlError: if :param preparedHtml: is False and the HTML
            fails to validate.
        """
        if self.htmlBody:
            # Inject the header into the data.
            data = injectHtmlHeader(self, prepared = preparedHtml)

            # If we are preparing the HTML, then we should
            if preparedHtml and charset:
                bs = bs4.BeautifulSoup(data, features = 'html.parser')
                if not bs.find('meta', {'http-equiv': 'Content-Type'}):
                    # Setup the attributes for the tag.
                    tagAttrs = {
                        'http-equiv': 'Content-Type',
                        'content': f'text/html; charset={charset}',
                    }
                    # Create the tag.
                    tag = bs4.Tag(parser = bs, name = 'meta', attrs = tagAttrs, can_be_empty_element = True)
                    # Add the tag to the head section.
                    if bs.find('head'):
                        bs.find('head').insert(0, tag)
                    else:
                        # If we are here, the head doesn't exist, so let's add
                        # it.
                        if bs.find('html'):
                            # This should always be true, but I want to be safe.
                            head = bs4.Tag(parser = bs, name = 'head')
                            head.insert(0, tag)
                            bs.find('html').insert(0, head)

                    data = bs.prettify('utf-8')

            return data
        else:
            return self.htmlBody

    def getSavePdfBody(self, **kwargs) -> bytes:
        """
        Returns the PDF body that will be used in saving based on the arguments.

        :param wkPath: Used to manually specify the path of the wkhtmltopdf
            executable. If not specified, the function will try to find it.
            Useful if wkhtmltopdf is not on the path. If :param pdf: is False,
            this argument is ignored.
        :param wkOptions: Used to specify additional options to wkhtmltopdf.
            this must be a list or list-like object composed of strings and
            bytes.
        :param kwargs: Used to allow kwargs expansion in the save function.
            Arguments absorbed by this are simply ignored.

        :raises ExecutableNotFound: The wkhtmltopdf executable could not be
            found.
        :raises WKError: Something went wrong in creating the PDF body.
        """
        # Immediately try to find the executable.
        wkPath = findWk(kwargs.get('wkPath'))

        # First thing is first, we need to parse our wkOptions if
        # they exist.
        wkOptions = kwargs.get('wkOptions')
        if wkOptions:
            try:
                # Try to convert to a list, whatever it is, and
                # fail if it is not possible.
                parsedWkOptions = [*wkOptions]
            except TypeError:
                raise TypeError(':param wkOptions: must be an iterable, not {type(wkOptions)}.')
        else:
            parsedWkOptions = []

        # Confirm that all of our options we now have are either
        # strings or bytes.
        if not all(isinstance(option, (str, bytes)) for option in parsedWkOptions):
            raise TypeError(':param wkOptions: must be an iterable of strings and bytes.')

        # We call the program to convert the html, but give tell it
        # the data will go in and come out through stdin and stdout,
        # respectively. This way we don't have to write temporary
        # files to the disk. We also ask that it be quiet about it.
        process = subprocess.Popen([wkPath, *parsedWkOptions, '-', '-'], shell = True, stdin = subprocess.PIPE, stdout = subprocess.PIPE, stderr = subprocess.PIPE)
        # Give the program the data and wait for the program to
        # finish.
        output = process.communicate(self.getSaveHtmlBody(**kwargs))
        if process.returncode != 0:
            raise WKError(output[1].decode('utf-8'))

        return output[0]

    def getSaveRtfBody(self, **kwargs) -> bytes:
        """
        Returns the RTF body that will be used in saving based on the arguments.

        :param kwargs: Used to allow kwargs expansion in the save function.
            Arguments absorbed by this are simply ignored.
        """
        # Inject the header into the data.
        return injectRtfHeader(self)
