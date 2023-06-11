from __future__ import annotations


__all__ = [
    'MessageBase',
]


import base64
import datetime
import email.utils
import functools
import html
import json
import logging
import os
import pathlib
import re
import subprocess
import zipfile

import bs4
import chardet
import compressed_rtf
import RTFDE

from email.parser import Parser as EmailParser
from typing import Callable, List, Optional, Union

from . import constants
from ._rtf.create_doc import createDocument
from ._rtf.inject_rtf import injectStartRTF
from .enums import BodyTypes, DeencapType, RecipientType
from .exceptions import (
        BadHtmlError, DataNotFoundError, DeencapMalformedData,
        DeencapNotEncapsulated, IncompatibleOptionsError, WKError
    )
from .msg import MSGFile
from .structures.report_tag import ReportTag
from .recipient import Recipient
from .utils import (
        addNumToDir, addNumToZipDir, createZipOpen, decodeRfc2047, findWk,
        htmlSanitize, inputToBytes, inputToString, isEncapsulatedRtf,
        prepareFilename, rtfSanitizeHtml, rtfSanitizePlain, validateHtml
    )

from imapclient.imapclient import decode_utf7


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
            extract_msg.exceptions. All other exceptions must be handled
            internally or they will not be caught. The original deencapsulation
            method will not run if this is set.
        """
        super().__init__(path, **kwargs)
        recipientSeparator = ';'
        self.__recipientSeparator = kwargs.get('recipientSeparator', ';')
        self.__ignoreRtfDeErrors = kwargs.get('ignoreRtfDeErrors', False)
        self.__deencap = kwargs.get('deencapsulationFunc')
        # Initialize properties in the order that is least likely to cause bugs.
        # TODO have each function check for initialization of needed data so
        # these lines will be unnecessary.
        self.props
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

    def _genRecipient(self, recipientType, recipientInt : RecipientType) -> Optional[str]:
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
                    value = decodeRfc2047(value)
                    value = value.replace(',', self.__recipientSeparator)

            # If the header had a blank field or didn't have the field, generate
            # it manually.
            if not value:
                # Check if the header has initialized.
                if self.headerInit():
                    logger.info(f'Header found, but "{recipientType}" is not included. Will be generated from other streams.')

                # Get a list of the recipients of the specified type.
                foundRecipients = tuple(recipient.formatted for recipient in self.recipients if recipient.type == recipientInt)

                # If we found recipients, join them with the recipient separator
                # and a space.
                if len(foundRecipients) > 0:
                    value = (self.__recipientSeparator + ' ').join(foundRecipients)

            # Code to fix the formatting so it's all a single line. This allows
            # the user to format it themself if they want. This should probably
            # be redone to use re or something, but I can do that later. This
            # shouldn't be a huge problem for now.
            if value:
                value = value.replace(' \r\n\t', ' ').replace('\r\n\t ', ' ').replace('\r\n\t', ' ')
                value = value.replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
                while value.find('  ') != -1:
                    value = value.replace('  ', ' ')

            # Set the field in the class.
            setattr(self, private, value)

            return value

    def deencapsulateBody(self, rtfBody : bytes, bodyType : DeencapType) -> Optional[Union[bytes, str]]:
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

    def dump(self) -> None:
        """
        Prints out a summary of the message.
        """
        print('Message')
        print('Subject:', self.subject)
        print('Date:', self.date)
        print('Body:')
        print(self.body)

    def getInjectableHeader(self, prefix : str, joinStr : str, suffix : str, formatter : Callable[[str, str], str]) -> str:
        """
        Using the specified prefix, suffix, formatter, and join string,
        generates the injectable header. Prefix is placed at the beginning,
        followed by a series of format strings joined together with the join
        string, with the suffix placed afterwards. Effectively makes this
        structure:
        {prefix}{formatter()}{joinStr}{formatter()}{joinStr}...{formatter()}{suffix}

        Formatter be a function that takes first a name variable then a value
        variable and formats the line.

        If self.headerFormatProperties is None, immediately returns an empty
        string.
        """
        formattedProps = []
        allProps = self.headerFormatProperties

        if allProps is None:
            return ''

        for entry in allProps:
            isGroup = False
            entryUsed = False
            # This is how we handle the groups.
            if isinstance(allProps[entry], dict):
                props = allProps[entry]
                isGroup = True
            else:
                props = {entry: allProps[entry]}

            for name in props:
                if props[name]:
                    if isinstance(props[name], tuple):
                        if props[name][1]:
                            value = props[name][0] or ''
                        elif props[name][0] is not None:
                            value = props[name][0]
                        else:
                            continue
                    else:
                        value = props[name]

                    entryUsed = True
                    formattedProps.append(formatter(name, value))

            # Now if we are working with a group, add an empty entry to get a
            # second join string between this section and the last, but *only*
            # if any of the entries were used.
            if isGroup and entryUsed:
                formattedProps.append('')

        # If the last entry is empty, remove it. We don't want extra spacing at
        # the end.
        if formattedProps[-1] == '':
            formattedProps.pop()

        return prefix + joinStr.join(formattedProps) + suffix

    def getJson(self) -> str:
        """
        Returns the JSON representation of the Message.
        """
        return json.dumps({
            'from': inputToString(self.sender, self.stringEncoding),
            'to': inputToString(self.to, self.stringEncoding),
            'cc': inputToString(self.cc, self.stringEncoding),
            'bcc': inputToString(self.bcc, self.stringEncoding),
            'subject': inputToString(self.subject, self.stringEncoding),
            'date': inputToString(self.date, self.stringEncoding),
            'body': decode_utf7(self.body),
        })

    def getSaveBody(self, **kwargs) -> bytes:
        """
        Returns the plain text body that will be used in saving based on the
        arguments.

        :param kwargs: Used to allow kwargs expansion in the save function.
            Arguments absorbed by this are simply ignored.
        """
        # Get the type of line endings.
        crlf = inputToString(self.crlf, 'utf-8')

        prefix = ''
        suffix = crlf + '-----------------' + crlf + crlf
        joinStr = crlf
        formatter = (lambda name, value : f'{name}: {value}')

        header = self.getInjectableHeader(prefix, joinStr, suffix, formatter).encode('utf-8')
        return header + inputToBytes(self.body, 'utf-8')

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
            data = self.injectHtmlHeader(prepared = preparedHtml)

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

        # First thing is first, we need to parse our wkOptions if they exist.
        wkOptions = kwargs.get('wkOptions')
        if wkOptions:
            try:
                # Try to convert to a list, whatever it is, and fail if it is
                # not possible.
                parsedWkOptions = [*wkOptions]
            except TypeError:
                raise TypeError(f':param wkOptions: must be an iterable, not {type(wkOptions)}.')
        else:
            parsedWkOptions = []

        # Confirm that all of our options we now have are either strings or
        # bytes.
        if not all(isinstance(option, (str, bytes)) for option in parsedWkOptions):
            raise TypeError(':param wkOptions: must be an iterable of strings and bytes.')

        processArgs = [wkPath, *parsedWkOptions, '-', '-']
        # Log the arguments.
        logger.info(f'Converting to PDF with the following arguments: {processArgs}')

        # Get the html body *before* calling Popen.
        htmlBody = self.getSaveHtmlBody(**kwargs)

        # We call the program to convert the html, but give tell it the data
        # will go in and come out through stdin and stdout, respectively. This
        # way we don't have to write temporary files to the disk. We also ask
        # that it be quiet about it.
        process = subprocess.run(processArgs, input = htmlBody, stdout = subprocess.PIPE, stderr = subprocess.PIPE)
        # Give the program the data and wait for the program to finish.
        #output = process.communicate(htmlBody)

        # If it errored, throw it as an exception.
        if process.returncode != 0:
            raise WKError(process.stderr.decode('utf-8'))

        return process.stdout

    def getSaveRtfBody(self, **kwargs) -> bytes:
        """
        Returns the RTF body that will be used in saving based on the arguments.

        :param kwargs: Used to allow kwargs expansion in the save function.
            Arguments absorbed by this are simply ignored.
        """
        # Inject the header into the data.
        return self.injectRtfHeader()

    def headerInit(self) -> bool:
        """
        Checks whether the header has been initialized.
        """
        try:
            self._header
            return True
        except AttributeError:
            return False

    def injectHtmlHeader(self, prepared : bool = False) -> bytes:
        """
        Returns the HTML body from the MSG file (will check that it has one) with
        the HTML header injected into it.

        :param prepared: Determines whether to be using the standard HTML (False) or
            the prepared HTML (True) body (Default: False).

        :raises AttributeError: if the correct HTML body cannot be acquired.
        :raises BadHtmlError: if :param preparedHtml: is False and the HTML fails to
            validate.
        """
        if not self.htmlBody:
            raise AttributeError('Cannot inject the HTML header without an HTML body attribute.')

        body = None

        # We don't do this all at once because the prepared body is not cached.
        if prepared:
            body = self.htmlBodyPrepared

            # If the body is not valid or not found, raise an AttributeError.
            if not body:
                raise AttributeError('Cannot find a prepared HTML body to inject into.')
        else:
            body = self.htmlBody

        # Validate the HTML.
        if not validateHtml(body):
            # If we are not preparing the HTML body, then raise an
            # exception.
            if not prepared:
                raise BadHtmlError('HTML body failed to pass validation.')

            # If we are here, then we need to do what we can to fix the HTML body.
            # Unfortunately this gets complicated because of the various ways the
            # body could be wrong. If only the <body> tag is missing, then we just
            # need to insert it at the end and be done. If both the <html> and
            # <body> tag are missing, we determine where to put the body tag (around
            # everything if there is no <head> tag, otherwise at the end) and then
            # wrap it all in the <html> tag.
            parser = bs4.BeautifulSoup(body, features = 'html.parser')
            if not parser.find('html') and not parser.find('body'):
                if parser.find('head') or parser.find('footer'):
                    # Create the parser we will be using for the corrections.
                    correctedHtml = bs4.BeautifulSoup(b'<html></html>', features = 'html.parser')
                    htmlTag = correctedHtml.find('html')

                    # Iterate over each of the direct descendents of the parser and
                    # add each to a new tag if they are not the head or footer.
                    bodyTag = parser.new_tag('body')
                    # What we are going to be doing will be causing some of the tags
                    # to be moved out of the parser, and so the iterator will end up
                    # pointing to the wrong place after that. To compensate we first
                    # create a tuple and iterate over that.
                    for tag in tuple(parser.children):
                        if tag.name.lower() in ('head', 'footer'):
                            correctedHtml.append(tag)
                        else:
                            bodyTag.append(tag)

                    # All the tags should now be properly in the body, so let's
                    # insert it.
                    if correctedHtml.find('head'):
                        correctedHtml.find('head').insert_after(bodyTag)
                    elif correctedHtml.find('footer'):
                        correctedHtml.find('footer').insert_before(bodyTag)
                    else:
                        # Neither a head or a body are present, so just append it to
                        # the main tag.
                        htmlTag.append(bodyTag)
                else:
                    # If there is no <html>, <head>, <footer>, or <body> tag, then
                    # we just add the tags to the beginning and end of the data and
                    # move on.
                    body = b'<html><body>' + body + b'</body></html>'
            elif parser.find('html'):
                # Found <html> but not <body>.
                # Iterate over each of the direct descendents of the parser and
                # add each to a new tag if they are not the head or footer.
                bodyTag = parser.new_tag('body')
                # What we are going to be doing will be causing some of the tags
                # to be moved out of the parser, and so the iterator will end up
                # pointing to the wrong place after that. To compensate we first
                # create a tuple and iterate over that.
                for tag in tuple(parser.find('html').children):
                    if tag.name and tag.name.lower() not in ('head', 'footer'):
                        bodyTag.append(tag)

                # All the tags should now be properly in the body, so let's
                # insert it.
                if parser.find('head'):
                    parser.find('head').insert_after(bodyTag)
                elif parser.find('footer'):
                    parser.find('footer').insert_before(bodyTag)
                else:
                    parser.find('html').insert(0, bodyTag)
            else:
                # Found <body> but not <html>. Just wrap everything in the <html>
                # tags.
                body = b'<html>' + body + b'</html>'

        def replace(bodyMarker):
            """
            Internal function to replace the body tag with itself plus the header.
            """
            # I recently had to change this and how it worked. Now we use a new
            # property of `MSGFile` that returns a special tuple of tuples to define
            # how to get all of the properties we are formatting. They are all
            # processed in the same way, making everything neat. By defining them
            # in each class, any class can specify a completely different set to be
            # used.
            return bodyMarker.group() + self.htmlInjectableHeader.encode('utf-8')

        # Use the previously defined function to inject the HTML header.
        return constants.RE_HTML_BODY_START.sub(replace, body, 1)

    def injectRtfHeader(self) -> bytes:
        """
        Returns the RTF body from this MSG file (will check that it has one)
        with the RTF header injected into it.

        :raises AttributeError: if the RTF body cannot be acquired.
        :raises RuntimeError: if all injection attempts fail.
        """
        if not self.rtfBody:
            raise AttributeError('Cannot inject the RTF header without an RTF body attribute.')

        # Try to determine which header to use. Also determines how to sanitize the
        # rtf.
        if isEncapsulatedRtf(self.rtfBody):
            injectableHeader = self.rtfEncapInjectableHeader
        else:
            injectableHeader = self.rtfPlainInjectableHeader

        def replace(bodyMarker):
            """
            Internal function to replace the body tag with itself plus the header.
            """
            return bodyMarker.group() + injectableHeader

        # This first method only applies to documents with encapsulated HTML
        # that is formatted in a nice way.
        if isEncapsulatedRtf(self.rtfBody):
            data = constants.RE_RTF_ENC_BODY_START.sub(replace, self.rtfBody, 1)
            if data != self.rtfBody:
                logger.debug('Successfully injected RTF header using encapsulation method.')
                return data
            logger.debug('RTF has encapsulated HTML, but injection method failed. It is likely dirty. Will use normal RTF injection method.')

        # If the normal encapsulated HTML injection fails or it isn't
        # encapsulated, use the internal _rtf module.
        logger.debug('Using _rtf module to inject RTF text header.')
        return createDocument(injectStartRTF(self.rtfBody, injectableHeader))

    def save(self, **kwargs) -> MessageBase:
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
        :param saveHeader: Turns on saving the header as a separate file when
        set.
        :param skipAttachments: Turns off saving attachments.
        :param skipHidden: If True, skips attachments marked as hidden.
            (Default: False)
        :param skipBodyNotFound: Suppresses errors if no valid body could be
            found, simply skipping the step of saving the body.
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
        skipHidden = kwargs.get('skipHidden', False)
        # Track if we should skip the body if no valid body is found instead of
        # raising an exception.
        skipBodyNotFound = kwargs.get('skipBodyNotFound', False)

        if pdf:
            kwargs['preparedHtml'] = True

        # ZipFile handling.
        if _zip:
            # `raw` and `zip` are incompatible.
            if raw:
                raise IncompatibleOptionsError('The options `raw` and `zip` are incompatible.')
            # If we are doing a zip file, first check that we have been given a
            # path.
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

        # TODO: insert code here that will handle checking all of the msg files
        # to see if the path with overflow.

        if customFilename:
            # First we need to validate it. If there are invalid characters,
            # this will detect it.
            if constants.RE_INVALID_FILENAME_CHARACTERS.search(customFilename):
                raise ValueError('Invalid character found in customFilename. Must not contain any of the following characters: \\/:*?"<>|')
            # Quick fix to remove spaces from the end of the filename, if any
            # are there.
            customFilename = customFilename.strip()
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
            pathCompare = str(path).replace('\\', '/').rstrip('/') + '/'
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
            headerText = self.headerText
            if not headerText:
                headerText = constants.HEADER_FORMAT.format(subject = self.subject, **self.header)

            with _open(str(path / 'header.txt'), mode) as f:
                f.write(headerText.encode('utf-8'))

        try:
            if not attachOnly:
                # Check what to save the body with.
                fext = 'json' if _json else 'txt'

                fallbackToPlain = False
                useHtml = False
                usePdf = False
                useRtf = False
                if html:
                    if self.htmlBody:
                        useHtml = True
                        fext = 'html'
                    elif not allowFallback:
                        if skipBodyNotFound:
                            fext = None
                        else:
                            raise DataNotFoundError('Could not find the htmlBody.')

                if pdf:
                    if self.htmlBody:
                        usePdf = True
                        fext = 'pdf'
                    elif not allowFallback:
                        if skipBodyNotFound:
                            fext = None
                        else:
                            raise DataNotFoundError('Count not find the htmlBody to convert to pdf.')

                if rtf or (html and not useHtml) or (pdf and not usePdf):
                    if self.rtfBody:
                        useRtf = True
                        fext = 'rtf'
                    elif not allowFallback:
                        if skipBodyNotFound:
                            fext = None
                        else:
                            raise DataNotFoundError('Could not find the rtfBody.')
                    else:
                        # This was the last resort before plain text, so fall
                        # back to that.
                        fallbackToPlain = True

                # After all other options, try to go with plain text if
                # possible.
                if not (rtf or html or pdf) or fallbackToPlain:
                    # We need to check if the plain text body was found. If it
                    # was found but was empty that is considered valid, so we
                    # specifically check against None.
                    if self.body is None:
                        if skipBodyNotFound:
                            fext = None
                        else:
                            if allowFallback:
                                raise DataNotFoundError('Could not find a valid body using current options.')
                            else:
                                raise DataNotFoundError('Plain text body could not be found.')


            if not skipAttachments:
                # Save the attachments.
                attachmentNames = [attachment.save(**kwargs) for attachment in self.attachments if not (skipHidden and attachment.hidden)]
                # Remove skipped attachments.
                attachmentNames = [x for x in attachmentNames if x and isinstance(x, str)]

            if not attachOnly and fext:
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
        finally:
            # Close the ZipFile if this function created it.
            if _zip and createdZip:
                _zip.close()

        # Return the instance so that functions can easily be chained.
        return self

    @property
    def bcc(self) -> Optional[str]:
        """
        Returns the bcc field, if it exists.
        """
        return self._genRecipient('bcc', RecipientType.BCC)

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
    def cc(self) -> Optional[str]:
        """
        Returns the cc field, if it exists.
        """
        return self._genRecipient('cc', RecipientType.CC)

    @property
    def compressedRtf(self) -> Optional[bytes]:
        """
        Returns the compressed RTF stream, if it exists.
        """
        return self._ensureSet('_compressedRtf', '__substg1.0_10090102', False)

    @property
    def crlf(self) -> str:
        """
        Returns the value of self.__crlf, should you need it for whatever
        reason.
        """
        self.body
        return self.__crlf

    @property
    def date(self) -> Optional[str]:
        """
        Returns the send date, if it exists.
        """
        try:
            return self._date
        except AttributeError:
            self._date = self._prop.date if self.isSent else None
            return self._date

    @property
    def deencapsulatedRtf(self) -> Optional[RTFDE.DeEncapsulator]:
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
                        logger.warning(f'RTFDE failed to decode rtfBody for message with subject "{self.subject}". Attempting to cut out unnecessary data and override decoding.')

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
            dirName = dirName.strip()

            self._defaultFolderName = dirName
            return dirName

    @functools.cached_property
    def detectedBodies(self) -> BodyTypes:
        """
        The types of bodies stored in the .msg file.
        """
        bodies = BodyTypes.NONE
        if self.sExists('__substg1.0_1000'):
            bodies |= BodyTypes.PLAIN
        if self.exists('__substg1.0_10090102'):
            bodies |= BodyTypes.RTF
        if self.exists('__substg1.0_10130102'):
            bodies |= BodyTypes.HTML

        return bodies

    @property
    def header(self) -> email.message.Message:
        """
        Returns the message header, if it exists. Otherwise it will generate
        one.
        """
        try:
            return self._header
        except AttributeError:
            headerText = self.headerText
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
    def headerFormatProperties(self) -> constants.HEADER_FORMAT_TYPE:
        """
        Returns a dictionary of properties, in order, to be formatted into the
        header. Keys are the names to use in the header while the values are one
        of the following:
        None: Signifies no data was found for the property and it should be
            omitted from the header.
        str: A string to be formatted into the header using the string encoding.
        Tuple[Union[str, None], bool]: A string should be formatted into the
            header. If the bool is True, then place an empty string if the value
            is None, otherwise follow the same behavior as regular None.

        Additional note: If the value is an empty string, it will be dropped as
        well by default.

        Additionally you can group members of a header together by placing them
        in an embedded dictionary. Groups will be spaced out using a second
        instance of the join string. If any member of a group is being printed,
        it will be spaced apart from the next group/item.

        If you class should not do *any* header injection, return None from this
        property.
        """
        # Checking outlook printing, default behavior is to completely omit
        # *any* field that is not present. So while for extensability the
        # option exists to have it be present even if no data is found, we are
        # specifically not doing that.
        return {
            '-basic info-': {
                'From': self.sender,
                'Sent': self.date,
                'To': self.to,
                'Cc': self.cc,
                'Bcc': self.bcc,
                'Subject': self.subject,
            },
            '-importance-': {
                'Importance': self.importanceString,
            },
        }

    @functools.cached_property
    def headerText(self) -> Optional[str]:
        """
        The raw text of the header stream, if it exists.
        """
        return self._getStringStream('__substg1.0_007D')

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
            elif self.rtfBody:
                logger.info('HTML body was not found, attempting to generate from RTF.')
                self._htmlBody = self.deencapsulateBody(self.rtfBody, DeencapType.HTML)
            # This is it's own if statement so we can ensure it will generate
            # even if there is an rtfBody, in the event it doesn't have HTML.
            if not self._htmlBody and self.body:
                # Convert the plain text body to html.
                logger.info('HTML body was not found, attempting to generate from plain text body.')
                correctedBody = html.escape(self.body).replace('\r', '').replace('\n', '<br />')
                self._htmlBody = f'<html><body>{correctedBody}</body></head>'.encode('utf-8')

            if not self._htmlBody:
                logger.info('HTML body could not be found nor generated.')

            return self._htmlBody

    @property
    def htmlBodyPrepared(self) -> Optional[bytes]:
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
    def htmlInjectableHeader(self) -> str:
        """
        The header that can be formatted and injected into the html body.
        """
        prefix = '<div id="injectedHeader"><div><p class="MsoNormal">'
        suffix = '<o:p></o:p></p></div></div>'
        joinStr = '<br/>'
        formatter = (lambda name, value : f'<b>{name}:</b>&nbsp;{inputToString(htmlSanitize(value), self.stringEncoding)}')

        return self.getInjectableHeader(prefix, joinStr, suffix, formatter)

    @property
    def inReplyTo(self) -> Optional[str]:
        """
        Returns the message id that this message is in reply to.
        """
        return self._ensureSet('_in_reply_to', '__substg1.0_1042')

    @property
    def isRead(self) -> bool:
        """
        Returns if this email has been marked as read.
        """
        return bool(self.props['0E070003'].value & 1)

    @property
    def isSent(self) -> bool:
        """
        Returns if this email has been marked as sent. Assumes True if no flags
        are found.
        """
        if not self.props.get('0E070003'):
            return True
        else:
            return not bool(self.props['0E070003'].value & 8)

    @property
    def messageId(self) -> Optional[str]:
        try:
            return self._messageId
        except AttributeError:
            headerResult = None
            if self.headerInit():
                headerResult = self.header['message-id']
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
    def receivedTime(self) -> Optional[datetime.datetime]:
        """
        The date and time the message was received by the server.
        """
        return self._ensureSetProperty('_receivedTime', '0E060040')

    @property
    def recipientSeparator(self) -> str:
        return self.__recipientSeparator

    @property
    def recipients(self) -> List[Recipient]:
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
    def reportTag(self) -> Optional[ReportTag]:
        """
        Data that is used to correlate the report and the original message.
        """
        return self._ensureSet('_reportTag', '__substg1.0_00310102', False, overrideClass = ReportTag)

    @property
    def rtfBody(self) -> Optional[bytes]:
        """
        Returns the decompressed Rtf body from the message.
        """
        try:
            return self._rtfBody
        except AttributeError:
            self._rtfBody = compressed_rtf.decompress(self.compressedRtf) if self.compressedRtf else None
            return self._rtfBody

    @property
    def rtfEncapInjectableHeader(self) -> bytes:
        """
        The header that can be formatted and injected into the plain RTF body.
        """
        prefix = r'\htmlrtf {\htmlrtf0 {\*\htmltag96 <div>}{\*\htmltag96 <div>}{\*\htmltag64 <p class=MsoNormal>}'
        suffix = r'{\*\htmltag244 <o:p>}{\*\htmlrag252 </o:p>}\htmlrtf \par\par\htmlrtf0 {\*\htmltag72 </p>}{\*\htmltag104 </div>}{\*\htmltag104 </div>}\htmlrtf }\htmlrtf0 '
        joinStr = r'{\*\htmltag116 <br />}\htmlrtf \line\htmlrtf0 '
        formatter = (lambda name, value : fr'\htmlrtf {{\b\htmlrtf0{{\*\htmltag84 <b>}}{name}: {{\*\htmltag92 </b>}}\htmlrtf \b0\htmlrtf0 {inputToString(rtfSanitizeHtml(value), self.stringEncoding)}\htmlrtf }}\htmlrtf0')

        return self.getInjectableHeader(prefix, joinStr, suffix, formatter).encode('utf-8')

    @property
    def rtfPlainInjectableHeader(self) -> bytes:
        """
        The header that can be formatted and injected into the encapsulated RTF
        body.
        """
        prefix = '{'
        suffix = r'\par\par}'
        joinStr = r'\line'
        formatter = (lambda name, value : fr'{{\b {name}: \b0 {inputToString(rtfSanitizePlain(value), self.stringEncoding)}}}')

        return self.getInjectableHeader(prefix, joinStr, suffix, formatter).encode('utf-8')

    @property
    def sender(self) -> Optional[str]:
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
                    self._sender = decodeRfc2047(headerResult)
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
    def subject(self) -> Optional[str]:
        """
        Returns the message subject, if it exists.
        """
        return self._ensureSet('_subject', '__substg1.0_0037')

    @property
    def to(self) -> Optional[str]:
        """
        Returns the to field, if it exists.
        """
        return self._genRecipient('to', RecipientType.TO)
