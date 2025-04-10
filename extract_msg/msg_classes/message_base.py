__all__ = [
    'MessageBase',
]


import base64
import datetime
import email.message
import email.utils
import enum
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
import compressed_rtf
import RTFDE
import RTFDE.exceptions

from email import policy
from email.charset import Charset, QP
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.parser import HeaderParser
from typing import Any, Callable, cast, Dict, List, Optional, Tuple, Type, Union

from .. import constants
from .._rtf.create_doc import createDocument
from .._rtf.inject_rtf import injectStartRTF
from ..enums import (
        BodyTypes, DeencapType, ErrorBehavior, RecipientType, SaveType
    )
from ..exceptions import (
        ConversionError, DataNotFoundError, DeencapMalformedData,
        DeencapNotEncapsulated, IncompatibleOptionsError, MimetypeFailureError,
        WKError
    )
from .msg import MSGFile
from ..structures.report_tag import ReportTag
from ..recipient import Recipient
from ..utils import (
        addNumToDir, addNumToZipDir, createZipOpen, decodeRfc2047, findWk,
        htmlSanitize, inputToBytes, inputToString, isEncapsulatedRtf,
        prepareFilename, rtfSanitizeHtml, rtfSanitizePlain, stripRtf,
        validateHtml
    )


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class MessageBase(MSGFile):
    """
    Base class for Message-like MSG files.
    """

    def __init__(self, path, **kwargs):
        """
        Supports all of the options from :meth:`MSGFile.__init__` with some
        additional ones.

        :param recipientSeparator: Optional, separator string to use between
            recipients.
        :param deencapsulationFunc: Optional, if specified must be a callable
            that will override the way that HTML/text is deencapsulated from the
            RTF body. This function must take exactly 2 arguments, the first
            being the RTF body from the message and the second being an instance
            of the enum ``DeencapType`` that will tell the function what type of
            body is desired. The function should return a string for plain text
            and bytes for HTML. If any problems occur, the function *must*
            either return ``None`` or raise one of the appropriate exceptions
            from :mod:`extract_msg.exceptions`. All other exceptions must be
            handled internally or they will not be caught. The original
            deencapsulation method will not run if this is set.
        """
        super().__init__(path, **kwargs)
        # The rest needs to be in a try-except block to ensure the file closes
        # if an error occurs.
        try:
            self.__headerInit = False
            self.__recipientSeparator: str = kwargs.get('recipientSeparator', ';')
            self.__deencap = kwargs.get('deencapsulationFunc')
            self.header

            # This variable keeps track of what the new line character should be.
            self._crlf = '\n'
            try:
                self.body
            except Exception as e:
                # Prevent an error in the body from preventing opening.
                logger.exception('Critical error accessing the body. File opened but accessing the body will throw an exception.')
            self._htmlEncoding = None
        except:
            try:
                self.close()
            except:
                pass
            raise

    def _genRecipient(self, recipientStr: str, recipientType: RecipientType) -> Optional[str]:
        """
        Method to generate the specified recipient field.
        """
        value = None
        # Check header first.
        if self.headerInit:
            value = cast(Optional[str], self.header[recipientStr])
            if value:
                value = decodeRfc2047(value)
                value = value.replace(',', self.__recipientSeparator)

        # If the header had a blank field or didn't have the field, generate
        # it manually.
        if not value:
            # Check if the header has initialized.
            if self.headerInit:
                logger.info(f'Header found, but "{recipientStr}" is not included. Will be generated from other streams.')

            # Get a list of the recipients of the specified type.
            foundRecipients = tuple(recipient.formatted for recipient in self.recipients if recipient.type is recipientType)

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

        return value

    def _getHtmlEncoding(self, soup: bs4.BeautifulSoup) -> None:
        """
        Helper function to set the html encoding.
        """
        if not self._htmlEncoding:
            try:
                self._htmlEncoding = cast(Optional[str], soup.original_encoding or soup.declared_html_encoding)
            except AttributeError:
                pass

    def asEmailMessage(self) -> EmailMessage:
        """
        Returns an instance of EmailMessage used to represent the contents of
        this message.

        :raises ConversionError: The function failed to convert one of the
            attachments into a form that it could attach, and the attachment
            data type was not None.
        """
        ret = EmailMessage()

        # Copy the headers.
        for key, value in self.header.items():
            if key.lower() != 'content-type':
                ret[key] = value.replace('\r\n', '').replace('\n', '')

        ret['Content-Type'] = 'multipart/mixed'

        # Attach the body to the EmailMessage instance.
        msgMain = MIMEMultipart('related')
        ret.attach(msgMain)
        bodyParts = MIMEMultipart('alternative')
        msgMain.attach(bodyParts)

        c = Charset('utf-8')
        c.body_encoding = QP

        if self.body:
            bodyParts.attach(MIMEText(self.body, 'plain', c))
        if self.htmlBody:
            bodyParts.attach(MIMEText(self.htmlBody.decode('utf-8'), 'html', c))

        # Process attachments.
        for att in self.attachments:
            if att.dataType:
                if hasattr(att.dataType, 'asEmailMessage'):
                    # Replace the extension with '.eml'.
                    filename = att.name or ''
                    if filename.lower().endswith('.msg'):
                        filename = filename[:-4] + '.eml'
                    msgMain.attach(att.data.asEmailMessage())
                else:
                    if issubclass(att.dataType, bytes):
                        data = att.data
                    elif issubclass(att.dataType, MSGFile):
                        if hasattr(att.dataType, 'asBytes'):
                            data = att.asBytes
                        else:
                            data = att.data.exportBytes()
                    else:
                        raise ConversionError(f'Could not find a suitable method to attach attachment data type "{att.dataType}".')
                    mime = att.mimetype or 'application/octet-stream'
                    mainType, subType = mime.split('/')[0], mime.split('/')[-1]
                    # Need to do this manually instead of using add_attachment.
                    attachment = EmailMessage()
                    attachment.set_content(data,
                                           maintype = mainType,
                                           subtype = subType,
                                           cid = att.contentId)
                    # This is just a very basic check.
                    attachment['Content-Disposition'] = f'{"inline" if att.hidden else "attachment"}; filename="{att.getFilename()}"'

                    # Add the attachment.
                    msgMain.attach(attachment)

        return ret

    def deencapsulateBody(self, rtfBody: bytes, bodyType: DeencapType) -> Optional[Union[bytes, str]]:
        """
        A method to deencapsulate the specified body from the RTF body.

        Returns a string for plain text and bytes for HTML. If specified, uses
        the deencapsulation override function. Returns ``None`` if nothing
        could be deencapsulated.

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
                        return self.deencapsulatedRtf.html

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
        if self.date:
            print('Date:', self.date.__format__(self.datetimeFormat))
        print('Body:')
        print(self.body)

    def getInjectableHeader(self, prefix: str, joinStr: str, suffix: str, formatter: Callable[[str, str], str]) -> str:
        """
        Using the specified prefix, suffix, formatter, and join string,
        generates the injectable header.

        Prefix is placed at the beginning, followed by a series of format
        strings joined together with the join string, with the suffix placed
        afterwards. Effectively makes this structure:
        {prefix}{formatter()}{joinStr}{formatter()}{joinStr}...{formatter()}{suffix}

        Formatter be a function that takes first a name variable then a value
        variable and formats the line.

        If self.headerFormatProperties is None, immediately returns an empty
        string.
        """
        allProps = self.headerFormatProperties

        if allProps is None:
            return ''

        formattedProps = []

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
            'from': self.sender,
            'to': self.to,
            'cc': self.cc,
            'bcc': self.bcc,
            'subject': self.subject,
            'date': self.date.__format__(self.datetimeFormat) if self.date else None,
            'body': self.body,
        })

    def getSaveBody(self, **_) -> bytes:
        """
        Returns the plain text body that will be used in saving based on the
        arguments.

        :param _: Used to allow kwargs expansion in the save function.
            Arguments absorbed by this are simply ignored.
        """
        # Get the type of line endings.
        crlf = inputToString(self.crlf, 'utf-8')

        prefix = ''
        suffix = crlf + '-----------------' + crlf + crlf
        joinStr = crlf
        formatter = (lambda name, value: f'{name}: {value}')

        header = self.getInjectableHeader(prefix, joinStr, suffix, formatter).encode('utf-8')
        return header + inputToBytes(self.body, 'utf-8')

    def getSaveHtmlBody(self, preparedHtml: bool = False, charset: str = 'utf-8', **_) -> bytes:
        """
        Returns the HTML body that will be used in saving based on the
        arguments.

        :param preparedHtml: Whether or not the HTML should be prepared for
            standalone use (add tags, inject images, etc.).
        :param charset: If the html is being prepared, the charset to use for
            the Content-Type meta tag to insert. This exists to ensure that
            something parsing the html can properly determine the encoding (as
            not having this tag can cause errors in some programs). Set this to
            ``None`` or an empty string to not insert the tag. (Default:
            'utf-8')
        :param _: Used to allow kwargs expansion in the save function.
            Arguments absorbed by this are simply ignored.
        """
        if self.htmlBody:
            # Inject the header into the data.
            data = self.injectHtmlHeader(prepared = preparedHtml)

            # If we are preparing the HTML, then we should
            if preparedHtml and charset:
                bs = bs4.BeautifulSoup(data, features = 'html.parser', from_encoding = self._htmlEncoding)
                self._getHtmlEncoding(bs)
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

                    data = bs.encode('utf-8')

            return data
        else:
            return self.htmlBody or b''

    def getSavePdfBody(self, wkPath = None, wkOptions = None, **kwargs) -> bytes:
        """
        Returns the PDF body that will be used in saving based on the arguments.

        :param wkPath: Used to manually specify the path of the wkhtmltopdf
            executable. If not specified, the function will try to find it.
            Useful if wkhtmltopdf is not on the path. If :param pdf: is
            ``False``, this argument is ignored.
        :param wkOptions: Used to specify additional options to wkhtmltopdf.
            this must be a list or list-like object composed of strings and
            bytes.
        :param kwargs: Used to allow kwargs expansion in the save function.
            Arguments absorbed by this are simply ignored, except for keyword arguments used by :meth:`getSaveHtmlBody`.

        :raises ExecutableNotFound: The wkhtmltopdf executable could not be
            found.
        :raises WKError: Something went wrong in creating the PDF body.
        """
        # Immediately try to find the executable.
        wkPath = findWk(wkPath)

        # First thing is first, we need to parse our wkOptions if they exist.
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

    def getSaveRtfBody(self, **_) -> bytes:
        """
        Returns the RTF body that will be used in saving based on the arguments.

        :param kwargs: Used to allow kwargs expansion in the save function.
            Arguments absorbed by this are simply ignored.
        """
        # Inject the header into the data.
        return self.injectRtfHeader()

    def injectHtmlHeader(self, prepared: bool = False) -> bytes:
        """
        Returns the HTML body from the MSG file (will check that it has one)
        with the HTML header injected into it.

        :param prepared: Determines whether to be using the standard HTML
            (``False``) or the prepared HTML (``True``) body. (Default: ``False``)

        :raises AttributeError: The correct HTML body cannot be acquired.
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
        if not validateHtml(body, self._htmlEncoding):
            logger.warning('HTML body failed to validate. Code will attempt to correct it.')

            # If we are here, then we need to do what we can to fix the HTML
            # body. Unfortunately this gets complicated because of the various
            # ways the body could be wrong. If only the <body> tag is missing,
            # then we just need to insert it at the end and be done. If both
            # the <html> and <body> tag are missing, we determine where to put
            # the body tag (around everything if there is no <head> tag,
            # otherwise at the end) and then wrap it all in the <html> tag.
            parser = bs4.BeautifulSoup(body, features = 'html.parser', from_encoding = self._htmlEncoding)
            self._getHtmlEncoding(parser)
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
            Internal function to replace the body tag with itself plus the
            header.
            """
            # I recently had to change this and how it worked. Now we use a new
            # property of `MSGFile` that returns a special tuple of tuples to define
            # how to get all of the properties we are formatting. They are all
            # processed in the same way, making everything neat. By defining them
            # in each class, any class can specify a completely different set to be
            # used.
            return bodyMarker.group() + self.htmlInjectableHeader.encode('utf-8')

        # Use the previously defined function to inject the HTML header.
        return constants.re.HTML_BODY_START.sub(replace, body, 1)

    def injectRtfHeader(self) -> bytes:
        """
        Returns the RTF body from this MSG file (will check that it has one)
        with the RTF header injected into it.

        :raises AttributeError: The RTF body cannot be acquired.
        :raises RuntimeError: All injection attempts failed.
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
            Internal function to replace the body tag with itself plus the
            header.
            """
            return bodyMarker.group() + injectableHeader

        # This first method only applies to documents with encapsulated HTML
        # that is formatted in a nice way.
        if isEncapsulatedRtf(self.rtfBody):
            data = constants.re.RTF_ENC_BODY_START.sub(replace, self.rtfBody, 1)
            if data != self.rtfBody:
                logger.debug('Successfully injected RTF header using encapsulation method.')
                return data
            logger.debug('RTF has encapsulated HTML, but injection method failed. It is likely dirty. Will use normal RTF injection method.')

        # If the normal encapsulated HTML injection fails or it isn't
        # encapsulated, use the internal _rtf module.
        logger.debug('Using _rtf module to inject RTF text header.')
        return createDocument(injectStartRTF(self.rtfBody, injectableHeader))

    def save(self, **kwargs) -> constants.SAVE_TYPE:
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
          has not been set, the name of the folder will be created using
          :property defaultFolderName:.
        * Setting :param maxNameLength: will force all file names to be
          shortened to fit in the space (with the extension included in the
          length). If a number is added to the directory that will not be
          included in the length, so it is recommended to plan for up to 5
          characters extra to be a part of the name. Default is 256.

        It should be noted that regardless of the value for
        :param maxNameLength:, the name of the file containing the body will always have the name 'message' followed by the full extension.

        There are several parameters used to determine how the message will be
        saved. By default, the message will be saved as plain text. Setting one
        of the following parameters to ``True`` will change that:

        * :param html: will output the message in HTML format.
        * :param json: will output the message in JSON format.
        * :param raw: will output the message in a raw format.
        * :param rtf: will output the message in RTF format.

        Usage of more than one formatting parameter will raise an exception.

        Using HTML or RTF will raise an exception if they could not be retrieved
        unless you have :param allowFallback: set to ``True``. Fallback will go
        in this order, starting at the top most format that is set:

        * HTML
        * RTF
        * Plain text

        If you want to save the contents into a ``ZipFile`` or similar object,
        either pass a path to where you want to create one or pass an instance
        to :param zip:. If :param zip: is set, :param customPath: will refer to
        a location inside the zip file.

        :param attachmentsOnly: Turns off saving the body and only saves the
            attachments when set.
        :param saveHeader: Turns on saving the header as a separate file when
            set.
        :param skipAttachments: Turns off saving attachments.
        :param skipHidden: If ``True``, skips attachments marked as hidden.
            (Default: ``False``)
        :param skipBodyNotFound: Suppresses errors if no valid body could be
            found, simply skipping the step of saving the body.
        :param charset: If the HTML is being prepared, the charset to use for
            the Content-Type meta tag to insert. This exists to ensure that
            something parsing the HTML can properly determine the encoding (as
            not having this tag can cause errors in some programs). Set this to
            ``None`` or an empty string to not insert the tag (Default:
            ``'utf-8'``).
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

        # Try to get the body, if needed, before messing with the path.
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

        createdZip = False
        try:
            # ZipFile handling.
            if _zip:
                # `raw` and `zip` are incompatible.
                if raw:
                    raise IncompatibleOptionsError('The options `raw` and `zip` are incompatible.')
                # If we are doing a zip file, first check that we have been
                # given a path.
                if isinstance(_zip, (str, pathlib.Path)):
                    # If we have a path then we use the zip file.
                    _zip = zipfile.ZipFile(_zip, 'a', zipfile.ZIP_DEFLATED)
                    kwargs['zip'] = _zip
                    createdZip = True
                # Path needs to be done in a special way if we are in a zip
                # file.
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

            # TODO: insert code here that will handle checking all of the msg
            # files to see if the path with overflow.

            if customFilename:
                # First we need to validate it. If there are invalid characters,
                # this will detect it.
                if constants.re.INVALID_FILENAME_CHARS.search(customFilename):
                    raise ValueError('Invalid character found in customFilename. Must not contain any of the following characters: \\/:*?"<>|')
                # Quick fix to remove spaces from the end of the filename, if
                # any are there.
                customFilename = customFilename.strip()
                path /= customFilename[:maxNameLength]
            elif useMsgFilename:
                if not self.filename:
                    raise ValueError(':param useMsgFilename: is only available if you are using an MSG file on the disk or have provided a filename.')
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
                        raise OSError(f'Failed to create directory "{path}". Does it already exist?')
            else:
                # In my testing I ended up with multiple files in a zip at the
                # same location so let's try to handle that.
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
                return (SaveType.FOLDER, str(path))

            # If the user has requested the headers for this file, save it now.
            if kwargs.get('saveHeader', False):
                headerText = self.headerText
                if not headerText:
                    headerText = constants.HEADER_FORMAT.format(subject = self.subject, **self.header)

                with _open(str(path / 'header.txt'), mode) as f:
                    f.write(headerText.encode('utf-8'))


            if not skipAttachments:
                # Save the attachments.
                attachmentReturns = [attachment.save(**kwargs) for attachment in self.attachments if not (skipHidden and attachment.hidden)]
                # Get the names from each.
                attachmentNames = []
                for x in attachmentReturns:
                    if isinstance(x[1], str):
                        attachmentNames.append(x[1])
                    elif isinstance(x[1], list):
                        attachmentNames.extend(x[1])

            if not attachOnly and fext:
                with _open(str(path / ('message.' + fext)), mode) as f:
                    if _json:
                        emailObj = json.loads(self.getJson())
                        if not skipAttachments:
                            emailObj['attachments'] = attachmentNames

                        f.write(json.dumps(emailObj).encode('utf-8'))
                    elif useHtml:
                        f.write(self.getSaveHtmlBody(**kwargs))
                    elif usePdf:
                        f.write(self.getSavePdfBody(**kwargs))
                    elif useRtf:
                        f.write(self.getSaveRtfBody(**kwargs))
                    else:
                        f.write(self.getSaveBody(**kwargs))

            return (SaveType.FOLDER, str(path))
        finally:
            # Close the ZipFile if this function created it.
            if _zip and createdZip:
                _zip.close()

    @functools.cached_property
    def bcc(self) -> Optional[str]:
        """
        The "Bcc" field, if it exists.
        """
        return self._genRecipient('bcc', RecipientType.BCC)

    @functools.cached_property
    def body(self) -> Optional[str]:
        """
        The message body, if it exists.
        """
        # If the body exists but is empty, that means it should be returned.
        if (body := self.getStringStream('__substg1.0_1000')) is not None:
            pass
        elif self.rtfBody:
            # If the body doesn't exist, see if we can get it from the RTF
            # body.
            body = self.deencapsulateBody(self.rtfBody, DeencapType.PLAIN)

        if body:
            body = inputToString(body, 'utf-8')
            if re.search('\n', body) is not None:
                if re.search('\r\n', body) is not None:
                    self._crlf = '\r\n'

        return body

    @functools.cached_property
    def cc(self) -> Optional[str]:
        """
        The "Cc" field, if it exists.
        """
        return self._genRecipient('cc', RecipientType.CC)

    @functools.cached_property
    def compressedRtf(self) -> Optional[bytes]:
        """
        The compressed RTF stream, if it exists.
        """
        return self.getStream('__substg1.0_10090102')

    @property
    def crlf(self) -> str:
        """
        The value of ``self.__crlf``, should you need it for whatever reason.
        """
        return self._crlf

    @functools.cached_property
    def date(self) -> Optional[datetime.datetime]:
        """
        The send date, if it exists.
        """
        return self.props.date if self.isSent else None

    @functools.cached_property
    def deencapsulatedRtf(self) -> Optional[RTFDE.DeEncapsulator]:
        """
        The instance of the deencapsulated RTF body.

        If there is no RTF body or the body is not encasulated, returns
        ``None``.
        """
        if self.rtfBody:
            # If there is an RTF body, we try to deencapsulate it.
            body = self.rtfBody
            # Sometimes you get MSG files whose RTF body has stuff
            # *after* the body, and RTFDE can't handle that. Here is
            # how we compensate.
            while body and body[-1] != 125:
                body = body[:-1]

            # Some files take a long time due to how they are structured and
            # how RTFDE works. The longer a file would normally take, the
            # better this fix works:
            body = stripRtf(body)

            try:
                deencapsultor = RTFDE.DeEncapsulator(body)
                deencapsultor.deencapsulate()
                return deencapsultor
            except RTFDE.exceptions.NotEncapsulatedRtf:
                logger.debug('RTF body is not encapsulated.')
            except RTFDE.exceptions.MalformedEncapsulatedRtf:
                if ErrorBehavior.RTFDE_MALFORMED not in self.errorBehavior:
                    raise
                logger.info('RTF body contains malformed encapsulated content.')
            except Exception:
                # If we are just ignoring the errors, log it then set to
                # None. Otherwise, continue the exception.
                if ErrorBehavior.RTFDE_UNKNOWN_ERROR not in self.errorBehavior:
                    raise
                logger.exception('Unhandled error happened while using RTFDE. You have choosen to ignore these errors.')
        return None

    @property
    def defaultFolderName(self) -> str:
        """
        Generates the default name of the save folder.
        """
        try:
            return self._defaultFolderName
        except AttributeError:
            d = self.parsedDate or tuple([0] * 9)

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

    @functools.cached_property
    def header(self) -> email.message.Message:
        """
        The message header, if it exists.

        If one does not exist as a stream, it will be generated from the other
        properties.
        """
        headerText = self.headerText
        if headerText:
            # Fix an issue with prefixed headers not parsing correctly.
            if headerText.startswith('Microsoft Mail Internet Headers Version 2.0'):
                headerText = headerText[43:].lstrip()
            header = HeaderParser(policy = policy.compat32).parsestr(headerText)
        else:
            logger.info('Header is empty or was not found. Header will be generated from other streams.')
            header = HeaderParser(policy = policy.compat32).parsestr('')
            if self.date:
                header.add_header('Date', email.utils.format_datetime(self.date))
            header.add_header('From', self.sender)
            header.add_header('To', self.to)
            header.add_header('Cc', self.cc)
            header.add_header('Bcc', self.bcc)
            header.add_header('Message-Id', self.messageId)
            # TODO find authentication results outside of header
            header.add_header('Authentication-Results', None)

        self.__headerInit = True
        return header

    @functools.cached_property
    def headerDict(self) -> Dict[str, Any]:
        """
        A dictionary of the entries in the header
        """
        headerDict = {x: self.header[x] for x in self.header}
        try:
            headerDict.pop('Received')
        except KeyError:
            pass
        return headerDict

    @property
    def headerFormatProperties(self) -> constants.HEADER_FORMAT_TYPE:
        """
        Returns a dictionary of properties, in order, to be formatted into the
        header.

        Keys are the names to use in the header while the values are one of the
        following:

        * ``None``: Signifies no data was found for the property and it should
          be omitted from the header.
        * ``str``: A string to be formatted into the header using the string
          encoding.
        * ``Tuple[Union[str, None], bool]``: A string should be formatted into
          the header. If the bool is ``True``, then place an empty string if
          the first value is ``None``, otherwise follow the same behavior as
          regular ``None``.

        Additional note: If the value is an empty string, it will be dropped as
        well by default.

        Additionally you can group members of a header together by placing them
        in an embedded dictionary. Groups will be spaced out using a second
        instance of the join string. If any member of a group is being printed,
        it will be spaced apart from the next group/item.

        If you class should not do *any* header injection, return ``None`` from
        this property.
        """
        # Checking outlook printing, default behavior is to completely omit
        # *any* field that is not present. So while for extensability the
        # option exists to have it be present even if no data is found, we are
        # specifically not doing that.
        return {
            '-basic info-': {
                'From': self.sender,
                'Sent': self.date.__format__(self.datetimeFormat) if self.date else None,
                'To': self.to,
                'Cc': self.cc,
                'Bcc': self.bcc,
                'Subject': self.subject,
            },
            '-importance-': {
                'Importance': self.importanceString,
            },
        }

    @property
    def headerInit(self) -> bool:
        """
        Checks whether the header has been initialized.
        """
        return self.__headerInit

    @functools.cached_property
    def headerText(self) -> Optional[str]:
        """
        The raw text of the header stream, if it exists.
        """
        return self.getStringStream('__substg1.0_007D')

    @functools.cached_property
    def htmlBody(self) -> Optional[bytes]:
        """
        The HTML body, if it exists.
        """
        if (htmlBody := self.getStream('__substg1.0_10130102')) is not None:
            pass
        elif self.rtfBody:
            logger.info('HTML body was not found, attempting to generate from RTF.')
            htmlBody = cast(bytes, self.deencapsulateBody(self.rtfBody, DeencapType.HTML))
        # This is it's own if statement so we can ensure it will generate
        # even if there is an rtfBody, in the event it doesn't have HTML.
        if not htmlBody and self.body:
            # Convert the plain text body to html.
            logger.info('HTML body was not found, attempting to generate from plain text body.')
            correctedBody = html.escape(self.body).replace('\r', '').replace('\n', '<br />')
            htmlBody = f'<html><body>{correctedBody}</body></head>'.encode('ascii', 'xmlcharrefreplace')

        if not htmlBody:
            logger.info('HTML body could not be found nor generated.')

        return htmlBody

    @functools.cached_property
    def htmlBodyPrepared(self) -> Optional[bytes]:
        """
        The HTML body that has (where possible) the embedded attachments
        inserted into the body.
        """
        # If we can't get an HTML body then we have nothing to do.
        if not self.htmlBody:
            return self.htmlBody

        # Create the BeautifulSoup instance to use.
        soup = bs4.BeautifulSoup(self.htmlBody, 'html.parser', from_encoding = self._htmlEncoding)
        self._getHtmlEncoding(soup)

        # Get a list of image tags to see if we can inject into. If the source
        # of an image starts with "cid:" that means it is one of the attachments
        # and is using the content id of that attachment.
        tags = (tag for tag in soup.findAll('img') if tag.get('src') and tag.get('src').startswith('cid:'))

        for tag in tags:
            # Iterate through the attachments until we get the right one.
            cid = tag['src'][4:]
            att = next((attachment for attachment in self.attachments if hasattr(attachment, 'cid') and attachment.cid == cid), None)
            # If we found anything, inject it.
            if att and isinstance(att.data, bytes):
                # Try to get the mimetype. If we can't, see if the item has an
                # extension and guess the mimtype for a few known ones.
                mime = att.mimetype
                if not mime:
                    ext = (att.name or '').split('.')[-1].lower()
                    if ext == 'png':
                        mime = 'image/png'
                    elif ext == 'jpg' or ext == 'jpeg':
                        mime = 'image/jpeg'
                    elif ext == 'gif':
                        mime = 'image/gif'
                    elif ext == 'tiff' or ext == 'tif':
                        mime = 'image/tif'
                    elif ext == 'bmp':
                        mime = 'image/bmp'
                    elif ext == 'svg':
                        mime = 'image/svg+xml'
                # Final check.
                if mime:
                    tag['src'] = (b'data:' + mime.encode() + b';base64,' + base64.b64encode(att.data)).decode('utf-8')
                else:
                    # We don't know what to actually put for this item, and we
                    # really should never end up here, so throw an error.
                    raise MimetypeFailureError('Could not get the mimetype to use for htmlBodyPrepared.')

        return soup.encode('utf-8')

    @functools.cached_property
    def htmlInjectableHeader(self) -> str:
        """
        The header that can be formatted and injected into the HTML body.
        """
        prefix = '<div id="injectedHeader"><div><p class="MsoNormal">'
        suffix = '<o:p></o:p></p></div></div>'
        joinStr = '<br/>'
        formatter = (lambda name, value: f'<b>{name}:</b>&nbsp;{inputToString(htmlSanitize(value), self.stringEncoding).encode("ascii", "xmlcharrefreplace").decode()}')

        return self.getInjectableHeader(prefix, joinStr, suffix, formatter)

    @functools.cached_property
    def inReplyTo(self) -> Optional[str]:
        """
        The message id that this message is in reply to.
        """
        return self.getStringStream('__substg1.0_1042')

    @functools.cached_property
    def isRead(self) -> bool:
        """
        Whether this email has been marked as read.
        """
        return bool(self.getPropertyVal('0E070003', 0) & 1)

    @functools.cached_property
    def isSent(self) -> bool:
        """
        Whether this email has been marked as sent.

        Assumes ``True`` if no flags are found.
        """
        return not bool(self.getPropertyVal('0E070003', 0) & 8)

    @functools.cached_property
    def messageId(self) -> Optional[str]:
        headerResult = None
        if self.headerInit:
            headerResult = self.header['message-id']
        if headerResult is not None:
            return headerResult

        if self.headerInit:
            logger.info('Header found, but "Message-Id" is not included. Will be generated from other streams.')
        return self.getStringStream('__substg1.0_1035')

    @functools.cached_property
    def parsedDate(self) -> Optional[Tuple[int, int, int, int, int, int, int, int, int]]:
        """
        A 9 tuple of the parsed date from the header.
        """
        return email.utils.parsedate(self.header['Date'])

    @functools.cached_property
    def receivedTime(self) -> Optional[datetime.datetime]:
        """
        The date and time the message was received by the server.
        """
        return self.getPropertyVal('0E060040')

    @property
    def recipientSeparator(self) -> str:
        return self.__recipientSeparator

    @functools.cached_property
    def recipients(self) -> List[Recipient]:
        """
        A list of all recipients.
        """
        recipientDirs = []
        prefixLen = self.prefixLen
        for dir_ in self.listDir():
            if dir_[prefixLen].startswith('__recip') and\
                    dir_[prefixLen] not in recipientDirs:
                recipientDirs.append(dir_[prefixLen])

        return [Recipient(recipientDir, self, self.recipientTypeClass) for recipientDir in recipientDirs]

    @property
    def recipientTypeClass(self) -> Type[enum.IntEnum]:
        """
        The class to use for a recipient's recipientType property.

        The default is :class:`extract_msg.enums.RecipientType`. If a subclass's
        attributes have different meanings to the values, you can override this
        property to return a valid enum.
        """
        return RecipientType

    @functools.cached_property
    def reportTag(self) -> Optional[ReportTag]:
        """
        Data that is used to correlate the report and the original message.
        """
        return self.getStreamAs('__substg1.0_00310102', ReportTag)

    @functools.cached_property
    def responseRequested(self) -> bool:
        """
        Whether to send Meeting Response objects to the organizer.
        """
        return bool(self.getPropertyVal('0063000B'))

    @functools.cached_property
    def rtfBody(self) -> Optional[bytes]:
        """
        The decompressed Rtf body from the message.
        """
        return compressed_rtf.decompress(self.compressedRtf) if self.compressedRtf else None

    @functools.cached_property
    def rtfEncapInjectableHeader(self) -> bytes:
        """
        The header that can be formatted and injected into the plain RTF body.
        """
        prefix = r'\htmlrtf {\htmlrtf0 {\*\htmltag96 <div>}{\*\htmltag96 <div>}{\*\htmltag64 <p class=MsoNormal>}'
        suffix = r'{\*\htmltag244 <o:p>}{\*\htmlrag252 </o:p>}\htmlrtf \par\par\htmlrtf0 {\*\htmltag72 </p>}{\*\htmltag104 </div>}{\*\htmltag104 </div>}\htmlrtf }\htmlrtf0 '
        joinStr = r'{\*\htmltag116 <br />}\htmlrtf \line\htmlrtf0 '
        formatter = (lambda name, value: fr'\htmlrtf {{\b\htmlrtf0{{\*\htmltag84 <b>}}{name}: {{\*\htmltag92 </b>}}\htmlrtf \b0\htmlrtf0 {inputToString(rtfSanitizeHtml(value), self.stringEncoding)}\htmlrtf }}\htmlrtf0')

        return self.getInjectableHeader(prefix, joinStr, suffix, formatter).encode('utf-8')

    @functools.cached_property
    def rtfPlainInjectableHeader(self) -> bytes:
        """
        The header that can be formatted and injected into the encapsulated RTF
        body.
        """
        prefix = '{'
        suffix = '\\par\\par}'
        joinStr = '\\line'
        formatter = (lambda name, value: fr'{{\b {name}: \b0 {inputToString(rtfSanitizePlain(value), self.stringEncoding)}}}')

        return self.getInjectableHeader(prefix, joinStr, suffix, formatter).encode('utf-8')

    @functools.cached_property
    def sender(self) -> Optional[str]:
        """
        The message sender, if it exists.
        """
        # Check header first
        if self.headerInit:
            headerResult = self.header['from']
            if headerResult is not None:
                return decodeRfc2047(headerResult)
            logger.info('Header found, but "sender" is not included. Will be generated from other streams.')
        # Extract from other fields
        text = self.getStringStream('__substg1.0_0C1A')
        email = self.getStringStream('__substg1.0_5D01')
        # Will not give an email address sometimes. Seems to exclude the email
        # address if YOU are the sender.
        result = None
        if text is None:
            result = email
        else:
            result = text
            if email is not None:
                result += ' <' + email + '>'

        return result

    @functools.cached_property
    def subject(self) -> Optional[str]:
        """
        The message subject, if it exists.
        """
        return self.getStringStream('__substg1.0_0037')

    @functools.cached_property
    def to(self) -> Optional[str]:
        """
        The "To" field, if it exists.
        """
        return self._genRecipient('to', RecipientType.TO)
