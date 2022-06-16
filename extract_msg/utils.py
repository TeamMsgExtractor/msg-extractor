"""
Utility functions of extract_msg.
"""

import argparse
import codecs
import collections
import copy
import datetime
import email.message
import email.policy
import glob
import json
import logging
import logging.config
import os
import pathlib
import shutil
import struct
# Not actually sure if this needs to be here for the logging, so just in case.
import sys
import zipfile

import bs4
import tzlocal

from html import escape as htmlEscape
from typing import Dict, List, Tuple, Union

from . import constants
from .enums import AttachmentType
from .exceptions import BadHtmlError, ConversionError, IncompatibleOptionsError, InvalidFileFormatError, InvaildPropertyIdError, UnknownCodepageError, UnknownTypeError, UnrecognizedMSGTypeError, UnsupportedMSGTypeError


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())
logging.addLevelName(5, 'DEVELOPER')


def addNumToDir(dirName : pathlib.Path) -> pathlib.Path:
    """
    Attempt to create the directory with a '(n)' appended.
    """
    for i in range(2, 100):
        try:
            newDirName = dirName.with_name(dirName.name + f' ({i})')
            os.makedirs(newDirName)
            return newDirName
        except Exception as e:
            pass
    return None


def addNumToZipDir(dirName : pathlib.Path, _zip):
    """
    Attempt to create the directory with a '(n)' appended.
    """
    for i in range(2, 100):
        newDirName = dirName.with_name(dirName.name + f' ({i})')
        pathCompare = str(newDirName).rstrip('/') + '/'
        if not any(x.startswith(pathCompare) for x in _zip.namelist()):
            return newDirName
    return None


def bitwiseAdjust(inp : int, mask : int) -> int:
    """
    Uses a given mask to adjust the location of bits after an operation like
    bitwise AND. This is useful for things like flags where you are trying to
    get a small portion of a larger number. Say for example, you had the number
    0xED (0b11101101) and you needed the adjusted result of the AND operation
    with 0x70 (0b01110000). The result of the AND operation (0b01100000) and the
    mask used to get it (0x70) are given to this function and the adjustment
    will be done automatically.

    :param mask: MUST be greater than 0.

    :raises ValueError: if the mask is not greater than 0.
    """
    if mask < 1:
        raise ValueError('Mask MUST be greater than 0')
    return inp >> bin(mask)[::-1].index('1')


def bitwiseAdjustedAnd(inp : int, mask : int) -> int:
    """
    Preforms the bitwise AND operation between :param inp: and :param mask: and
    adjusts the results based on the rules of the bitwiseAdjust function.

    :raises ValueError: if the mask is not greater than 0.
    """
    if mask < 1:
        raise ValueError('Mask MUST be greater than 0')
    return (inp & mask) >> bin(mask)[::-1].index('1')


def bytesToGuid(bytesInput : bytes) -> str:
    """
    Converts a bytes instance to a GUID.
    """
    guidVals = constants.ST_GUID.unpack(bytesInput)
    return f'{{{guidVals[0]:08X}-{guidVals[1]:04X}-{guidVals[2]:04X}-{guidVals[3][:2].hex().upper()}-{guidVals[3][2:].hex().upper()}}}'


def ceilDiv(n : int, d : int) -> int:
    """
    Returns the int from the ceil division of n / d.
    ONLY use ints as inputs to this function.

    For ints, this is faster and more accurate for numbers
    outside the precision range of float.
    """
    return -(n // -d)


def createZipOpen(func):
    """
    Creates a wrapper for the open function of a ZipFile that will automatically
    set the current date as the modified time to the current time.
    """
    def _open(name, mode, *args, **kwargs):
        if mode == 'w':
            name = zipfile.ZipInfo(name, datetime.datetime.now().timetuple())

        return func(name, mode, *args, **kwargs)

    return _open


def divide(string, length : int) -> list:
    """
    Divides a string into multiple substrings of equal length. If there is not
    enough for the last substring to be equal, it will simply use the rest of
    the string. Can also be used for things like lists and tuples.

    :param string: string to be divided.
    :param length: length of each division.
    :returns: list containing the divided strings.

    Example:
    >>>> a = divide('Hello World!', 2)
    >>>> print(a)
    ['He', 'll', 'o ', 'Wo', 'rl', 'd!']
    >>>> a = divide('Hello World!', 5)
    >>>> print(a)
    ['Hello', ' Worl', 'd!']
    """
    return [string[length * x:length * (x + 1)] for x in range(int(ceilDiv(len(string), length)))]


def findWk(path = None):
    """
    Attempt to find the path of the wkhtmltopdf executable. If :param path: is
    provided, verifies that it is executable and returns the path if it is.

    :raises ExecutableNotFound: A valid executable could not be found.
    """
    if path:
        if os.path.isfile(path):
            # Check if executable.
            if os.access(path, os.X_OK):
                return path
            else:
                raise ExecutableNotFound('Path provided was not a valid executable (execution bit not set).')
        else:
            raise ExecutableNotFound('Path provided was not a valid executable (not a file).')

    candidate = shutil.which('wkhtmltopdf')
    if candidate:
        return candidate

    raise ExecutableNotFound('Could not find wkhtmltopdf.')


def fromTimeStamp(stamp) -> datetime.datetime:
    """
    Returns a datetime from the UTC timestamp given the current timezone.
    """
    return datetime.datetime.fromtimestamp(stamp, tzlocal.get_localzone())


def getCommandArgs(args):
    """
    Parse command-line arguments.

    :raises IncompatibleOptionsError: if options are provided that are
        incompatible.
    """
    parser = argparse.ArgumentParser(description = constants.MAINDOC, prog = 'extract_msg')
    outFormat = parser.add_mutually_exclusive_group()
    inputFormat = parser.add_mutually_exclusive_group()
    # --use-content-id, --cid
    parser.add_argument('--use-content-id', '--cid', dest='cid', action='store_true',
                        help='Save attachments by their Content ID, if they have one. Useful when working with the HTML body.')
    # --dev
    parser.add_argument('--dev', dest='dev', action='store_true',
                        help='Changes to use developer mode. Automatically enables the --verbose flag. Takes precedence over the --validate flag.')
    # --validate
    parser.add_argument('--validate', dest='validate', action='store_true',
                        help='Turns on file validation mode. Turns off regular file output.')
    # --json
    outFormat.add_argument('--json', dest='json', action='store_true',
                        help='Changes to write output files as json.')
    # --file-logging
    parser.add_argument('--file-logging', dest='fileLogging', action='store_true',
                        help='Enables file logging. Implies --verbose.')
    # --verbose
    parser.add_argument('--verbose', dest='verbose', action='store_true',
                        help='Turns on console logging.')
    # --log PATH
    parser.add_argument('--log', dest='log',
                        help='Set the path to write the file log to.')
    # --config PATH
    parser.add_argument('--config', dest='configPath',
                        help='Set the path to load the logging config from.')
    # --out PATH
    parser.add_argument('--out', dest='outPath',
                        help='Set the folder to use for the program output. (Default: Current directory)')
    # --use-filename
    parser.add_argument('--use-filename', dest='useFilename', action='store_true',
                        help='Sets whether the name of each output is based on the msg filename.')
    # --dump-stdout
    parser.add_argument('--dump-stdout', dest='dumpStdout', action='store_true',
                        help='Tells the program to dump the message body (plain text) to stdout. Overrides saving arguments.')
    # --html
    outFormat.add_argument('--html', dest='html', action='store_true',
                        help='Sets whether the output should be HTML. If this is not possible, will error.')
    # --pdf
    outFormat.add_argument('--pdf', dest='pdf', action='store_true',
                        help='Saves the body as a PDF. If this is not possible, will error.')
    # --wk-path PATH
    parser.add_argument('--wk-path', dest='wkPath',
                        help='Overrides the path for finding wkhtmltopdf.')
    # --wk-options OPTIONS
    parser.add_argument('--wk-options', dest='wkOptions', nargs='*',
                        help='Sets additional options to be used in wkhtmltopdf. Should be a series of options and values, replacing the - or -- in the beginning with + or ++, respectively. For example: --wk-options "+O Landscape"')
    # --prepared-html
    parser.add_argument('--prepared-html', dest='preparedHtml', action='store_true',
                        help='When used in conjunction with --html, sets whether the HTML output should be prepared for embedded attachments.')
    # --charset
    parser.add_argument('--charset', dest='charset', default='utf-8',
                        help='Character set to use for the prepared HTML in the added tag. (Default: utf-8)')
    # --raw
    outFormat.add_argument('--raw', dest='raw', action='store_true',
                        help='Sets whether the output should be raw. If this is not possible, will error.')
    # --rtf
    outFormat.add_argument('--rtf', dest='rtf', action='store_true',
                        help='Sets whether the output should be RTF. If this is not possible, will error.')
    # --allow-fallback
    parser.add_argument('--allow-fallback', dest='allowFallback', action='store_true',
                        help='Tells the program to fallback to a different save type if the selected one is not possible.')
    # --zip
    parser.add_argument('--zip', dest='zip',
                        help='Path to use for saving to a zip file.')
    # --attachments-only
    outFormat.add_argument('--attachments-only', dest='attachmentsOnly', action='store_true',
                        help='Specify to only save attachments from an msg file.')
    # --out-name NAME
    inputFormat.add_argument('--out-name', dest='outName',
                        help='Name to be used with saving the file output. Cannot be used if you are saving more than one file.')
    # --glob
    inputFormat.add_argument('--glob', '--wildcard', dest='glob', action='store_true',
                        help='Interpret all paths as having wildcards. Incompatible with --out-name.')
    # --ignore-rtfde
    parser.add_argument('--ignore-rtfde', dest='ignoreRtfDeErrors', action='store_true',
                        help='Ignores all errors thrown from RTFDE when trying to save. Useful for allowing fallback to continue when an exception happens.')
    # --progress
    parser.add_argument('--progress', dest='progress', action='store_true',
                        help='Shows what file the program is currently working on during it\'s progress.')
    # [msg files]
    parser.add_argument('msgs', metavar='msg', nargs='+',
                        help='An MSG file to be parsed.')

    options = parser.parse_args(args)

    if options.dev or options.fileLogging:
        options.verbose = True

    # Handle the wkOptions if they exist.
    if options.wkOptions:
        wkOptions = []
        for option in options.wkOptions:
            if option.startswith('++'):
                option = '--' + option[2:]
            elif option.startswith('+'):
                option = '-' + option[1:]

            # Now that we have corrected to the correct start, split the argument if
            # necessary.
            split = option.split(' ')
            if len(split) == 1:
                # No spaces means we just pass that directly.
                wkOptions.append(option)
            else:
                wkOptions.append(split[0])
                wkOptions.append(' '.join(split[1:]))

        options.wkOptions = wkOptions

    # If dump_stdout is True, we need to unset all arguments used in files.
    # Technically we actually only *need* to unset `out_path`, but that may
    # change in the future, so let's be thorough.
    if options.dumpStdout:
        options.outPath = None
        options.json = False
        options.rtf = False
        options.html = False
        options.useFilename = False
        options.cid = False

    if options.glob:
        fileLists = []
        for path in options.msgs:
            fileLists += glob.glob(path)

        if len(fileLists) == 0:
            raise ValueError('Could not find any msg files using the specified wildcards.')
        options.msgs = fileLists

    # Make it so outName can only be used on single files.
    if options.outName and options.fileArgs and len(options.fileArgs) > 0:
        raise ValueError('--out-name is not supported when saving multiple MSG files.')

    return options


def getEncodingName(codepage : int) -> str:
    """
    Returns the name of the encoding with the specified codepage.

    :raises UnknownCodepageError: if the codepage is unrecognized.
    :raises UnsupportedEncodingError: if the codepage is not supported.
    """
    if codepage not in constants.CODE_PAGES:
        raise UnknownCodepageError(str(codepage))
    try:
        codecs.lookup(constants.CODE_PAGES[codepage])
        return constants.CODE_PAGES[codepage]
    except LookupError:
        raise UnsupportedEncodingError(f'The codepage {codepage} ({constants.CODE_PAGES[codepage]}) is not currently supported by your version of Python.')


def getFullClassName(inp):
    return inp.__class__.__module__ + '.' + inp.__class__.__name__


def hasLen(obj) -> bool:
    """
    Checks if :param obj: has a __len__ attribute.
    """
    return hasattr(obj, '__len__')


def injectHtmlHeader(msgFile, prepared : bool = False) -> bytes:
    """
    Returns the HTML body from the MSG file (will check that it has one) with
    the HTML header injected into it.

    :param prepared: Determines whether to be using the standard HTML (False) or
        the prepared HTML (True) body (Default: False).

    :raises AttributeError: if the correct HTML body cannot be acquired.
    :raises BadHtmlError: if :param preparedHtml: is False and the HTML fails to
        validate.
    """
    if not hasattr(msgFile, 'htmlBody') or not msgFile.htmlBody:
        raise AttributeError('Cannot inject the HTML header without an HTML body attribute.')

    body = None

    # We don't do this all at once because the prepared body is not cached.
    if prepared:
        if hasattr(msgFile, 'htmlBodyPrepared'):
            body = msgFile.htmlBodyPrepared

        # If the body is not valid or not found, raise an AttributeError.
        if not body:
            raise AttributeError('Cannot find a prepared HTML body to inject into.')
    else:
        body = msgFile.htmlBody

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
        return bodyMarker.group() + constants.HTML_INJECTABLE_HEADER.format(
        **{
            'sender': inputToString(htmlEscape(msgFile.sender) if msgFile.sender else '', 'utf-8'),
            'to': inputToString(htmlEscape(msgFile.to) if msgFile.to else '', 'utf-8'),
            'cc': inputToString(htmlEscape(msgFile.cc) if msgFile.cc else '', 'utf-8'),
            'bcc': inputToString(htmlEscape(msgFile.bcc) if msgFile.bcc else '', 'utf-8'),
            'date': inputToString(msgFile.date, 'utf-8'),
            'subject': inputToString(htmlEscape(msgFile.subject) if msgFile.subject else '', 'utf-8'),
        }).encode('utf-8')

    # Use the previously defined function to inject the HTML header.
    return constants.RE_HTML_BODY_START.sub(replace, body, 1)


def injectRtfHeader(msgFile) -> bytes:
    """
    Returns the RTF body from the MSG file (will check that it has one) with the
    RTF header injected into it.

    :raises AttributeError: if the RTF body cannot be acquired.
    :raises RuntimeError: if all injection attempts fail.
    """
    if not hasattr(msgFile, 'rtfBody') or not msgFile.rtfBody:
        raise AttributeError('Cannot inject the RTF header without an RTF body attribute.')

    # Try to determine which header to use. Also determines how to sanitize the
    # rtf.
    if isEncapsulatedRtf(msgFile.rtfBody):
        injectableHeader = constants.RTF_ENC_INJECTABLE_HEADER
        rtfSanitize = rtfSanitizeHtml
    else:
        injectableHeader = constants.RTF_PLAIN_INJECTABLE_HEADER
        rtfSanitize = rtfSanitizePlain

    def replace(bodyMarker):
        """
        Internal function to replace the body tag with itself plus the header.
        """
        return bodyMarker.group() + injectableHeader.format(
        **{
            'sender': inputToString(rtfSanitize(msgFile.sender) if msgFile.sender else '', 'utf-8'),
            'to': inputToString(rtfSanitize(msgFile.to) if msgFile.to else '', 'utf-8'),
            'cc': inputToString(rtfSanitize(msgFile.cc) if msgFile.cc else '', 'utf-8'),
            'bcc': inputToString(rtfSanitize(msgFile.bcc) if msgFile.bcc else '', 'utf-8'),
            'date': inputToString(msgFile.date, 'utf-8'),
            'subject': inputToString(rtfSanitize(msgFile.subject), 'utf-8'),
        }).encode('utf-8')

    # Use the previously defined function to inject the RTF header. We are
    # trying a few different methods to determine where to place the header.
    data = constants.RE_RTF_BODY_START.sub(replace, msgFile.rtfBody, 1)
    # If after any method the data does not match the RTF body, then we have
    # succeeded.
    if data != msgFile.rtfBody:
        logger.debug('Successfully injected RTF header using first method.')
        return data

    # This second method only applies to encapsulated HTML, so we need to check
    # for that first.
    if isEncapsulatedRtf(msgFile.rtfBody):
        data = constants.RE_RTF_ENC_BODY_START_1.sub(replace, msgFile.rtfBody, 1)
        if data != msgFile.rtfBody:
            logger.debug('Successfully injected RTF header using second method.')
            return data

        # This third method is a lot less reliable, and actually would just
        # simply violate the encapuslated html, so for this one we don't even
        # try to worry about what the html will think about it. If it injects,
        # we swap to basic and then inject again, more worried about it working
        # than looking nice inside.
        if constants.RE_RTF_ENC_BODY_UGLY.sub(replace, msgFile.rtfBody, 1) != msgFile.rtfBody:
            injectableHeader = constants.RTF_PLAIN_INJECTABLE_HEADER
            data = constants.RE_RTF_ENC_BODY_UGLY.sub(replace, msgFile.rtfBody, 1)
            logger.debug('Successfully injected RTF header using third method.')
            return data

    # Severe fallback attempts.
    data = constants.RE_RTF_BODY_FALLBACK_FS.sub(replace, msgFile.rtfBody, 1)
    if data != msgFile.rtfBody:
        logger.debug('Successfully injected RTF header using forth method.')
        return data

    data = constants.RE_RTF_BODY_FALLBACK_F.sub(replace, msgFile.rtfBody, 1)
    if data != msgFile.rtfBody:
        logger.debug('Successfully injected RTF header using fifth method.')
        return data

    data = constants.RE_RTF_BODY_FALLBACK_PLAIN.sub(replace, msgFile.rtfBody, 1)
    if data != msgFile.rtfBody:
        logger.debug('Successfully injected RTF header using sixth method.')
        return data

    raise RuntimeError('All injection attempts failed. Please report this to the developer.')


def inputToBytes(stringInputVar, encoding) -> bytes:
    """
    Converts the input into bytes.

    :raises ConversionError: if the input cannot be converted.
    """
    if isinstance(stringInputVar, bytes):
        return stringInputVar
    elif isinstance(stringInputVar, str):
        return stringInputVar.encode(encoding)
    elif stringInputVar is None:
        return b''
    else:
        raise ConversionError('Cannot convert to bytes.')


def inputToMsgpath(inp) -> list:
    """
    Converts the input into an msg path.
    """
    if isinstance(inp, (list, tuple)):
        inp = '/'.join(inp)
    ret = inputToString(inp, 'utf-8').replace('\\', '/').split('/')
    return ret if ret[0] != '' else []


def inputToString(bytesInputVar, encoding) -> str:
    """
    Converts the input into a string.

    :raises ConversionError: if the input cannot be converted.
    """
    if isinstance(bytesInputVar, str):
        return bytesInputVar
    elif isinstance(bytesInputVar, bytes):
        return bytesInputVar.decode(encoding)
    elif bytesInputVar is None:
        return ''
    else:
        raise ConversionError('Cannot convert to str type.')


def isEncapsulatedRtf(inp : bytes) -> bool:
    """
    Currently the destection is made to be *extremly* basic, but this will work
    for now. In the future this will be fixed to that literal text in the body
    of a message won't cause false detection.
    """
    return b'\\fromhtml' in inp


def isEmptyString(inp : str) -> bool:
    """
    Returns true if the input is None or is an Empty string.
    """
    return (inp == '' or inp is None)


def knownMsgClass(classType : str) -> bool:
    """
    Checks if the specified class type is recognized by the module. Usually used
    for checking if a type is simply unsupported rather than unknown.
    """
    classType = classType.lower()
    if classType == 'ipm':
        return True

    for item in constants.KNOWN_CLASS_TYPES:
        if classType.startswith(item):
            return True

    return False


def filetimeToUtc(inp : int) -> float:
    """
    Converts a FILETIME into a unix timestamp.
    """
    return (inp - 116444736000000000) / 10000000.0


def msgpathToString(inp) -> str:
    """
    Converts an msgpath (one of the internal paths inside an msg file) into a
    string.
    """
    if inp is None:
        return None
    if isinstance(inp, (list, tuple)):
        inp = '/'.join(inp)
    inp.replace('\\', '/')
    return inp


def openMsg(path, **kwargs):
    """
    Function to automatically open an MSG file and detect what type it is.

    :param path: Path to the msg file in the system or is the raw msg file.
    :param prefix: Used for extracting embeded msg files inside the main one.
        Do not set manually unless you know what you are doing.
    :param parentMsg: Used for syncronizing named properties instances. Do not
        set this unless you know what you are doing.
    :param attachmentClass: Optional, the class the Message object will use for
        attachments. You probably should not change this value unless you know
        what you are doing.
    :param signedAttachmentClass: Optional, the class the object will use for
        signed attachments.
    :param filename: Optional, the filename to be used by default when saving.
    :param delayAttachments: Optional, delays the initialization of attachments
        until the user attempts to retrieve them. Allows MSG files with bad
        attachments to be initialized so the other data can be retrieved.
    :param overrideEncoding: Optional, overrides the specified encoding of the
        MSG file.
    :param attachmentErrorBehavior: Optional, the behaviour to use in the event
        of an error when parsing the attachments.
    :param recipientSeparator: Optional, Separator string to use between
        recipients.
    :param ignoreRtfDeErrors: Optional, specifies that any errors that occur
        from the usage of RTFDE should be ignored (default: False).

    If :param strict: is set to `True`, this function will raise an exception
    when it cannot identify what MSGFile derivitive to use. Otherwise, it will
    log the error and return a basic MSGFile instance.

    :raises UnsupportedMSGTypeError: if the type is recognized but not suppoted.
    :raises UnrecognizedMSGTypeError: if the type is not recognized.
    """
    from .appointment import Appointment
    from .contact import Contact
    from .message import Message
    from .msg import MSGFile
    from .message_signed import MessageSigned
    from .task import Task

    msg = MSGFile(path, **kwargs)
    # After rechecking the docs, all comparisons should be case-insensitive, not
    # case-sensitive. My reading ability is great.
    #
    # Also after consideration, I realized we need to be very careful here, as
    # other file types (like doc, ppt, etc.) might open but not return a class
    # type. If the stream is not found, classType returns None, which has no
    # lower function. So let's make sure we got a good return first.
    if not msg.classType:
        if kwargs.get('strict', True):
            raise InvalidFileFormatError('File was confirmed to be an olefile, but was not an MSG file.')
        else:
            # If strict mode is off, we'll just return an MSGFile anyways.
            logging.critical('Received file that was an olefile but was not an MSG file. Returning MSGFile anyways because strict mode is off.')
            return msg
    classType = msg.classType.lower()
    if classType.startswith('ipm.contact') or classType.startswith('ipm.distlist'):
        msg.close()
        return Contact(path, **kwargs)
    elif classType.startswith('ipm.note') or classType.startswith('report'):
        msg.close()
        if classType.endswith('smime.multipartsigned'):
            return MessageSigned(path, **kwargs)
        else:
            return Message(path, **kwargs)
    elif classType.startswith('ipm.appointment') or classType.startswith('ipm.schedule'):
        msg.close()
        return Appointment(path, **kwargs)
    elif classType.startswith('ipm.task'):
        msg.close()
        return Task(path, **kwargs)
    elif classType == 'ipm': # Unspecified format. It should be equal to this and not just start with it.
        return msg
    elif kwargs.get('strict', True):
        # Because we are closing it, we need to store it in a variable first.
        ct = msg.classType
        msg.close()
        if knownMsgClass(classType):
            raise UnsupportedMSGTypeError(f'MSG type "{ct}" currently is not supported by the module. If you would like support, please make a feature request.')
        raise UnrecognizedMSGTypeError(f'Could not recognize msg class type "{ct}".')
    else:
        logger.error(f'Could not recognize msg class type "{msg.classType}". This most likely means it hasn\'t been implemented yet, and you should ask the developers to add support for it.')
        return msg


def openMsgBulk(path, **kwargs) -> Union[List, Tuple]:
    """
    Takes the same arguments as openMsg, but opens a collection of msg files
    based on a wild card. Returns a list if successful, otherwise returns a
    tuple.

    :param ignoreFailures: If this is True, will return a list of all successful
        files, ignoring any failures. Otherwise, will close all that
        successfully opened, and return a tuple of the exception and the path of
        the file that failed.
    """
    files = []
    for x in glob.glob(path):
        try:
            files.append(openMsg(x, **kwargs))
        except Exception as e:
            if not kwargs.get('ignoreFailures', False):
                for msg in files:
                    msg.close()
                return (e, x)

    return files


def parseType(_type : int, stream, encoding, extras):
    """
    Converts the data in :param stream: to a much more accurate type, specified
    by :param _type:.
    :param _type: the data's type.
    :param stream: is the data to be converted.
    :param encoding: is the encoding to be used for regular strings.
    :param extras: is used in the case of types like PtypMultipleString.
    For that example, extras should be a list of the bytes from rest of the
    streams.

    :raises NotImplementedError: for types with no current support. Most of
        these types have no documentation of existing in an MSG file.
    """
    # WARNING Not done. Do not try to implement anywhere where it is not already implemented.
    value = stream
    lengthExtras = len(extras)
    if _type == 0x0000:  # PtypUnspecified
        pass
    elif _type == 0x0001:  # PtypNull
        if value != b'\x00\x00\x00\x00\x00\x00\x00\x00':
            # DEBUG
            logger.warning('Property type is PtypNull, but is not equal to 0.')
        return None
    elif _type == 0x0002:  # PtypInteger16
        return constants.STI16.unpack(value)[0]
    elif _type == 0x0003:  # PtypInteger32
        return constants.STI32.unpack(value)[0]
    elif _type == 0x0004:  # PtypFloating32
        return constants.STF32.unpack(value)[0]
    elif _type == 0x0005:  # PtypFloating64
        return constants.STF64.unpack(value)[0]
    elif _type == 0x0006:  # PtypCurrency
        return (constants.STI64.unpack(value)[0]) / 10000.0
    elif _type == 0x0007:  # PtypFloatingTime
        value = constants.STF64.unpack(value)[0]
        return constants.PYTPFLOATINGTIME_START + datetime.timedelta(days = value)
    elif _type == 0x000A:  # PtypErrorCode
        value = constants.STUI32.unpack(value)[0]
        # TODO parsing for this.
        # I can't actually find any msg properties that use this, so it should
        # be okay to release this function without support for it.
        raise NotImplementedError('Parsing for type 0x000A (PtypErrorCode) has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType 0x000A PtypErrorCode".')
    elif _type == 0x000B:  # PtypBoolean
        return constants.ST3.unpack(value)[0] == 1
    elif _type == 0x000D:  # PtypObject/PtypEmbeddedTable
        # TODO parsing for this
        # Wait, that's the extension for an attachment folder, so parsing this
        # might not be as easy as we would hope. The function may be released
        # without support for this.
        raise NotImplementedError('Current version of extract-msg does not support the parsing of PtypObject/PtypEmbeddedTable in this function.')
    elif _type == 0x0014:  # PtypInteger64
        return constants.STI64.unpack(value)[0]
    elif _type == 0x001E:  # PtypString8
        return value.decode(encoding)
    elif _type == 0x001F:  # PtypString
        return value.decode('utf-16-le')
    elif _type == 0x0040:  # PtypTime
        rawtime = constants.ST3.unpack(value)[0]
        if rawtime < 116444736000000000:
            # We can't properly parse this with our current setup, so
            # we will rely on olefile to handle this one.
            value = olefile.olefile.filetime2datetime(rawtime)
        else:
            if rawtime != 915151392000000000:
                value = fromTimeStamp(filetimeToUtc(rawtime))
            else:
                # Temporarily just set to max time to signify a null date.
                value = datetime.datetime.max
        return value
    elif _type == 0x0048:  # PtypGuid
        return bytesToGuid(value)
    elif _type == 0x00FB:  # PtypServerId
        count = constants.STUI16.unpack(value[:2])
        # If the first byte is a 1 then it uses the ServerID structure.
        if value[3] == 1:
            from .data import ServerID
            return ServerID(value)
        else:
            return (count, value[2:count + 2])
    elif _type == 0x00FD:  # PtypRestriction
        # TODO parsing for this
        raise NotImplementedError('Parsing for type 0x00FD (PtypRestriction) has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType 0x00FD PtypRestriction".')
    elif _type == 0x00FE:  # PtypRuleAction
        # TODO parsing for this
        raise NotImplementedError('Parsing for type 0x00FE (PtypRuleAction) has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType 0x00FE PtypRuleAction".')
    elif _type == 0x0102:  # PtypBinary
        return value
    elif _type & 0x1000 == 0x1000:  # PtypMultiple
        # TODO parsing for `multiple` types
        if _type in (0x101F, 0x101E): # PtypMultipleString/PtypMultipleString8
            ret = [x.decode(encoding) for x in extras]
            lengths = struct.unpack(f'<{len(ret)}i', stream)
            lengthLengths = len(lengths)
            if lengthLengths > lengthExtras:
                logger.warning(f'Error while parsing multiple type. Expected {lengthLengths} stream{"s" if lengthLengths != 1 else ""}, got {lengthExtras}. Ignoring.')
            for x, y in enumerate(extras):
                if lengths[x] != len(y):
                    logger.warning(f'Error while parsing multiple type. Expected length {lengths[x]}, got {len(y)}. Ignoring.')
            return ret
        elif _type == 0x1102: # PtypMultipleBinary
            ret = copy.deepcopy(extras)
            lengths = tuple(constants.STUI32.unpack(stream[pos*8:(pos+1)*8])[0] for pos in range(len(stream) // 8))
            lengthLengths = len(lengths)
            if lengthLengths > lengthExtras:
                logger.warning(f'Error while parsing multiple type. Expected {lengthLengths} stream{"s" if lengthLengths != 1 else ""}, got {lengthExtras}. Ignoring.')
            for x, y in enumerate(extras):
                if lengths[x] != len(y):
                    logger.warning(f'Error while parsing multiple type. Expected length {lengths[x]}, got {len(y)}. Ignoring.')
            return ret
        elif _type in (0x1002, 0x1003, 0x1004, 0x1005, 0x1007, 0x1014, 0x1040, 0x1048):
            if stream != len(extras):
                logger.warning(f'Error while parsing multiple type. Expected {stream} entr{"y" if stream == 1 else "ies"}, got {len(extras)}. Ignoring.')
            if _type == 0x1002: # PtypMultipleInteger16
                return tuple(constants.STMI16.unpack(x)[0] for x in extras)
            if _type == 0x1003: # PtypMultipleInteger32
                return tuple(constants.STMI32.unpack(x)[0] for x in extras)
            if _type == 0x1004: # PtypMultipleFloating32
                return tuple(constants.STMF32.unpack(x)[0] for x in extras)
            if _type == 0x1005: # PtypMultipleFloating64
                return tuple(constants.STMF64.unpack(x)[0] for x in extras)
            if _type == 0x1007: # PtypMultipleFloatingTime
                values = tuple(constants.STMF64.unpack(x)[0] for x in extras)
                return tuple(constants.PYTPFLOATINGTIME_START + datetime.timedelta(days = amount) for amount in values)
            if _type == 0x1014: # PtypMultipleInteger64
                return tuple(constants.STMI64.unpack(x)[0] for x in extras)
            if _type == 0x1040: # PtypMultipleTime
                return tuple(filetimeToUtc(constants.ST3.unpack(x)[0]) for x in extras)
            if _type == 0x1048: # PtypMultipleGuid
                return tuple(bytesToGuid(x) for x in extras)
        else:
            raise NotImplementedError(f'Parsing for type {_type} has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType {_type}".')
    return value


def prepareFilename(filename) -> str:
    """
    Adjusts :param filename: so that it can succesfully be used as an actual
    file name.
    """
    # I would use re here, but it tested to be slightly slower than this.
    return ''.join(i for i in filename if i not in r'\/:*?"<>|' + '\x00')


def properHex(inp, length : int = 0) -> str:
    """
    Takes in various input types and converts them into a hex string whose
    length will always be even.
    """
    a = ''
    if isinstance(inp, str):
        a = ''.join([hex(ord(inp[x]))[2:].rjust(2, '0') for x in range(len(inp))])
    elif isinstance(inp, bytes):
        a = inp.hex()
    elif isinstance(inp, int):
        a = hex(inp)[2:]
    if len(a) % 2 != 0:
        a = '0' + a
    return a.rjust(length, '0').upper()


def roundUp(inp : int, mult : int) -> int:
    """
    Rounds :param inp: up to the nearest multiple of :param mult:.
    """
    return inp + (mult - inp) % mult


def rtfSanitizeHtml(inp : str) -> str:
    """
    Sanitizes input to an RTF stream that has encapsulated HTML.
    """
    if not inp:
        return ''
    output = ''
    for char in inp:
        # Check if it is in the right range to be printed directly.
        if 32 <= ord(char) < 128:
            # Quick check for handling the HTML escapes. Will eventually
            # upgrade this code to actually handle all the HTML escapes
            # but this will do for now.
            if char == '<':
                output += r'{\*\htmltag84 &lt;}\htmlrtf <\htmlrtf0 '
            elif char == '>':
                output += r'{\*\htmltag84 &gt;}\htmlrtf >\htmlrtf0'
            else:
                if char in ('\\', '{', '}'):
                    output += '\\'
                output += char
        elif ord(char) < 32 or 128 <= ord(char) <= 255:
            # Otherwise, see if it is just a small escape.
            output += "\\'" + properHex(char, 2)
        else:
            # Handle Unicode characters.
            output += '\\u' + str(ord(char)) + '?'

    return output


def rtfSanitizePlain(inp : str) -> str:
    """
    Sanitizes input to a plain RTF stream.
    """
    if not inp:
        return ''
    output = ''
    for char in inp:
        # Check if it is in the right range to be printed directly.
        if 32 <= ord(char) < 128:
            if char in ('\\', '{', '}'):
                output += '\\'
            output += char
        elif ord(char) < 32 or 128 <= ord(char) <= 255:
            # Otherwise, see if it is just a small escape.
            output += "\\'" + properHex(char, 2)
        else:
            # Handle Unicode characters.
            output += '\\u' + str(ord(char)) + '?'

    return output


def setupLogging(defaultPath = None, defaultLevel = logging.WARN, logfile = None, enableFileLogging : bool = False,
                  env_key = 'EXTRACT_MSG_LOG_CFG') -> bool:
    """
    Setup logging configuration

    Args:
    :param defaultPath: Default path to use for the logging configuration file.
    :param defaultLevel: Default logging level.
    :param env_key: Environment variable name to search for, for setting logfile
        path.
    :param enableFileLogging: Whether to use a file to log or not.

    Returns:
        bool: True if the configuration file was found and applied, False otherwise
    """
    shippedConfig = pathlib.Path(__file__).parent / 'logging-config'
    if os.name == 'nt':
        null = 'NUL'
        shippedConfig /= 'logging-nt.json'
    elif os.name == 'posix':
        null = '/dev/null'
        shippedConfig /= 'logging-posix.json'
    # Find logging.json if not provided
    defaultPath = pathlib.Path(defaultPath) if defaultPath else shippedConfig

    paths = [
        defaultPath,
        pathlib.Path('logging.json'),
        pathlib.Path('../logging.json'),
        pathlib.Path('../../logging.json'),
        shippedConfig,
    ]

    path = None

    for configPath in paths:
        if configPath.exists():
            path = configPath
            break

    value = os.getenv(env_key, None)
    if value and os.path.exists(value) and os.path.isfile(value):
        path = pathlib.Path(value)

    if not path:
        print('Unable to find logging.json configuration file')
        print('Make sure a valid logging configuration file is referenced in the defaultPath'
              ' argument, is inside the extract_msg install location, or is available at one '
              'of the following file-paths:')
        print(str(paths[1:]))
        logging.basicConfig(level = defaultLevel)
        logging.warning('The extract_msg logging configuration was not found - using a basic configuration.'
                        f'Please check the extract_msg installation directory for "logging-{os.name}.json".')
        return False

    with open(path, 'rt') as f:
        config = json.load(f)

    for x in config['handlers']:
        if 'filename' in config['handlers'][x]:
            if enableFileLogging:
                config['handlers'][x]['filename'] = tmp = os.path.expanduser(
                    os.path.expandvars(logfile if logfile else config['handlers'][x]['filename']))
                tmp = pathlib.Path(tmp).parent
                if not tmp.exists:
                    os.makedirs(tmp)
            else:
                config['handlers'][x]['filename'] = null

    try:
        logging.config.dictConfig(config)
    except ValueError as e:
        print('Failed to configure the logger. Did your installation get messed up?')
        print(e)

    logging.getLogger().setLevel(defaultLevel)
    return True


def unwrapMsg(msg : "MSGFile") -> Dict:
    """
    Takes a recursive message-attachment structure and unwraps it into a linear
    dictionary for easy iteration. Dictionary contains 4 keys: "attachments" for
    main message attachments, not including embedded MSG files, "embedded" for
    attachments representing embedded MSG files, "msg" for all MSG files
    (including the original in the first index), and "raw_attachments" for raw
    attachments from signed messages.
    """
    from .message_signed_base import MessageSignedBase

    # Here is where we store main attachments.
    attachments = []
    # Here is where we are going to store embedded msg files.
    msgFiles = [msg]
    # Here is where we store embedded attachments.
    embedded = []
    # Here is where we store raw attachments from signed messages.
    raw = []

    # Normally we would need a recursive function to unwrap a recursive
    # structure like the message-attachment structure. Essentially, a function
    # that calls itself. Here, I have designed code capable of circumventing
    # this to do it in a single function, which is a lot more efficient and
    # safer. That is why we store the `toProcess` and use a while loop
    # surrounding a for loop. The for loop would be the main body of the
    # function, while the append to toProcess would be the recursive call.
    toProcess = collections.deque((msg,))

    while len(toProcess) > 0:
        # Remove the last item from the list of things to process, and store it
        # in `currentItem`. We will be processing it in the for loop.
        currentItem = toProcess.popleft()
        # iterate through the attachments and
        for att in currentItem.attachments:
            # If it is a regular attachment, add it to the list. Otherwise, add
            # it to be processed
            if att.type in (AttachmentType.DATA, AttachmentType.SIGNED):
                attachments.append(att)
            elif att.type == AttachmentType.MSG:
                # Here we do two things. The first is we store it to the output
                # so we can return it. The second is we add it to the processing
                # list. The reason this is two steps is because we need to be
                # able to remove items from the processing list, but can't
                # do that from the output.
                embedded.append(att)
                msgFiles.append(att.data)
                toProcess.append(att.data)
        if isinstance(currentItem, MessageSignedBase):
            raw += currentItem._rawAttachments

    return {
        'attachments': attachments,
        'embedded': embedded,
        'msg': msgFiles,
        'raw_attachments': raw,
    }


def unwrapMultipart(mp : Union[bytes, str, email.message.Message]) -> Dict:
    """
    Unwraps a recursive multipart structure into a dictionary of linear lists.
    Similar to unwrapMsg, but for multipart. Dictionary contains 3 keys:
    "attachments" which contains a list of dicts containing processed attachment
    data as well as the Message instance associated with it, "plain_body" which
    contains the plain text body, and "html_body" which contains the HTML body.

    For clarification, each instance of processed attachment data is a dict
    with keys identical to the args used for the SignedAttachment constructor.
    This makes it easy to expand for use in constructing a SignedAttachment. The
    only argument missing is "msg" to ensure this function will not require one.

    :param mp: The bytes that make up a multipart, the string that makes up a
        multipart, or a Message instance from the email module created from the
        multipart to unwrap. If providing a Message instance, prefer it to be an
        instance of EmailMessage. If you are doing so, make sure it's policy is
        default.
    """
    # In the event we are generating it, these are the kwargs to use.
    genKwargs = {
        '_class': email.message.EmailMessage,
        'policy': email.policy.default,
    }
    # Convert our input into something usable.
    if isinstance(mp, email.message.EmailMessage):
        if mp.policy == email.policy.default:
            mpMessage = mp
        else:
            mpMessage = email.message_from_bytes(mp.as_bytes(), **genKwargs)
    elif isinstance(mp, email.message.Message):
        mpMessage = email.message_from_bytes(mp.as_bytes(), **genKwargs)
    elif isinstance(mp, bytes):
        mpMessage = email.message_from_bytes(mp, **genKwargs)
    elif isinstance(mp, str):
        mpMessage = email.message_from_string(mp, **genKwargs)
    else:
        raise TypeError(f'Unsupported type "{type(mp)}" provided to unwrapMultipart.')

    # Okay, now that we have it in a useable form, let's do the most basic
    # unwrapping possible. Once the most basic unwrapping is done, we can
    # actually process the data. For this, we only care if the section is
    # multipart or not. If it is, it get's unwrapped too.
    #
    # In case you are curious, this is effectively doing a breadth first
    # traversal of the tree.
    dataNodes = []

    toProcess = collections.deque((mpMessage,))
    # I do know about the walk method, but it might *also* walk embedded
    # messages which we very much don't want.
    while len(toProcess) > 0:
        currentItem = toProcess.popleft()
        # 'multipart' indicates that it shouldn't contain any data itself, just
        # other nodes to go through.
        if currentItem.get_content_maintype() == 'multipart':
            payload = currentItem.get_payload()
            # For multipart, the payload should be a list, but handle it not
            # being one.
            if isinstance(payload, list):
                toProcess.extend(payload)
            else:
                logging.warn('Found multipart node that did not return a list. Appending as a data node.')
                dataNodes.append(currentItem)
        else:
            # The opposite is *not* true. If it's not multipart, always add as a
            # data node.
            dataNodes.append(currentItem)

    # At this point, all of our nodes should have processed and we should now
    # have data nodes. Now let's process them. For anything that was parsed as
    # a message, we actually want to get it's raw bytes back so it can be saved.
    # If they user wants to process that message in some way, they can do it
    # themself.
    attachments = []
    plainBody = None
    htmlBody = None

    for node in dataNodes:
        # Let's setup our attachment we are going to use.
        attachment = {
            'data': None,
            'name': node.get_filename(),
            'mimetype': node.get_content_type(),
            'node': node,
        }

        # Finally, we need to get the data. As we need to ensure it is bytes,
        # we may have to do some special processing.
        data = node.get_content()
        if isinstance(data, bytes):
            # If the data is bytes, we are perfectly good.
            pass
        elif isinstance(data, email.message.Message):
            # If it is a message, get it's bytes directly.
            data = data.as_bytes()
        elif isinstance(data, str):
            # If it is a string, let's reverse encode it where possible.
            # First thing we want to check is if we can find the encoding type.
            # If we can, use that to reverse the process. Otherwise use utf-8.
            data = data.encode(node.get_content_charset('utf-8'))
        else:
            # We throw an exception to describe the problem if we can't reverse
            # the problem.
            raise TypeError(f'Attempted to get bytes for attachment, but could not convert {type(data)} to bytes.')

        attachment['data'] = data

        # Now for the fun part, figuring out if we actually have an attachment.
        if attachment['name']:
            attachments.append(attachment)
        elif attachment['mimetype'] == 'text/plain':
            if plainBody:
                logger.warn('Found multiple candidates for plain text body.')
            plainBody = data
        elif attachment['mimetype'] == 'text/html':
            if htmlBody:
                logger.warn('Found multiple candidates for HTML body.')
            htmlBody = data

    return {
        'attachments': attachments,
        'plain_body': plainBody,
        'html_body': htmlBody,
    }

def validateHtml(html : bytes) -> bool:
    """
    Checks whether the HTML is considered valid. To be valid, the HTML must, at
    minimum, contain an <html> tag, a <body> tag, and closing tags for each.
    """
    bs = bs4.BeautifulSoup(html, 'html.parser')
    if not bs.find('html') or not bs.find('body'):
        return False
    return True


def verifyPropertyId(id : str) -> None:
    """
    Determines whether a property ID is valid for vertain functions. Property
    IDs MUST be a 4 digit hexadecimal string. Property is valid if no exception
    is raised.

    :raises InvaildPropertyIdError: if the it is not a 4 digit hexadecimal
        number.
    """
    if not isinstance(id, str):
        raise InvaildPropertyIdError('ID was not a 4 digit hexadecimal string')
    elif len(id) != 4:
        raise InvaildPropertyIdError('ID was not a 4 digit hexadecimal string')
    else:
        try:
            int(id, 16)
        except ValueError:
            raise InvaildPropertyIdError('ID was not a 4 digit hexadecimal string')


def verifyType(_type):
    """
    Verifies that the type is valid. Raises an exception if it is not.

    :raises UnknownTypeError: if the type is not recognized.
    """
    if _type is not None:
        if (_type not in constants.VARIABLE_LENGTH_PROPS_STRING) and (_type not in constants.FIXED_LENGTH_PROPS_STRING):
            raise UnknownTypeError(f'Unknown type {_type}.')


def windowsUnicode(string):
    return str(string, 'utf-16-le') if string is not None else None
