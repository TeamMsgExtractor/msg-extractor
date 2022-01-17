"""
Utility functions of extract_msg.
"""

import argparse
import codecs
import copy
import datetime
import json
import logging
import logging.config
import os
import pathlib
import struct
# Not actually sure if this needs to be here for the logging, so just in case.
import sys

import tzlocal

from html import escape as htmlEscape

from . import constants
from .exceptions import ConversionError, IncompatibleOptionsError, InvaildPropertyIdError, UnknownCodepageError, UnknownTypeError, UnrecognizedMSGTypeError, UnsupportedMSGTypeError


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())
logging.addLevelName(5, 'DEVELOPER')


def addNumToDir(dirName):
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

def bitwiseAdjust(inp, mask):
    """
    Uses a given mask to adjust the location of bits after an operation like
    bitwise AND. This is useful for things like flags where you are trying to
    get a small portion of a larger number. Say for example, you had the number
    0xED (0b11101101) and you needed the adjusted result of the AND operation
    with 0x70 (0b01110000). The result of the and operation (0b01100000) and the
    mask used to get it (0x70) are give and the output gets adjusted to be 0x6
    (0b110).

    :param mask: MUST be greater than 0.
    """
    if mask < 1:
        raise ValueError('Mask MUST be greater than 0')
    return inp >> bin(mask)[::-1].index('1')

def bitwiseAdjustedAnd(inp, mask):
    """
    Preforms the bitwise AND operation between :param inp: and :param mask: and
    adjusts the results based on the rules of the bitwiseAdjust function.
    """
    if mask < 1:
        raise ValueError('Mask MUST be greater than 0')
    return (inp & mask) >> bin(mask)[::-1].index('1')

def bytesToGuid(bytes_input):
    hexinput = [properHex(byte) for byte in bytes_input]
    hexs = [hexinput[3] + hexinput[2] + hexinput[1] + hexinput[0], hexinput[5] + hexinput[4], hexinput[7] + hexinput[6], hexinput[8] + hexinput[9], ''.join(hexinput[10:16])]
    return '{{{}-{}-{}-{}-{}}}'.format(*hexs).upper()

def ceilDiv(n, d):
    """
    Returns the int from the ceil division of n / d.
    ONLY use ints as inputs to this function.

    For ints, this is faster and more accurate for numbers
    outside the precision range of float.
    """
    return -(n // -d)

def divide(string, length):
    """
    Taken (with permission) from https://github.com/TheElementalOfDestruction/creatorUtils

    Divides a string into multiple substrings of equal length.
    If there is not enough for the last substring to be equal,
    it will simply use the rest of the string.
    Can also be used for things like lists and tuples.

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

def fromTimeStamp(stamp):
    return datetime.datetime.fromtimestamp(stamp, tzlocal.get_localzone())

def getCommandArgs(args):
    """
    Parse command-line arguments
    """
    parser = argparse.ArgumentParser(description=constants.MAINDOC, prog='extract_msg')
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
    parser.add_argument('--json', dest='json', action='store_true',
                        help='Changes to write output files as json.')
    # --file-logging
    parser.add_argument('--file-logging', dest='file_logging', action='store_true',
                        help='Enables file logging. Implies --verbose.')
    # --verbose
    parser.add_argument('--verbose', dest='verbose', action='store_true',
                        help='Turns on console logging.')
    # --log PATH
    parser.add_argument('--log', dest='log',
                        help='Set the path to write the file log to.')
    # --config PATH
    parser.add_argument('--config', dest='config_path',
                        help='Set the path to load the logging config from.')
    # --out PATH
    parser.add_argument('--out', dest='out_path',
                        help='Set the folder to use for the program output. (Default: Current directory)')
    # --use-filename
    parser.add_argument('--use-filename', dest='use_filename', action='store_true',
                        help='Sets whether the name of each output is based on the msg filename.')
    # --dump-stdout
    parser.add_argument('--dump-stdout', dest='dump_stdout', action='store_true',
                        help='Tells the program to dump the message body (plain text) to stdout. Overrides saving arguments.')
    # --html
    parser.add_argument('--html', dest='html', action='store_true',
                       help='Sets whether the output should be html. If this is not possible, will error.')
    # --raw
    parser.add_argument('--raw', dest='raw', action='store_true',
                       help='Sets whether the output should be html. If this is not possible, will error.')
    # --rtf
    parser.add_argument('--rtf', dest='rtf', action='store_true',
                       help='Sets whether the output should be rtf. If this is not possible, will error.')
    # --allow-fallback
    parser.add_argument('--allow-fallback', dest='allowFallbac', action='store_true',
                       help='Tells the program to fallback to a different save type if the selected one is not possible.')
    # --out-name NAME
    parser.add_argument('--out-name', dest = 'out_name',
                        help = 'Name to be used with saving the file output. Should come immediately after the file name.')
    # [msg files]
    parser.add_argument('msgs', metavar='msg', nargs='+',
                        help='An msg file to be parsed')

    options = parser.parse_args(args)
    # Check if more than one of the following arguments has been specified
    if options.html + options.rtf + options.json > 1:
       raise IncompatibleOptionsError('Only one of these options may be selected at a time: --html, --json, --raw, --rtf')

    if options.dev or options.file_logging:
        options.verbose = True

    # If dump_stdout is True, we need to unset all arguments used in files.
    # Technically we actually only *need* to unset `out_path`, but that may
    # change in the future, so let's be thorough.
    if options.dump_stdout:
        options.out_path = None
        options.json = False
        options.rtf = False
        options.html = False
        options.use_filename = False
        options.cid = False

    file_args = options.msgs
    file_tables = []  # This is where we will store the separated files and their arguments
    temp_table = []  # temp_table will store each table while it is still being built.
    need_arg = True  # This tells us if the last argument was something like
    # --out-name which requires a string name after it.
    # We start on true to make it so that we use don't have to have something checking if we are on the first table.
    for x in file_args:  # Iterate through each
        if need_arg:
            temp_table.append(x)
            need_arg = False
        elif x in constants.KNOWN_FILE_FLAGS:
            temp_table.append(x)
            if x in constants.NEEDS_ARG:
                need_arg = True
        else:
            file_tables.append(temp_table)
            temp_table = [x]

    file_tables.append(temp_table)
    options.msgs = file_tables
    return options

def getContFileDir(_file_):
    """
    Takes in the path to a file and tries to return the containing folder.
    """
    return '/'.join(_file_.replace('\\', '/').split('/')[:-1])

def getEncodingName(codepage):
    """
    Returns the name of the encoding with the specified codepage.
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

def hasLen(obj):
    """
    Checks if :param obj: has a __len__ attribute.
    """
    return hasattr(obj, '__len__')

def injectHtmlHeader(msgFile):
    """
    Returns the HTML body from the MSG file (will check that it has one) with
    the HTML header injected into it.
    """
    if not hasattr(msgFile, 'htmlBody') or not msgFile.htmlBody:
        raise AttributeError('Cannot inject the HTML header without an HTML body attribute.')

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
            'subject': inputToString(htmlEscape(msgFile.subject), 'utf-8'),
        }).encode('utf-8')

    # Use the previously defined function to inject the HTML header.
    return constants.RE_HTML_BODY_START.sub(replace, msgFile.htmlBody, 1)

def injectRtfHeader(msgFile):
    """
    Returns the RTF body from the MSG file (will check that it has one) with the
    RTF header injected into it.
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

    raise Exception('All injection attempts failed.')

def inputToBytes(stringInputVar, encoding):
    if isinstance(stringInputVar, bytes):
        return stringInputVar
    elif isinstance(stringInputVar, str):
        return stringInputVar.encode(encoding)
    elif stringInputVar is None:
        return b''
    else:
        raise ConversionError('Cannot convert to bytes.')

def inputToMsgpath(inp):
    """
    Converts the input into an msg path.
    """
    if isinstance(inp, (list, tuple)):
        inp = '/'.join(inp)
    ret = inputToString(inp, 'utf-8').replace('\\', '/').split('/')
    return ret if ret[0] != '' else []

def inputToString(bytesInputVar, encoding):
    if isinstance(bytesInputVar, str):
        return bytesInputVar
    elif isinstance(bytesInputVar, bytes):
        return bytesInputVar.decode(encoding)
    elif bytesInputVar is None:
        return ''
    else:
        raise ConversionError('Cannot convert to str type.')

def isEncapsulatedRtf(inp):
    """
    Currently the destection is made to be *extremly* basic, but this will work
    for now. In the future this will be fixed to that literal text in the body
    of a message won't cause false detection.
    """
    return b'\\fromhtml' in inp

def isEmptyString(inp):
    """
    Returns true if the input is None or is an Empty string.
    """
    return (inp == '' or inp is None)

def knownMsgClass(classType):
    """
    Checks if the specified class type is recognized by the module. Usually used
    for checking if a type is simply unsupported rather than unknown.
    """
    classType = classType.lower()
    if classType == 'ipm':
        return True

    for item in constants.KNOWN_CLASS_TYPES:
        if classType.startsWith(item):
            return True

    return False

def msgEpoch(inp):
    """
    Taken (with permission) from https://github.com/TheElementalOfDestruction/creatorUtils
    """
    return (inp - 116444736000000000) / 10000000.0

def msgpathToString(inp):
    """
    Converts an msgpath (one of the internal paths inside an msg file) into a string.
    """
    if inp is None:
        return None
    if isinstance(inp, (list, tuple)):
        inp = '/'.join(inp)
    inp.replace('\\', '/')
    return inp

def openMsg(path, prefix = '', attachmentClass = None, filename = None, delayAttachments = False, overrideEncoding = None, attachmentErrorBehavior = constants.ATTACHMENT_ERROR_THROW, recipientSeparator = ';', strict = True):
    """
    Function to automatically open an MSG file and detect what type it is.

    :param path: Path to the msg file in the system or is the raw msg file.
    :param prefix: Used for extracting embeded msg files inside the main one.
        Do not set manually unless you know what you are doing.
    :param attachmentClass: Optional, the class the Message object will use for
        attachments. You probably should not change this value unless you know
        what you are doing.
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

    If :param strict: is set to `True`, this function will raise an exception
    when it cannot identify what MSGFile derivitive to use. Otherwise, it will
    log the error and return a basic MSGFile instance.

    Raises UnsupportedMSGTypeError and UnrecognizedMSGTypeError.
    """
    from .appointment import Appointment
    from .attachment import Attachment
    from .contact import Contact
    from .message import Message
    from .msg import MSGFile

    attachmentClass = Attachment if attachmentClass is None else attachmentClass

    msg = MSGFile(path, prefix, attachmentClass, filename, overrideEncoding, attachmentErrorBehavior)
    # After rechecking the docs, all comparisons should be case-insensitive, not case-sensitive. My reading ability is great.
    classType = msg.classType.lower()
    if classType.startswith('ipm.contact') or classType.startswith('ipm.distlist'):
        msg.close()
        return Contact(path, prefix, attachmentClass, filename, overrideEncoding, attachmentErrorBehavior)
    elif classType.startswith('ipm.note') or classType.startswith('report'):
        msg.close()
        return Message(path, prefix, attachmentClass, filename, delayAttachments, overrideEncoding, attachmentErrorBehavior, recipientSeparator)
    elif classType.startswith('ipm.appointment') or classType.startswith('ipm.schedule'):
        msg.close()
        return Appointment(path, prefix, attachmentClass, filename, delayAttachments, overrideEncoding, attachmentErrorBehavior, recipientSeparator)
    elif classType == 'ipm': # Unspecified format. It should be equal to this and not just start with it.
        return msg
    elif strict:
        ct = msg.classType
        msg.close()
        if knownMsgClass(classType):
            raise UnsupportedMSGTypeError(f'MSG type "{ct}" currently is not supported by the module. If you would like support, please make a feature request.')
        raise UnrecognizedMSGTypeError(f'Could not recognize msg class type "{ct}".')
    else:
        logger.error(f'Could not recognize msg class type "{msg.classType}". This most likely means it hasn\'t been implemented yet, and you should ask the developers to add support for it.')
        return msg

def parseType(_type, stream, encoding, extras):
    """
    Converts the data in :param stream: to a much more accurate type, specified
    by :param _type:.
    :param _type: the data's type.
    :param stream: is the data to be converted.
    :param encoding: is the encoding to be used for regular strings.
    :param extras: is used in the case of types like PtypMultipleString.
    For that example, extras should be a list of the bytes from rest of the
    streams.

    WARNING: Not done. Do not try to implement anywhere where it is not already implemented
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
        # I can't actually find any msg properties that use this, so it should be okay to release this function without support for it.
        raise NotImplementedError('Parsing for type 0x000A has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType 0x000A"')
    elif _type == 0x000B:  # PtypBoolean
        return constants.ST3.unpack(value)[0] == 1
    elif _type == 0x000D:  # PtypObject/PtypEmbeddedTable
        # TODO parsing for this
        # Wait, that's the extension for an attachment folder, so parsing this might not be as easy as we would hope. The function may be released without support for this.
        raise NotImplementedError('Current version of extract-msg does not support the parsing of PtypObject/PtypEmbeddedTable in this function.')
    elif _type == 0x0014:  # PtypInteger64
        return constants.STI64.unpack(value)[0]
    elif _type == 0x001E:  # PtypString8
        return value.decode(encoding)
    elif _type == 0x001F:  # PtypString
        return value.decode('utf-16-le')
    elif _type == 0x0040:  # PtypTime
        rawtime = constants.ST3.unpack(value)[0]
        if rawtime != 915151392000000000:
            value = fromTimeStamp(msgEpoch(rawtime))
        else:
            # Temporarily just set to max time to signify a null date.
            value = datetime.datetime.max
        return value
    elif _type == 0x0048:  # PtypGuid
        return bytesToGuid(value)
    elif _type == 0x00FB:  # PtypServerId
        # TODO parsing for this
        raise NotImplementedError('Parsing for type 0x00FB has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType 0x00FB"')
    elif _type == 0x00FD:  # PtypRestriction
        # TODO parsing for this
        raise NotImplementedError('Parsing for type 0x00FD has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType 0x00FD"')
    elif _type == 0x00FE:  # PtypRuleAction
        # TODO parsing for this
        raise NotImplementedError('Parsing for type 0x00FE has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType 0x00FE"')
    elif _type == 0x0102:  # PtypBinary
        return value
    elif _type & 0x1000 == 0x1000:  # PtypMultiple
        # TODO parsing for `multiple` types
        if _type in (0x101F, 0x101E):
            ret = [x.decode(encoding) for x in extras]
            lengths = struct.unpack(f'<{len(ret)}i', stream)
            lengthLengths = len(lengths)
            if lengthLengths > lengthExtras:
                logger.warning(f'Error while parsing multiple type. Expected {lengthLengths} stream{"s" if lengthLengths != 1 else ""}, got {lengthExtras}. Ignoring.')
            for x, y in enumerate(extras):
                if lengths[x] != len(y):
                    logger.warning(f'Error while parsing multiple type. Expected length {lengths[x]}, got {len(y)}. Ignoring.')
            return ret
        elif _type == 0x1102:
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
            if _type == 0x1002:
                return tuple(constants.STMI16.unpack(x)[0] for x in extras)
            if _type == 0x1003:
                return tuple(constants.STMI32.unpack(x)[0] for x in extras)
            if _type == 0x1004:
                return tuple(constants.STMF32.unpack(x)[0] for x in extras)
            if _type == 0x1005:
                return tuple(constants.STMF64.unpack(x)[0] for x in extras)
            if _type == 0x1007:
                values = tuple(constants.STMF64.unpack(x)[0] for x in extras)
                raise NotImplementedError('Parsing for type 0x1007 has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType 0x1007"')
            if _type == 0x1014:
                return tuple(constants.STMI64.unpack(x)[0] for x in extras)
            if _type == 0x1040:
                return tuple(msgEpoch(constants.ST3.unpack(x)[0]) for x in extras)
            if _type == 0x1048:
                return tuple(bytesToGuid(x) for x in extras)
        else:
            raise NotImplementedError(f'Parsing for type {_type} has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType {_type}"')
    return value

def prepareFilename(filename):
    """
    Adjusts :param filename: so that it can succesfully be used as an actual
    file name.
    """
    # I would use re here, but it tested to be slightly slower than this.
    return ''.join(i for i in filename if i not in r'\/:*?"<>|' + '\x00')

def properHex(inp, length = 0):
    """
    Taken (with permission) from
    https://github.com/TheElementalOfDestruction/creatorUtils
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

def roundUp(inp, mult):
    """
    Rounds :param inp: up to the nearest multiple of :param mult:.
    """
    return inp + (mult - inp) % mult

def rtfSanitizeHtml(inp):
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

def rtfSanitizePlain(inp):
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

def setupLogging(defaultPath=None, defaultLevel=logging.WARN, logfile=None, enableFileLogging=False,
                  env_key='EXTRACT_MSG_LOG_CFG'):
    """
    Setup logging configuration

    Args:
        defaultPath (str): Default path to use for the logging configuration file
        defaultLevel (int): Default logging level
        env_key (str): Environment variable name to search for, for setting logfile path

    Returns:
        bool: True if the configuration file was found and applied, False otherwise
    """
    shippedConfig = getContFileDir(__file__) + '/logging-config/'
    if os.name == 'nt':
        null = 'NUL'
        shippedConfig += 'logging-nt.json'
    elif os.name == 'posix':
        null = '/dev/null'
        shippedConfig += 'logging-posix.json'
    # Find logging.json if not provided
    if not defaultPath:
        defaultPath = shippedConfig

    paths = [
        defaultPath,
        'logging.json',
        '../logging.json',
        '../../logging.json',
        shippedConfig,
    ]

    path = None

    for configPath in paths:
        if os.path.exists(configPath):
            path = configPath
            break

    value = os.getenv(env_key, None)
    if value and os.path.exists(value):
        path = value

    if path is None:
        print('Unable to find logging.json configuration file')
        print('Make sure a valid logging configuration file is referenced in the defaultPath'
              ' argument, is inside the extract_msg install location, or is available at one '
              'of the following file-paths:')
        print(str(paths[1:]))
        logging.basicConfig(level=defaultLevel)
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
                tmp = getContFileDir(tmp)
                if not os.path.exists(tmp):
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

def verifyPropertyId(id):
    """
    Determines whether a property ID is valid for the functions that this function
    is called from. Property IDs MUST be a 4 digit hexadecimal string.
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
    if _type is not None:
        if (_type not in constants.VARIABLE_LENGTH_PROPS_STRING) and (_type not in constants.FIXED_LENGTH_PROPS_STRING):
            raise UnknownTypeError(f'Unknown type {_type}')

def windowsUnicode(string):
    return str(string, 'utf-16-le') if string is not None else None
