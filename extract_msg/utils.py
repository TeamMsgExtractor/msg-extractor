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
import struct
import sys

import tzlocal

from extract_msg import constants
from extract_msg.compat import os_ as os
from extract_msg.exceptions import ConversionError, IncompatibleOptionsError, InvaildPropertyIdError, UnknownCodepageError, UnknownTypeError, UnrecognizedMSGTypeError

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())
logging.addLevelName(5, 'DEVELOPER')

if sys.version_info[0] >= 3:  # Python 3
    get_input = input

    def properHex(inp, length = 0):
        """
        Taken (with permission) from https://github.com/TheElementalOfDestruction/creatorUtils
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

    def windowsUnicode(string):
        return str(string, 'utf_16_le') if string is not None else None

else:  # Python 2
    get_input = raw_input

    def properHex(inp, length = 0):
        """
        Converts the input into a hexadecimal string without the beginning "0x". The string
        will also always have a length that is a multiple of 2 (unless :param length: has
        been specified). :param length: only specifies the MINIMUM length that the string
        will use.
        """
        a = ''
        if isinstance(inp, (str, unicode)):
            a = ''.join([hex(ord(inp[x]))[2:].rjust(2, '0') for x in range(len(inp))])
        elif isinstance(inp, int):
            a = hex(inp)[2:]
        elif isinstance(inp, long):
            a = hex(inp)[2:-1]
        if len(a) % 2 != 0:
            a = '0' + a
        return a.rjust(length, '0').upper()

    def windowsUnicode(string):
        return unicode(string, 'utf_16_le') if string is not None else None

def addNumToDir(dirName):
    """
    Attempt to create the directory with a '(n)' appended.
    """
    for i in range(2, 100):
        try:
            newDirName = dirName + ' (' + str(i) + ')'
            os.makedirs(newDirName)
            return newDirName
        except Exception as e:
            pass
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

def prepareFilename(filename):
    """
    Adjusts :param filename: so that it can succesfully be used as an actual
    file name.
    """
    # I would use re here, but it tested to be slightly slower than this.
    return ''.join(i for i in filename if i not in r'\/:*?"<>|' + '\x00')

def fromTimeStamp(stamp):
    return datetime.datetime.fromtimestamp(stamp, tzlocal.get_localzone())

def get_command_args(args):
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
    #parser.add_argument('--html', dest='html', action='store_true',
    #                    help='Sets whether the output should be html. If this is not possible, will error.')
    # --rtf
    #parser.add_argument('--rtf', dest='rtf', action='store_true',
    #                    help='Sets whether the output should be rtf. If this is not possible, will error.')
    # --allow-fallback
    #parser.add_argument('--allow-fallback', dest='allowFallbac', action='store_true',
    #                    help='Tells the program to fallback to a different save type if the selected one is not possible.')
    # --out-name NAME
    # parser.add_argument('--out-name', dest = 'out_name',
    #                     help = 'Name to be used with saving the file output. Should come immediately after the file name.')
    # [msg files]
    parser.add_argument('msgs', metavar='msg', nargs='+',
                        help='An msg file to be parsed')

    options = parser.parse_args(args)
    # Check if more than one of the following arguments has been specified
    #valid = 0
    #if options.html:
    #    valid += 1
    #if options.rtf:
    #    valid += 1
    #if options.json:
    #    valid += 1
    #if valid > 1:
    #    raise IncompatibleOptionsError('Only one of these options may be selected at a time: --html, --rtf, --json')

    if options.dev or options.file_logging:
        options.verbose = True

    # If dump_stdout is True, we need to unset all arguments used in files.
    # Technically we actually only *need* to unset `out_path`, but that may
    # change in the future, so let's be thorough.
    if options.dump_stdout:
        options.out_path = None
        options.json = False
        #options.rtf = False
        #options.html = False
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
        raise UnsupportedEncodingError('The codepage {} ({}) is not currently supported by your version of Python.'.format(codepage, constants.CODE_PAGES[codepage]))

def get_full_class_name(inp):
    return inp.__class__.__module__ + '.' + inp.__class__.__name__

def has_len(obj):
    """
    Checks if :param obj: has a __len__ attribute.
    """
    try:
        obj.__len__
        return True
    except AttributeError:
        return False

def inputToBytes(string_input_var, encoding):
    if isinstance(string_input_var, constants.BYTES):
        return string_input_var
    elif isinstance(string_input_var, constants.STRING):
        return string_input_var.encode(encoding)
    elif string_input_var is None:
        return b''
    else:
        raise ConversionError('Cannot convert to BYTES type')

def inputToMsgpath(inp):
    """
    Converts the input into an msg path.
    """
    if isinstance(inp, (list, tuple)):
        inp = '/'.join(inp)
    ret = inputToString(inp, 'utf-8').replace('\\', '/').split('/')
    return ret if ret[0] != '' else []

def inputToString(bytes_input_var, encoding):
    if isinstance(bytes_input_var, constants.STRING):
        return bytes_input_var
    elif isinstance(bytes_input_var, constants.BYTES):
        return bytes_input_var.decode(encoding)
    elif bytes_input_var is None:
        return ''
    else:
        raise ConversionError('Cannot convert to STRING type')

def isEmptyString(inp):
    """
    Returns true if the input is None or is an Empty string.
    """
    return (inp == '' or inp is None)

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

def openMsg(path, prefix = '', attachmentClass = None, filename = None, delayAttachments = False, overrideEncoding = None, attachmentErrorBehavior = constants.ATTACHMENT_ERROR_THROW, strict = True):
    """
    Function to automatically open an MSG file and detect what type it is.

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

    If :param strict: is set to `True`, this function will raise an exception
    when it cannot identify what MSGFile derivitive to use. Otherwise, it will
    log the error and return a basic MSGFile instance.
    """
    from extract_msg.appointment import Appointment
    from extract_msg.attachment import Attachment
    from extract_msg.contact import Contact
    from extract_msg.message import Message
    from extract_msg.msg import MSGFile

    attachmentClass = Attachment if attachmentClass is None else attachmentClass

    msg = MSGFile(path, prefix, attachmentClass, filename, overrideEncoding, attachmentErrorBehavior)
    classtype = msg.classType
    if classtype.startswith('IPM.Contact') or classtype.startswith('IPM.DistList'):
        msg.close()
        return Contact(path, prefix, attachmentClass, filename, overrideEncoding, attachmentErrorBehavior)
    elif classtype.startswith('IPM.Note') or classtype.startswith('REPORT'):
        msg.close()
        return Message(path, prefix, attachmentClass, filename, delayAttachments, overrideEncoding, attachmentErrorBehavior)
    elif classtype.startswith('IPM.Appointment') or classtype.startswith('IPM.Schedule'):
        msg.close()
        return Appointment(path, prefix, attachmentClass, filename, delayAttachments, overrideEncoding, attachmentErrorBehavior)
    elif strict:
        msg.close()
        raise UnrecognizedMSGTypeError('Could not recognize msg class type "{}". It is recommended you report this to the developers.'.format(msg.classType))
    else:
        logger.error('Could not recognize msg class type "{}". It is recommended you report this to the developers.'.format(msg.classType))
        return msg

def parseType(_type, stream, encoding, extras):
    """
    Converts the data in :param stream: to a
    much more accurate type, specified by
    :param _type: the data's type.
    :param stream: is the data to be converted.
    :param encoding: is the encoding to be used for regular strings.
    :param extras: is used in the case of types like PtypMultipleString.
    For that example, extras should be a list of the bytes from rest of the streams.

    WARNING: Not done. Do not try to implement anywhere where it is not already implemented
    """
    # WARNING Not done. Do not try to implement anywhere where it is not already implemented
    value = stream
    length_extras = len(extras)
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
        # TODO parsing for this
        # I can't actually find any msg properties that use this, so it should be okay to release this function without support for it.
        raise NotImplementedError('Parsing for type 0x000A has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType 0x000A"')
    elif _type == 0x000B:  # PtypBoolean
        return bool(constants.ST3.unpack(value)[0])
    elif _type == 0x000D:  # PtypObject/PtypEmbeddedTable
        # TODO parsing for this
        # Wait, that's the extension for an attachment folder, so parsing this might not be as easy as we would hope. The function may be released without support for this.
        raise NotImplementedError('Current version of extract-msg does not support the parsing of PtypObject/PtypEmbeddedTable in this function.')
    elif _type == 0x0014:  # PtypInteger64
        return constants.STI64.unpack(value)[0]
    elif _type == 0x001E:  # PtypString8
        return value.decode(encoding)
    elif _type == 0x001F:  # PtypString
        return value.decode('utf_16_le')
    elif _type == 0x0040:  # PtypTime
        return fromTimeStamp(msgEpoch(constants.ST3.unpack(value)[0])).__format__('%a, %d %b %Y %H:%M:%S %z')
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
            lengths = struct.unpack('<{}i'.format(len(ret)), stream)
            length_lengths = len(lengths)
            if length_lengths > length_extras:
                logger.warning('Error while parsing multiple type. Expected {} stream{}, got {}. Ignoring.'.format(length_lengths, 's' if length_lengths > 1 or length_lengths == 0 else '', length_extras))
            for x, y in enumerate(extras):
                if lengths[x] != len(y):
                    logger.warning('Error while parsing multiple type. Expected length {}, got {}. Ignoring.'.format(lengths[x], len(y)))
            return ret
        elif _type == 0x1102:
            ret = copy.deepcopy(extras)
            lengths = tuple(constants.STUI32.unpack(stream[pos*8:(pos+1)*8])[0] for pos in range(len(stream) // 8))
            length_lengths = len(lengths)
            if length_lengths > length_extras:
                logger.warning('Error while parsing multiple type. Expected {} stream{}, got {}. Ignoring.'.format(length_lengths, 's' if length_lengths > 1 or length_lengths == 0 else '', length_extras))
            for x, y in enumerate(extras):
                if lengths[x] != len(y):
                    logger.warning('Error while parsing multiple type. Expected length {}, got {}. Ignoring.'.format(lengths[x], len(y)))
            return ret
        elif _type in (0x1002, 0x1003, 0x1004, 0x1005, 0x1007, 0x1014, 0x1040, 0x1048):
            if stream != len(extras):
                logger.warning('Error while parsing multiple type. Expected {} entr{}, got {}. Ignoring.'.format(stream, ('y' if stream == 1 else 'ies'), len(extras)))
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
            raise NotImplementedError('Parsing for type {} has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType {}"'.format(_type, _type))
    return value

def roundUp(inp, mult):
    """
    Rounds :param inp: up to the nearest multiple of :param mult:.
    """
    return inp + (mult - inp) % mult

def setup_logging(default_path=None, default_level=logging.WARN, logfile=None, enable_file_logging=False,
                  env_key='EXTRACT_MSG_LOG_CFG'):
    """
    Setup logging configuration

    Args:
        default_path (str): Default path to use for the logging configuration file
        default_level (int): Default logging level
        env_key (str): Environment variable name to search for, for setting logfile path

    Returns:
        bool: True if the configuration file was found and applied, False otherwise
    """
    shipped_config = getContFileDir(__file__) + '/logging-config/'
    if os.name == 'nt':
        null = 'NUL'
        shipped_config += 'logging-nt.json'
    elif os.name == 'posix':
        null = '/dev/null'
        shipped_config += 'logging-posix.json'
    # Find logging.json if not provided
    if not default_path:
        default_path = shipped_config

    paths = [
        default_path,
        'logging.json',
        '../logging.json',
        '../../logging.json',
        shipped_config,
    ]

    path = None

    for config_path in paths:
        if os.path.exists(config_path):
            path = config_path
            break

    value = os.getenv(env_key, None)
    if value and os.path.exists(value):
        path = value

    if path is None:
        print('Unable to find logging.json configuration file')
        print('Make sure a valid logging configuration file is referenced in the default_path'
              ' argument, is inside the extract_msg install location, or is available at one '
              'of the following file-paths:')
        print(str(paths[1:]))
        logging.basicConfig(level=default_level)
        logging.warning('The extract_msg logging configuration was not found - using a basic configuration.'
                        'Please check the extract_msg installation directory for "logging-{}.json".'.format(os.name))
        return False

    with open(path, 'rt') as f:
        config = json.load(f)

    for x in config['handlers']:
        if 'filename' in config['handlers'][x]:
            if enable_file_logging:
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

    logging.getLogger().setLevel(default_level)
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
            raise UnknownTypeError('Unknown type {}'.format(_type))
