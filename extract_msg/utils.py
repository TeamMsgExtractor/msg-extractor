"""
Utility functions of extract_msg.
"""

import argparse
import datetime
import json
import logging
import logging.config
import sys

import tzlocal

from extract_msg import constants
from extract_msg.compat import os_ as os

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())
logging.addLevelName(5, 'DEVELOPER')

if sys.version_info[0] >= 3:  # Python 3
    stri = (str,)

    get_input = input


    def encode(inp):
        return inp


    def properHex(inp):
        """
        Taken (with permission) from https://github.com/TheElementalOfCreation/creatorUtils
        """
        a = ''
        if isinstance(inp, stri):
            a = ''.join([hex(ord(inp[x]))[2:].rjust(2, '0') for x in range(len(inp))])
        elif isinstance(inp, bytes):
            a = inp.hex()
        elif isinstance(inp, int):
            a = hex(inp)[2:]
        if len(a) % 2 != 0:
            a = '0' + a
        return a


    def windowsUnicode(string):
        return str(string, 'utf_16_le') if string is not None else None


    def xstr(s):
        return '' if s is None else str(s)

else:  # Python 2
    stri = (str, unicode)

    get_input = raw_input


    def encode(inp):
        return inp.encode('utf-8') if inp is not None else None


    def properHex(inp):
        """
        Taken (with permission) from https://github.com/TheElementalOfCreation/creatorUtils
        """
        a = ''
        if isinstance(inp, stri):
            a = ''.join([hex(ord(inp[x]))[2:].rjust(2, '0') for x in range(len(inp))])
        elif isinstance(inp, int):
            a = hex(inp)[2:]
        elif isinstance(inp, long):
            a = hex(inp)[2:-1]
        if len(a) % 2 != 0:
            a = '0' + a
        return a


    def windowsUnicode(string):
        return unicode(string, 'utf_16_le') if string is not None else None


    def xstr(s):
        if isinstance(s, unicode):
            return s.encode('utf-8')
        else:
            return '' if s is None else str(s)


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


def divide(string, length):
    """
    Taken (with permission) from https://github.com/TheElementalOfCreation/creatorUtils

    Divides a string into multiple substrings of equal length
    :param string: string to be divided.
    :param length: length of each division.
    :returns: list containing the divided strings.

    Example:
    >>>> a = divide('Hello World!', 2)
    >>>> print(a)
    ['He', 'll', 'o ', 'Wo', 'rl', 'd!']
    """
    return [string[length * x:length * (x + 1)] for x in range(int(len(string) / length))]


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
                        help='Enables file logging. Implies --verbose')
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
    # --out-name NAME
    # parser.add_argument('--out-name', dest = 'out_name',
    #                     help = 'Name to be used with saving the file output. Should come immediately after the file name')
    # [msg files]
    parser.add_argument('msgs', metavar='msg', nargs='+',
                        help='An msg file to be parsed')

    options = parser.parse_args(args)
    if options.dev or options.file_logging:
        options.verbose = True
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


def has_len(obj):
    """
    Checks if :param obj: has a __len__ attribute.
    """
    try:
        obj.__len__
        return True
    except AttributeError:
        return False


def msgEpoch(inp):
    """
    Taken (with permission) from https://github.com/TheElementalOfCreation/creatorUtils
    """
    return (inp - 116444736000000000) / 10000000.0


def parse_type(_type, stream):
    """
    Converts the data in :param stream: to a
    much more accurate type, specified by
    :param _type:, if possible.
    :param stream # TODO what is stream?

    Some types require that :param prop_value: be specified. This can be retrieved from the Properties instance.

    WARNING: Not done. Do not try to implement anywhere where it is not already implemented
    """
    # WARNING Not done. Do not try to implement anywhere where it is not already implemented
    value = stream
    if _type == 0x0000:  # PtypUnspecified
        pass
    elif _type == 0x0001:  # PtypNull
        if value != b'\x00\x00\x00\x00\x00\x00\x00\x00':
            # DEBUG
            logger.warning('Property type is PtypNull, but is not equal to 0.')
        value = None
    elif _type == 0x0002:  # PtypInteger16
        value = constants.STI16.unpack(value)[0]
    elif _type == 0x0003:  # PtypInteger32
        value = constants.STI32.unpack(value)[0]
    elif _type == 0x0004:  # PtypFloating32
        value = constants.STF32.unpack(value)[0]
    elif _type == 0x0005:  # PtypFloating64
        value = constants.STF64.unpack(value)[0]
    elif _type == 0x0006:  # PtypCurrency
        value = (constants.STI64.unpack(value)[0]) / 10000.0
    elif _type == 0x0007:  # PtypFloatingTime
        value = constants.STF64.unpack(value)[0]
        # TODO parsing for this
        pass
    elif _type == 0x000A:  # PtypErrorCode
        value = constants.STI32.unpack(value)[0]
        # TODO parsing for this
        pass
    elif _type == 0x000B:  # PtypBoolean
        value = bool(constants.ST3.unpack(value)[0])
    elif _type == 0x000D:  # PtypObject/PtypEmbeddedTable
        # TODO parsing for this
        pass
    elif _type == 0x0014:  # PtypInteger64
        value = constants.STI64.unpack(value)[0]
    elif _type == 0x001E:  # PtypString8
        # TODO parsing for this
        pass
    elif _type == 0x001F:  # PtypString
        value = value.decode('utf_16_le')
    elif _type == 0x0040:  # PtypTime
        value = constants.ST3.unpack(value)[0]
    elif _type == 0x0048:  # PtypGuid
        # TODO parsing for this
        pass
    elif _type == 0x00FB:  # PtypServerId
        # TODO parsing for this
        pass
    elif _type == 0x00FD:  # PtypRestriction
        # TODO parsing for this
        pass
    elif _type == 0x00FE:  # PtypRuleAction
        # TODO parsing for this
        pass
    elif _type == 0x0102:  # PtypBinary
        # TODO parsing for this
        # Smh, how on earth am I going to code this???
        pass
    elif _type & 0x1000 == 0x1000:  # PtypMultiple
        # TODO parsing for `multiple` types
        pass
    return value


def getContFileDir(_file_):
    """
    Takes in the path to a file and tries to return the containing folder.
    """
    return '/'.join(_file_.replace('\\', '/').split('/')[:-1])


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


def get_full_class_name(inp):
    return inp.__class__.__module__ + '.' + inp.__class__.__name__
