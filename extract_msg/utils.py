"""
Utility functions of extract_msg.
"""

import argparse
import datetime
import logging
import json
import os
import sys
import tzlocal

from extract_msg import constants

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

if sys.version_info[0] >= 3:  # Python 3
    stri = (str,)


    def encode(inp):
        return inp


    def properHex(inp):
        """
        Taken (with permission) from https://github.com/TheElementalOfCreation/creatorUtils
        """
        a = ''
        if isinstance(inp, stri):
            a = ''.join([hex(ord(inp[x]))[2:].rjust(2, '0') for x in range(len(inp))])
        if isinstance(inp, bytes):
            a = inp.hex()
        elif isinstance(inp, int):
            a = hex(inp)[2:]
        if len(a) % 2 != 0:
            a = '0' + a
        return a


    def windowsUnicode(string):
        if string is None:
            return None
        return str(string, 'utf_16_le')


    def xstr(s):
        return '' if s is None else str(s)

else:  # Python 2
    stri = (str, unicode)


    def encode(inp):
        return inp.encode('utf-8')


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
        if string is None:
            return None
        return unicode(string, 'utf_16_le')


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


def get_command_args():
    """Parse command-line arguments"""

    parser = argparse.ArgumentParser(description='MS MAPI/OLE formatted email parser')

    parser.add_argument('-V', '--version', dest='version', required=False, action='store_false',
                        help='Print version and exit')
    parser.add_argument('-v', '--verbose', dest='verbose', required=False, action='store_false',
                        help='Enable verbose (e.g. INFO) logging')
    parser.add_argument('-s', '--server', dest='address', required=False, action='store',
                        help='IR-Flow server address')
    parser.add_argument('-u', '--api_user', dest='api_user', required=False, action='store',
                        help='IR-Flow API User')
    parser.add_argument('-k', '--api_key', dest='api_key', required=False, action='store',
                        help='IR-Flow API Key')
    parser.add_argument('-P', '--protocol', dest='protocol', required=False, action='store', default='https',
                        help='Configure http or https - defaults to https')
    parser.add_argument('-i', '--irflow_config', dest='irflow_config', required=False, action='store',
                        default='../irflow_api.conf',
                        help='Path to a valid IR-Flow configuration file')

    # Output flags
    parser.add_argument('--regex', dest='regex', required=False, action='store', default="True",
                        help='Return values to stdout for regex parsing')
    parser.add_argument('--debug', dest='debug', required=False, action='store_true',
                        help='Enable debug logging level')

    options = parser.parse_args()
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


def setup_logging(default_path=None, default_level=logging.WARN, env_key='EXTRACT_MSG_LOG_CFG'):
    """Setup logging configuration

    Args:
        default_path (str): Default path to use for the logging configuration file
        default_level (int): Default logging level
        env_key (str): Environment variable name to search for, for setting logfile path

    Returns:
        bool: True if the configuration file was found and applied, False otherwise
    """
    # Find logging.json if not provided
    if default_path:
        default_path = default_path
    else:
        default_path = '/home/irflow/irflow-integrations/integrations/logging.json'

    paths = [
        default_path,
        '../logging.json',
    ]

    path = None

    for config_path in paths:
        if os.path.exists(config_path):
            path = config_path
            break

    if path is None:
        print('Unable to find logging.json configuration file')
        print('Make sure a valid logging.json in in the extract_msg root directory')
        print(str(paths[1:]))
        logging.basicConfig(level=default_level)
        logger.warning('No logging.json found, using a basic configuration.')
        return False

    value = os.getenv(env_key, None)
    if value and os.path.exists(value):
        path = value

    with open(path, 'rt') as f:
        config = json.load(f)

    try:
        logging.config.dictConfig(config)
    except ValueError as e:
        print('Failed to find file - make sure your integrations log file is properly configured')
        print(e)

    return True
