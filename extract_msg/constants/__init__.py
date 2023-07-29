"""
The constants used in extract_msg. If you modify any of these without explicit
instruction to do so from one of the contributers, please do not complain about
bugs.
"""

__all__ = [
    # Modules.
    'ps',
    're',
    'st',

    # Constants.
    'DEFAULT_CLSID',
    'FIXED_LENGTH_PROPS',
    'FIXED_LENGTH_PROPS_STRING',
    'HEADER_FORMAT',
    'HEADER_FORMAT_TYPE',
    'HEADER_FORMAT_VALUE_TYPE',
    'KNOWN_CLASS_TYPES',
    'KNOWN_FILE_FLAGS',
    'MAINDOC',
    'MULTIPLE_16_BYTES',
    'MULTIPLE_16_BYTES_HEX',
    'MULTIPLE_2_BYTES',
    'MULTIPLE_2_BYTES_HEX',
    'MULTIPLE_4_BYTES',
    'MULTIPLE_4_BYTES_HEX',
    'MULTIPLE_8_BYTES',
    'MULTIPLE_8_BYTES_HEX',
    'NEEDS_ARG',
    'NULL_DATE',
    'PTYPES',
    'PYTPFLOATINGTIME_START',
    'REFUSED_CLASS_TYPES',
    'REPOSITORY_URL',
    'SAVE_TYPE',
    'VARIABLE_LENGTH_PROPS',
    'VARIABLE_LENGTH_PROPS_STRING',
]


import datetime

from typing import Dict, List, Tuple, Union

from . import ps, re, st
from ..enums import SaveType


# Typing Constants.
HEADER_FORMAT_VALUE_TYPE = Union[str, Tuple[Union[str, None], bool], None]
# Basically a dict of HEADER_FORMAT_TYPE and dicts containing them.
HEADER_FORMAT_TYPE = Dict[str, Union[HEADER_FORMAT_VALUE_TYPE, Dict[str, HEADER_FORMAT_VALUE_TYPE]]]
SAVE_TYPE = Tuple[SaveType, Union[List[str], str, None]]



FIXED_LENGTH_PROPS = (
    0x0000,
    0x0001,
    0x0002,
    0x0003,
    0x0004,
    0x0005,
    0x0006,
    0x0007,
    0x000A,
    0x000B,
    0x0014,
    0x0040,
    0x0048,
)

FIXED_LENGTH_PROPS_STRING = (
    '0000',
    '0001',
    '0002',
    '0003',
    '0004',
    '0005',
    '0006',
    '0007',
    '000A',
    '000B',
    '0014',
    '0040',
    '0048',
)

VARIABLE_LENGTH_PROPS = (
    0x000D,
    0x001E,
    0x001F,
    0x00FB,
    0x00FD,
    0x00FE,
    0X0102,
    0x1002,
    0x1003,
    0x1004,
    0x1005,
    0x1006,
    0x1007,
    0x1014,
    0x101E,
    0x101F,
    0x1040,
    0x1048,
    0x1102,
)

VARIABLE_LENGTH_PROPS_STRING = (
    '000D',
    '001E',
    '001F',
    '00FB',
    '00FD',
    '00FE',
    '0102',
    '1002',
    '1003',
    '1004',
    '1005',
    '1006',
    '1007',
    '1014',
    '101E',
    '101F',
    '1040',
    '1048',
    '1102',
)

# Multiple type properties that take up 2 bytes.
MULTIPLE_2_BYTES = (
    '1002',
)

MULTIPLE_2_BYTES_HEX = (
    0x1002,
)

# Multiple type properties that take up 4 bytes.
MULTIPLE_4_BYTES = (
    '1003',
    '1004',
)

MULTIPLE_4_BYTES_HEX = (
    0x1003,
    0x1004,
)

# Multiple type properties that take up 8 bytes.
MULTIPLE_8_BYTES = (
    '1005',
    '1007',
    '1014',
    '1040',
)

MULTIPLE_8_BYTES_HEX = (
    0x1005,
    0x1007,
    0x1014,
    0x1040,
)

# Multiple type properties that take up 16 bytes.
MULTIPLE_16_BYTES = (
    '1048',
)

MULTIPLE_16_BYTES_HEX = (
    0x1048,
)


# Used to format the header for saving only the header.
HEADER_FORMAT = """From: {From}
To: {To}
Cc: {Cc}
Bcc: {Bcc}
Subject: {subject}
Date: {Date}
Message-ID: {Message-Id}
"""


KNOWN_CLASS_TYPES = (
    'ipm.activity',
    'ipm.appointment', # [MS-OXOCAL]
    'ipm.contact', # [MS-OXOCNTC]
    'ipm.configuration', # [MS-OXOCFG]
    'ipm.distlist',
    'ipm.document',
    'ipm.ole.class',
    'ipm.outlook.recall',
    'ipm.note',
    'ipm.post',
    'ipm.stickynote',
    'ipm.recall.report',
    'ipm.remote',
    'ipm.report',
    'ipm.resend',
    'ipm.schedule',
    'ipm.task',
    'ipm.taskrequest',
    'report',
)

# Each item is a tuple of the lowercase class type and the issue number
# associated with it.
REFUSED_CLASS_TYPES = (
    ('ipm.outlook.recall', '235'),
)

PYTPFLOATINGTIME_START = datetime.datetime(1899, 12, 30)
NULL_DATE = datetime.datetime(4500, 8, 31, 23, 59)

# Constants used for argparse stuff.
KNOWN_FILE_FLAGS = (
    '--out-name',
)
NEEDS_ARG = (
    '--out-name',
)
REPOSITORY_URL = 'https://github.com/TeamMsgExtractor/msg-extractor'
MAINDOC = f"""extract_msg:
\tExtracts emails and attachments saved in Microsoft Outlook's .msg files.

{REPOSITORY_URL}"""

# Default class ID for the root entry for OleWriter. This should be
# referencing Outlook if I understand it correctly.
DEFAULT_CLSID = b'\x0b\r\x02\x00\x00\x00\x00\x00\xc0\x00\x00\x00\x00\x00\x00F'

PTYPES = {
    0x0000: 'PtypUnspecified',
    0x0001: 'PtypNull',
    0x0002: 'PtypInteger16',  # Signed short.
    0x0003: 'PtypInteger32',  # Signed int.
    0x0004: 'PtypFloating32',  # Float.
    0x0005: 'PtypFloating64',  # Double.
    0x0006: 'PtypCurrency',
    0x0007: 'PtypFloatingTime',
    0x000A: 'PtypErrorCode',
    0x000B: 'PtypBoolean',
    0x000D: 'PtypObject/PtypEmbeddedTable/Storage',
    0x0014: 'PtypInteger64',  # Signed longlong.
    0x001E: 'PtypString8',
    0x001F: 'PtypString',
    0x0040: 'PtypTime', # Use filetimeToUtc to convert to unix time stamp.
    0x0048: 'PtypGuid',
    0x00FB: 'PtypServerId',
    0x00FD: 'PtypRestriction',
    0x00FE: 'PtypRuleAction',
    0x0102: 'PtypBinary',
    0x1002: 'PtypMultipleInteger16',
    0x1003: 'PtypMultipleInteger32',
    0x1004: 'PtypMultipleFloating32',
    0x1005: 'PtypMultipleFloating64',
    0x1006: 'PtypMultipleCurrency',
    0x1007: 'PtypMultipleFloatingTime',
    0x1014: 'PtypMultipleInteger64',
    0x101E: 'PtypMultipleString8',
    0x101F: 'PtypMultipleString',
    0x1040: 'PtypMultipleTime',
    0x1048: 'PtypMultipleGuid',
    0x1102: 'PtypMultipleBinary',
}

# END CONSTANTS.