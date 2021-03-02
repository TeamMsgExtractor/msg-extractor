"""
The constants used in extract_msg. If you modify any of these
without explicit instruction to do so from one of the
contributers, please do not complain about bugs.
"""

import datetime
import struct
import sys

import ebcdic

if sys.version_info[0] >= 3:
    BYTES = bytes
    STRING = str
else:
    BYTES = str
    STRING = unicode


# DEFINE CONSTANTS
# WARNING DO NOT CHANGE ANY OF THESE VALUES UNLESS YOU KNOW
# WHAT YOU ARE DOING! FAILURE TO FOLLOW THIS INSTRUCTION
# CAN AND WILL BREAK THIS SCRIPT!

# Constants used by named.py
NUMERICAL_NAMED = 0
STRING_NAMED = 1
GUID_PS_MAPI = '{00020328-0000-0000-C000-000000000046}'
GUID_PS_PUBLIC_STRINGS = '{00020329-0000-0000-C000-000000000046}'
GUID_PSETID_COMMON = '{00062008-0000-0000-C000-000000000046}'
GUID_PSETID_ADDRESS = '{00062004-0000-0000-C000-000000000046}'
GUID_PS_INTERNET_HEADERS = '{00020386-0000-0000-C000-000000000046}'
GUID_PSETID_APPOINTMENT = '{00062002-0000-0000-C000-000000000046}'
GUID_PSETID_MEETING = '{6ED8DA90-450B-101B-98DA-00AA003F1305}'
GUID_PSETID_LOG = '{0006200A-0000-0000-C000-000000000046}'
GUID_PSETID_MESSAGING = '{41F28F13-83F4-4114-A584-EEDB5A6B0BFF}'
GUID_PSETID_NOTE = '{0006200E-0000-0000-C000-000000000046}'
GUID_PSETID_POSTRSS = '{00062041-0000-0000-C000-000000000046}'
GUID_PSETID_TASK = '{00062003-0000-0000-C000-000000000046}'
GUID_PSETID_UNIFIEDMESSAGING = '{4442858E-A9E3-4E80-B900-317A210CC15B}'
GUID_PSETID_AIRSYNC = '{71035549-0739-4DCB-9163-00F0580DBBDF}'
GUID_PSETID_SHARING = '{00062040-0000-0000-C000-000000000046}'
GUID_PSETID_XMLEXTRACTEDENTITIES = '{23239608-685D-4732-9C55-4C95CB4E8E33}'
GUID_PSETID_ATTACHMENT = '{96357F7F-59E1-47D0-99A7-46515C183B54}'

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

# Multiple type properties that take up 2 bytes
MULTIPLE_2_BYTES = (
    '1002',
)

MULTIPLE_2_BYTES_HEX = (
    0x1002,
)

# Multiple type properties that take up 4 bytes
MULTIPLE_4_BYTES = (
    '1003',
    '1004',
)

MULTIPLE_4_BYTES_HEX = (
    0x1003,
    0x1004,
)

# Multiple type properties that take up 4 bytes
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

# Multiple type properties that take up 4 bytes
MULTIPLE_16_BYTES = (
    '1048',
)

MULTIPLE_16_BYTES_HEX = (
    0x1048,
)


# This is a dictionary matching the code page number to it's encoding name.
# The list used to make this can be found here:
# https://docs.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
### TODO:
# Many of these code pages are not supported by Python. As such, we should
# Really implement them ourselves to make sure that if someone wants to use an
# msg file with one of those encodings, they are able to. Perhaps we should
# create a seperate module for that?
# Code pages that currently don't have a supported encoding will be preceded by
# `# UNSUPPORTED`.
# For some of these, it is also possible that the name we are trying to find
# them with is not known to Python. I have already confirmed this for a few of
# them, and adjusted their names to ones that python would recognize. It is
# Possible I missed a few.
CODE_PAGES = {
    37: 'IBM037', # IBM EBCDIC US-Canada
    437: 'IBM437', # OEM United States
    500: 'IBM500', # IBM EBCDIC International
    708: 'ASMO-708', # Arabic (ASMO 708)
    # UNSUPPORTED
    709: '', # Arabic (ASMO-449+, BCON V4)
    # UNSUPPORTED
    710: '', # Arabic - Transparent Arabic
    # UNSUPPORTED
    720: 'DOS-720', # Arabic (Transparent ASMO); Arabic (DOS)
    737: 'cp737', # OEM Greek (formerly 437G); Greek (DOS)
    775: 'ibm775', # OEM Baltic; Baltic (DOS)
    850: 'ibm850', # OEM Multilingual Latin 1; Western European (DOS)
    852: 'ibm852', # OEM Latin 2; Central European (DOS)
    855: 'IBM855', # OEM Cyrillic (primarily Russian)
    857: 'ibm857', # OEM Turkish; Turkish (DOS)
    # UNSUPPORTED
    858: 'IBM00858', # OEM Multilingual Latin 1 + Euro symbol
    860: 'IBM860', # OEM Portuguese; Portuguese (DOS)
    861: 'ibm861', # OEM Icelandic; Icelandic (DOS)
    862: 'cp862', # OEM Hebrew; Hebrew (DOS)
    863: 'IBM863', # OEM French Canadian; French Canadian (DOS)
    864: 'IBM864', # OEM Arabic; Arabic (864)
    865: 'IBM865', # OEM Nordic; Nordic (DOS)
    866: 'cp866', # OEM Russian; Cyrillic (DOS)
    869: 'ibm869', # OEM Modern Greek; Greek, Modern (DOS)
    870: 'cp870', # IBM870 # IBM EBCDIC Multilingual/ROECE (Latin 2); IBM EBCDIC Multilingual Latin 2
    # UNSUPPORTED
    874: 'windows-874', # ANSI/OEM Thai (ISO 8859-11); Thai (Windows)
    875: 'cp875', # IBM EBCDIC Greek Modern
    932: 'shift_jis', # ANSI/OEM Japanese; Japanese (Shift-JIS)
    936: 'gb2312', # ANSI/OEM Simplified Chinese (PRC, Singapore); Chinese Simplified (GB2312)
    949: 'ks_c_5601-1987', # ANSI/OEM Korean (Unified Hangul Code)
    950: 'big5', # ANSI/OEM Traditional Chinese (Taiwan; Hong Kong SAR, PRC); Chinese Traditional (Big5)
    1026: 'IBM1026', # IBM EBCDIC Turkish (Latin 5)
    1047: 'cp1047', # IBM EBCDIC Latin 1/Open System
    1140: 'cp1140', # IBM EBCDIC US-Canada (037 + Euro symbol); IBM EBCDIC (US-Canada-Euro)
    1141: 'cp1141', # IBM EBCDIC Germany (20273 + Euro symbol); IBM EBCDIC (Germany-Euro)
    1142: 'cp1142', # IBM EBCDIC Denmark-Norway (20277 + Euro symbol); IBM EBCDIC (Denmark-Norway-Euro)
    1143: 'cp1143', # IBM EBCDIC Finland-Sweden (20278 + Euro symbol); IBM EBCDIC (Finland-Sweden-Euro)
    1144: 'cp1144', # IBM EBCDIC Italy (20280 + Euro symbol); IBM EBCDIC (Italy-Euro)
    1145: 'cp1145', # IBM EBCDIC Latin America-Spain (20284 + Euro symbol); IBM EBCDIC (Spain-Euro)
    1146: 'cp1146', # IBM EBCDIC United Kingdom (20285 + Euro symbol); IBM EBCDIC (UK-Euro)
    1147: 'cp1147', # IBM EBCDIC France (20297 + Euro symbol); IBM EBCDIC (France-Euro)
    1148: 'cp1148ms', # IBM EBCDIC International (500 + Euro symbol); IBM EBCDIC (International-Euro)
    1149: 'cp1149', # IBM EBCDIC Icelandic (20871 + Euro symbol); IBM EBCDIC (Icelandic-Euro)
    1200: 'utf-16-le', # Unicode UTF-16, little endian byte order (BMP of ISO 10646); available only to managed applications
    1201: 'utf-16-be', # Unicode UTF-16, big endian byte order; available only to managed applications
    1250: 'windows-1250', # ANSI Central European; Central European (Windows)
    1251: 'windows-1251', # ANSI Cyrillic; Cyrillic (Windows)
    1252: 'windows-1252', # ANSI Latin 1; Western European (Windows)
    1253: 'windows-1253', # ANSI Greek; Greek (Windows)
    1254: 'windows-1254', # ANSI Turkish; Turkish (Windows)
    1255: 'windows-1255', # ANSI Hebrew; Hebrew (Windows)
    1256: 'windows-1256', # ANSI Arabic; Arabic (Windows)
    1257: 'windows-1257', # ANSI Baltic; Baltic (Windows)
    1258: 'windows-1258', # ANSI/OEM Vietnamese; Vietnamese (Windows)
    1361: 'Johab', # Korean (Johab)
    10000: 'macintosh', # MAC Roman; Western European (Mac)
    10001: 'x-mac-japanese', # Japanese (Mac)
    # UNSUPPORTED
    10002: 'x-mac-chinesetrad', # MAC Traditional Chinese (Big5); Chinese Traditional (Mac)
    10003: 'x-mac-korean', # Korean (Mac)
    # UNSUPPORTED
    10004: 'x-mac-arabic', # Arabic (Mac)
    # UNSUPPORTED
    10005: 'x-mac-hebrew', # Hebrew (Mac)
    # UNSUPPORTED
    10006: 'x-mac-greek', # Greek (Mac)
    # UNSUPPORTED
    10007: 'x-mac-cyrillic', # Cyrillic (Mac)
    # UNSUPPORTED
    10008: 'x-mac-chinesesimp', # MAC Simplified Chinese (GB 2312); Chinese Simplified (Mac)
    # UNSUPPORTED
    10010: 'x-mac-romanian', # Romanian (Mac)
    # UNSUPPORTED
    10017: 'x-mac-ukrainian', # Ukrainian (Mac)
    # UNSUPPORTED
    10021: 'x-mac-thai', # Thai (Mac)
    # UNSUPPORTED
    10029: 'x-mac-ce', # MAC Latin 2; Central European (Mac)
    # UNSUPPORTED
    10079: 'x-mac-icelandic', # Icelandic (Mac)
    # UNSUPPORTED
    10081: 'x-mac-turkish', # Turkish (Mac)
    # UNSUPPORTED
    10082: 'x-mac-croatian', # Croatian (Mac)
    12000: 'utf-32', # Unicode UTF-32, little endian byte order; available only to managed applications
    12001: 'utf-32BE', # Unicode UTF-32, big endian byte order; available only to managed applications
    # UNSUPPORTED
    20000: 'x-Chinese_CNS', # CNS Taiwan; Chinese Traditional (CNS)
    # UNSUPPORTED
    20001: 'x-cp20001', # TCA Taiwan
    # UNSUPPORTED
    20002: 'x_Chinese-Eten', # Eten Taiwan; Chinese Traditional (Eten)
    # UNSUPPORTED
    20003: 'x-cp20003', # IBM5550 Taiwan
    # UNSUPPORTED
    20004: 'x-cp20004', # TeleText Taiwan
    # UNSUPPORTED
    20005: 'x-cp20005', # Wang Taiwan
    # UNSUPPORTED
    20105: 'x-IA5', # IA5 (IRV International Alphabet No. 5, 7-bit); Western European (IA5)
    # UNSUPPORTED
    20106: 'x-IA5-German', # IA5 German (7-bit)
    # UNSUPPORTED
    20107: 'x-IA5-Swedish', # IA5 Swedish (7-bit)
    # UNSUPPORTED
    20108: 'x-IA5-Norwegian', # IA5 Norwegian (7-bit)
    20127: 'us-ascii', # US-ASCII (7-bit)
    # UNSUPPORTED
    20261: 'x-cp20261', # T.61
    # UNSUPPORTED
    20269: 'x-cp20269', # ISO 6937 Non-Spacing Accent
    20273: 'IBM273', # IBM EBCDIC Germany
    20277: 'cp277', # IBM EBCDIC Denmark-Norway
    20278: 'cp278', # IBM EBCDIC Finland-Sweden
    20280: 'cp280', # IBM EBCDIC Italy
    20284: 'cp284', # IBM EBCDIC Latin America-Spain
    20285: 'cp285', # IBM EBCDIC United Kingdom
    20290: 'cp290', # IBM EBCDIC Japanese Katakana Extended
    20297: 'cp297', # IBM EBCDIC France
    20420: 'cp420', # IBM EBCDIC Arabic
    # UNSUPPORTED
    20423: 'IBM423', # IBM EBCDIC Greek
    20424: 'IBM424', # IBM EBCDIC Hebrew
    20833: 'cp833', # IBM EBCDIC Korean Extended
    20838: 'cp838', # IBM EBCDIC Thai
    20866: 'koi8-r', # Russian (KOI8-R); Cyrillic (KOI8-R)
    20871: 'cp871', # IBM EBCDIC Icelandic
    # UNSUPPORTED
    20880: 'IBM880', # IBM EBCDIC Cyrillic Russian
    # UNSUPPORTED
    20905: 'IBM905', # IBM EBCDIC Turkish
    # UNSUPPORTED
    20924: 'IBM00924', # IBM EBCDIC Latin 1/Open System (1047 + Euro symbol)
    20932: 'EUC-JP', # Japanese (JIS 0208-1990 and 0212-1990)
    # UNSUPPORTED
    20936: 'x-cp20936', # Simplified Chinese (GB2312); Chinese Simplified (GB2312-80)
    # UNSUPPORTED
    20949: 'x-cp20949', # Korean Wansung
    21025: 'cp1025', # IBM EBCDIC Cyrillic Serbian-Bulgarian
    # UNSUPPORTED
    21027: '', # (deprecated)
    21866: 'koi8-u', # Ukrainian (KOI8-U); Cyrillic (KOI8-U)
    28591: 'iso-8859-1', # ISO 8859-1 Latin 1; Western European (ISO)
    28592: 'iso-8859-2', # ISO 8859-2 Central European; Central European (ISO)
    28593: 'iso-8859-3', # ISO 8859-3 Latin 3
    28594: 'iso-8859-4', # ISO 8859-4 Baltic
    28595: 'iso-8859-5', # ISO 8859-5 Cyrillic
    28596: 'iso-8859-6', # ISO 8859-6 Arabic
    28597: 'iso-8859-7', # ISO 8859-7 Greek
    28598: 'iso-8859-8', # ISO 8859-8 Hebrew; Hebrew (ISO-Visual)
    28599: 'iso-8859-9', # ISO 8859-9 Turkish
    28603: 'iso-8859-13', # ISO 8859-13 Estonian
    28605: 'iso-8859-15', # ISO 8859-15 Latin 9
    # UNSUPPORTED
    29001: 'x-Europa', # Europa 3
    # UNSUPPORTED
    38598: 'iso-8859-8-i', # ISO 8859-8 Hebrew; Hebrew (ISO-Logical)
    50220: 'iso-2022-jp', # ISO 2022 Japanese with no halfwidth Katakana; Japanese (JIS)
    50221: 'csISO2022JP', # ISO 2022 Japanese with halfwidth Katakana; Japanese (JIS-Allow 1 byte Kana)
    50222: 'iso-2022-jp', # ISO 2022 Japanese JIS X 0201-1989; Japanese (JIS-Allow 1 byte Kana - SO/SI)
    50225: 'iso-2022-kr', # ISO 2022 Korean
    # UNSUPPORTED
    50227: 'x-cp50227', # ISO 2022 Simplified Chinese; Chinese Simplified (ISO 2022)
    # UNSUPPORTED
    50229: '', # ISO 2022 Traditional Chinese
    # UNSUPPORTED
    50930: '', # EBCDIC Japanese (Katakana) Extended
    # UNSUPPORTED
    50931: '', # EBCDIC US-Canada and Japanese
    # UNSUPPORTED
    50933: '', # EBCDIC Korean Extended and Korean
    # UNSUPPORTED
    50935: '', # EBCDIC Simplified Chinese Extended and Simplified Chinese
    # UNSUPPORTED
    50936: '', # EBCDIC Simplified Chinese
    # UNSUPPORTED
    50937: '', # EBCDIC US-Canada and Traditional Chinese
    # UNSUPPORTED
    50939: '', # EBCDIC Japanese (Latin) Extended and Japanese
    51932: 'euc-jp', # EUC Japanese
    51936: 'EUC-CN', # EUC Simplified Chinese; Chinese Simplified (EUC)
    51949: 'euc-kr', # EUC Korean
    # UNSUPPORTED
    51950: '', # EUC Traditional Chinese
    52936: 'hz-gb-2312', # HZ-GB2312 Simplified Chinese; Chinese Simplified (HZ)
    54936: 'GB18030', # Windows XP and later: GB18030 Simplified Chinese (4 byte); Chinese Simplified (GB18030)
    # UNSUPPORTED
    57002: 'x-iscii-de', # ISCII Devanagari
    # UNSUPPORTED
    57003: 'x-iscii-be', # ISCII Bangla
    # UNSUPPORTED
    57004: 'x-iscii-ta', # ISCII Tamil
    # UNSUPPORTED
    57005: 'x-iscii-te', # ISCII Telugu
    # UNSUPPORTED
    57006: 'x-iscii-as', # ISCII Assamese
    # UNSUPPORTED
    57007: 'x-iscii-or', # ISCII Odia
    # UNSUPPORTED
    57008: 'x-iscii-ka', # ISCII Kannada
    # UNSUPPORTED
    57009: 'x-iscii-ma', # ISCII Malayalam
    # UNSUPPORTED
    57010: 'x-iscii-gu', # ISCII Gujarati
    # UNSUPPORTED
    57011: 'x-iscii-pa', # ISCII Punjabi
    65000: 'utf-7', # Unicode (UTF-7)
    65001: 'utf-8', # Unicode (UTF-8)
}

INTELLIGENCE_DUMB = 0
INTELLIGENCE_SMART = 1
INTELLIGENCE_TUPLE = (
    'INTELLIGENCE_DUMB',
    'INTELLIGENCE_SMART',
)

TYPE_MESSAGE = 0
TYPE_MESSAGE_EMBED = 1
TYPE_ATTACHMENT = 2
TYPE_RECIPIENT = 3
TYPE_TUPLE = (
    'TYPE_MESSAGE',
    'TYPE_MESSAGE_EMBED',
    'TYPE_ATTACHMENT',
    'TYPE_RECIPIENT',
)

RECIPIENT_SENDER = 0
RECIPIENT_TO = 1
RECIPIENT_CC = 2
RECIPIENT_BCC = 3
RECIPIENT_TUPLE = (
    'RECIPIENT_SENDER',
    'RECIPIENT_TO',
    'RECIPIENT_CC',
    'RECIPIENT_BCC',
)

# PidTagImportance
IMPORTANCE_LOW = 0
IMPORTANCE_MEDIUM = 1
IMPORTANCE_HIGH = 2
IMPORTANCE_TUPLE = (
    'IMPORTANCE_LOW',
    'IMPORTANCE_MEDIUM',
    'IMPORTANCE_HIGH',
)

# PidTagSensitivity
SENSITIVITY_NORMAL = 0
SENSITIVITY_PERSONAL = 1
SENSITIVITY_PRIVATE = 2
SENSITIVITY_CONFIDENTIAL = 3
SENSITIVITY_TUPLE = (
    'SENSITIVITY_NORMAL',
    'SENSITIVITY_PERSONAL',
    'SENSITIVITY_PRIVATE',
    'SENSITIVITY_CONFIDENTIAL',
)

# PidTagPriority
PRIORITY_URGENT = 0x00000001
PRIORITY_NORMAL = 0x00000000
PRIORITY_NOT_URGENT = 0xFFFFFFFF

PYTPFLOATINGTIME_START = datetime.datetime(1899, 12, 30)

# Constants used for argparse stuff
KNOWN_FILE_FLAGS = [
    '--out-name',
]
NEEDS_ARG = [
    '--out-name',
]
MAINDOC = "extract_msg:\n\tExtracts emails and attachments saved in Microsoft Outlook's .msg files.\n\n" \
          "https://github.com/mattgwwalker/msg-extractor"

# Define pre-compiled structs to make unpacking slightly faster
# General structs
ST1 = struct.Struct('<8x4I')
ST2 = struct.Struct('<H2xI8x')
ST3 = struct.Struct('<Q')
# Structs used by data.py
ST_DATA_UI32 = struct.Struct('<I')
ST_DATA_UI16 = struct.Struct('<H')
ST_DATA_UI8 = struct.Struct('<B')
# Structs used by named.py
STNP_NAM = struct.Struct('<i')
STNP_ENT = struct.Struct('<IHH') # Struct used for unpacking the entries in the entry stream
# Structs used by prop.py
STFIX = struct.Struct('<8x8s')
STVAR = struct.Struct('<8xi4s')
# Structs to help with email type to python type conversions
STI16 = struct.Struct('<h6x')
STI32 = struct.Struct('<I4x')
STI64 = struct.Struct('<q')
STF32 = struct.Struct('<f4x')
STF64 = struct.Struct('<d')
STUI32 = struct.Struct('<I4x')
STMI16 = struct.Struct('<h')
STMI32 = struct.Struct('<i')
STMI64 = struct.Struct('<q')
STMF32 = struct.Struct('<f')
STMF64 = struct.Struct('<d')
# PermanentEntryID parsing struct
STPEID = struct.Struct('<B3x16s4xI')

PTYPES = {
    0x0000: 'PtypUnspecified',
    0x0001: 'PtypNull',
    0x0002: 'PtypInteger16',  # Signed short
    0x0003: 'PtypInteger32',  # Signed int
    0x0004: 'PtypFloating32',  # Float
    0x0005: 'PtypFloating64',  # Double
    0x0006: 'PtypCurrency',
    0x0007: 'PtypFloatingTime',
    0x000A: 'PtypErrorCode',
    0x000B: 'PtypBoolean',
    0x000D: 'PtypObject/PtypEmbeddedTable/Storage',
    0x0014: 'PtypInteger64',  # Signed longlong
    0x001E: 'PtypString8',
    0x001F: 'PtypString',
    0x0040: 'PtypTime',  # Use msgEpoch to convert to unix time stamp
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

# Display types
DT_MAILUSER = 0x0000
DT_DISTLIST = 0x0001
DT_FORUM = 0x0002
DT_AGENT = 0x0003
DT_ORGANIZATION = 0x0004
DT_PRIVATE_DISTLIST = 0x0005
DT_REMOTE_MAILUSER = 0x0006
DT_CONTAINER = 0x0100
DT_TEMPLATE = 0x0101
DT_ADDRESS_TEMPLATE = 0x0102
DT_SEARCH = 0x0200

# Rule action types
RA_OP_MOVE = 0x01
RA_OP_COPY = 0x02
RA_OP_REPLY = 0x03
RA_OP_OOF_REPLY = 0x04
RA_OP_DEFER_ACTION = 0x05
RA_OP_BOUNCE = 0x06
RA_OP_FORWARD = 0x07
RA_OP_DELEGATE = 0x08
RA_OP_TAG = 0x09
RA_OP_DELETE = 0x0A
RA_OP_MARK_AS_READ = 0x0B

# Recipiet Row Flag Types
RF_NOTYPE = 0x0
RF_X500DN = 0x1
RF_MSMAIL = 0x2
RF_SMTP = 0x3
RF_FAX = 0x4
RF_PROFESSIONALOFFICESYSTEM = 0x5
RF_PERSONALDESTRIBUTIONLIST1 = 0x6
RF_PERSONALDESTRIBUTIONLIST2 = 0x7

# Attachment Error behavior types
ATTACHMENT_ERROR_THROW = 0
ATTACHMENT_ERROR_NOT_IMPLEMENTED = 1
ATTACHMENT_ERROR_BROKEN = 2


# This property information was sourced from
# http://www.fileformat.info/format/outlookmsg/index.htm
# on 2013-07-22.
# It was extended by The Elemental of Destruction on 2018-10-12
PROPERTIES = {
    '00010102': 'Template data',
    '0002000B': 'Alternate recipient allowed',
    '0004001F': 'Auto forward comment',
    '00040102': 'Script data',
    '0005000B': 'Auto forwarded',
    '000F000F': 'Deferred delivery time',
    '00100040': 'Deliver time',
    '00150040': 'Expiry time',
    '00170003': 'Importance',
    '001A001F': 'Message class',
    '0023001F': 'Originator delivery report requested',
    '00250102': 'Parent key',
    '00260003': 'Priority',
    '0029000B': 'Read receipt requested',
    '002A0040': 'Receipt time',
    '002B000B': 'Recipient reassignment prohibited',
    '002E0003': 'Original sensitivity',
    '00300040': 'Reply time',
    '00310102': 'Report tag',
    '00320040': 'Report time',
    '00360003': 'Sensitivity',
    '0037001F': 'Subject',
    '00390040': 'Client Submit Time',
    '003A001F': '',
    '003B0102': '',
    '003D001F': 'Subject prefix',
    '003F0102': '',
    '0040001F': 'Received by name',
    '00410102': '',
    '0042001F': 'Sent repr name',
    '00430102': '',
    '0044001F': 'Rcvd repr name',
    '00450102': '',
    '0046001F': '',
    '00470102': '',
    '0049001F': '',
    '004B001F': '',
    '004C0102': '',
    '004D001F': 'Org author name',
    '004E0040': '',
    '004F0102': '',
    '0050001F': 'Reply rcipnt names',
    '00510102': '',
    '00520102': '',
    '00530102': '',
    '00540102': '',
    '00550040': '',
    '0057000B': '',
    '0058000B': '',
    '0059000B': '',
    '005A001F': 'Org sender name',
    '005B0102': '',
    '005C0102': '',
    '005D001F': '',
    '005E0102': '',
    '005F0102': '',
    '00600040': '',
    '00610040': '',
    '00620003': '',
    '0063000B': '',
    '0064001F': 'Sent repr adrtype',
    '0065001F': 'Sent repr email',
    '0066001F': '',
    '00670102': '',
    '0068001F': '',
    '0069001F': '',
    '0070001F': 'Topic',
    '00710102': '',
    '0072001F': '',
    '0073001F': '',
    '0074001F': '',
    '0075001F': 'Rcvd by adrtype',
    '0076001F': 'Rcvd by email',
    '0077001F': 'Repr adrtype',
    '0078001F': 'Repr email',
    '007D001F': 'Message header',
    '007F0102': '',
    '0080001F': '',
    '0081001F': '',
    '08070003': '',
    '0809001F': '',
    '0C040003': '',
    '0C050003': '',
    '0C06000B': '',
    '0C08000B': '',
    '0C150003': '',
    '0C17000B': '',
    '0C190102': '',
    '0C1A001F': 'Sender name',
    '0C1B001F': '',
    '0C1D0102': '',
    '0C1E001F': 'Sender adr type',
    '0C1F001F': 'Sender email',
    '0C200003': '',
    '0C21001F': '',
    '0E01000B': '',
    '0E02001F': 'Display BCC',
    '0E03001F': 'Display CC',
    '0E04001F': 'Display To',
    '0E060040': '',
    '0E070003': '',
    '0E080003': '',
    '0E080014': '',
    '0E090102': '',
    '0E0F000B': '',
    '0E12000D': '',
    '0E13000D': '',
    '0E170003': '',
    '0E1B000B': '',
    '0E1D001F': 'Subject (normalized)',
    '0E1F000B': '',
    '0E200003': '',
    '0E210003': '',
    '0E28001F': 'Recvd account1 (uncertain)',
    '0E29001F': 'Recvd account2 (uncertain)',
    '1000001F': 'Message body',
    '1008': 'RTF sync body tag', # Where did this come from ??? It's not listed in the docs
    '10090102': 'Compressed RTF body',
    '1013001F': 'HTML body',
    '1035001F': 'Message ID (uncertain)',
    '1046001F': 'Sender email (uncertain)',
    '3001001F': 'Display name',
    '3002001F': 'Address type',
    '3003001F': 'Email address',
    '30070040': 'Creation date',
    '39FE001F': '7-bit email (uncertain)',
    '39FF001F': '7-bit display name',

    # Attachments (37xx)
    '37010102': 'Attachment data',
    '37020102': '',
    '3703001F': 'Attachment extension',
    '3704001F': 'Attachment short filename',
    '37050003': 'Attachment attach method',
    '3707001F': 'Attachment long filename',
    '370E001F': 'Attachment mime tag',
    '3712001F': 'Attachment ID (uncertain)',

    # Address book (3Axx):
    '3A00001F': 'Account',
    '3A02001F': 'Callback phone no',
    '3A05001F': 'Generation',
    '3A06001F': 'Given name',
    '3A08001F': 'Business phone',
    '3A09001F': 'Home phone',
    '3A0A001F': 'Initials',
    '3A0B001F': 'Keyword',
    '3A0C001F': 'Language',
    '3A0D001F': 'Location',
    '3A11001F': 'Surname',
    '3A15001F': 'Postal address',
    '3A16001F': 'Company name',
    '3A17001F': 'Title',
    '3A18001F': 'Department',
    '3A19001F': 'Office location',
    '3A1A001F': 'Primary phone',
    '3A1B101F': 'Business phone 2',
    '3A1C001F': 'Mobile phone',
    '3A1D001F': 'Radio phone no',
    '3A1E001F': 'Car phone no',
    '3A1F001F': 'Other phone',
    '3A20001F': 'Transmit dispname',
    '3A21001F': 'Pager',
    '3A220102': 'User certificate',
    '3A23001F': 'Primary Fax',
    '3A24001F': 'Business Fax',
    '3A25001F': 'Home Fax',
    '3A26001F': 'Country',
    '3A27001F': 'Locality',
    '3A28001F': 'State/Province',
    '3A29001F': 'Street address',
    '3A2A001F': 'Postal Code',
    '3A2B001F': 'Post Office Box',
    '3A2C001F': 'Telex',
    '3A2D001F': 'ISDN',
    '3A2E001F': 'Assistant phone',
    '3A2F001F': 'Home phone 2',
    '3A30001F': 'Assistant',
    '3A44001F': 'Middle name',
    '3A45001F': 'Dispname prefix',
    '3A46001F': 'Profession',
    '3A47001F': '',
    '3A48001F': 'Spouse name',
    '3A4B001F': 'TTYTTD radio phone',
    '3A4C001F': 'FTP site',
    '3A4E001F': 'Manager name',
    '3A4F001F': 'Nickname',
    '3A51001F': 'Business homepage',
    '3A57001F': 'Company main phone',
    '3A58101F': 'Childrens names',
    '3A59001F': 'Home City',
    '3A5A001F': 'Home Country',
    '3A5B001F': 'Home Postal Code',
    '3A5C001F': 'Home State/Provnce',
    '3A5D001F': 'Home Street',
    '3A5F001F': 'Other adr City',
    '3A60': 'Other adr Country',
    '3A61': 'Other adr PostCode',
    '3A62': 'Other adr Province',
    '3A63': 'Other adr Street',
    '3A64': 'Other adr PO box',

    '3FF7': 'Server (uncertain)',
    '3FF8': 'Creator1 (uncertain)',
    '3FFA': 'Creator2 (uncertain)',
    '3FFC': 'To email (uncertain)',
    '403D': 'To adrtype (uncertain)',
    '403E': 'To email (uncertain)',
    '5FF6': 'To (uncertain)',
}


# END CONSTANTS

def int_to_data_type(integer):
    """
    Returns the name of the data type constant that has the value of :param integer:
    """
    return TYPE_TUPLE[integer]


def int_to_intelligence(integer):
    """
    Returns the name of the intelligence level constant that has the value of :param integer:
    """
    return INTELLIGENCE_TUPLE[integer]

def int_to_recipient_type(integer):
    """
    Returns the name of the recipient type constant that has the value of :param integer:
    """
    return RECIPIENT_TUPLE[integer]
