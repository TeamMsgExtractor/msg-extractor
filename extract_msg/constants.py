"""
The constants used in extract_msg. If you modify any of these
without explicit instruction to do so from one of the
contributers, please do not complain about bugs.
"""

import struct



# DEFINE CONSTANTS
# WARNING DO NOT CHANGE ANY OF THESE VALUES UNLESS YOU KNOW
# WHAT YOU ARE DOING! FAILURE TO FOLLOW THIS INSTRUCTION
# CAN AND WILL BREAK THIS SCRIPT!

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

INTELLIGENCE_DUMB = 0
INTELLIGENCE_SMART = 1
INTELLIGENCE_DICT = {
    INTELLIGENCE_DUMB: 'INTELLIGENCE_DUMB',
    INTELLIGENCE_SMART: 'INTELLIGENCE_SMART',
}

TYPE_MESSAGE = 0
TYPE_MESSAGE_EMBED = 1
TYPE_ATTACHMENT = 2
TYPE_RECIPIENT = 3
TYPE_DICT = {
    TYPE_MESSAGE: 'TYPE_MESSAGE',
    TYPE_MESSAGE_EMBED: 'TYPE_MESSAGE_EMBED',
    TYPE_ATTACHMENT: 'TYPE_ATTACHMENT',
    TYPE_RECIPIENT: 'TYPE_RECIPIENT',
}

RECIPIENT_SENDER = 0
RECIPIENT_TO = 1
RECIPIENT_CC = 2
RECIPIENT_BCC = 3
RECIPIENT_DICT = {
    RECIPIENT_SENDER: 'RECIPIENT_SENDER',
    RECIPIENT_TO: 'RECIPIENT_TO',
    RECIPIENT_CC: 'RECIPIENT_CC',
    RECIPIENT_BCC: 'RECIPIENT_BCC',
}

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
# Structs used by prop.py
STFIX = struct.Struct('<8x8s')
STVAR = struct.Struct('<8xi4s')
# Structs to help with email type to python type conversions
STI16 = struct.Struct('<h6x')
STI32 = struct.Struct('<i4x')
STI64 = struct.Struct('<q')
STF32 = struct.Struct('<f4x')
STF64 = struct.Struct('<d')

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

# This property information was sourced from
# http://www.fileformat.info/format/outlookmsg/index.htm
# on 2013-07-22.
# It was extended by The Elemental of Creation on 2018-10-12
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

def int_to_recipient_type(integer):
    """
    Returns the name of the recipient type constant that has the value of :param integer:
    """
    return RECIPIENT_DICT[integer]


def int_to_data_type(integer):
    """
    Returns the name of the data type constant that has the value of :param integer:
    """
    return TYPE_DICT[integer]


def int_to_intelligence(integer):
    """
    Returns the name of the intelligence level constant that has the value of :param integer:
    """
    return INTELLIGENCE_DICT[integer]
