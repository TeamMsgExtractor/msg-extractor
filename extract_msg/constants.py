import struct

# DEFINE CONSTANTS
# WARNING DO NOT CHANGE ANY OF THESE VALUES UNLESS YOU KNOW
# WHAT YOU ARE DOING! FAILURE TO FOLLOW THIS INSTRUCTION
# CAN AND WILL BREAK THIS SCRIPT!

INTELLIGENCE_DUMB  = 0
INTELLIGENCE_SMART = 1
INTELLIGENCE_DICT = {
    INTELLIGENCE_DUMB:  'INTELLIGENCE_DUMB',
    INTELLIGENCE_SMART: 'INTELLIGENCE_SMART',
}

TYPE_MESSAGE       = 0
TYPE_MESSAGE_EMBED = 1
TYPE_ATTACHMENT    = 2
TYPE_RECIPIENT     = 3
TYPE_DICT = {
    TYPE_MESSAGE:       'TYPE_MESSAGE',
    TYPE_MESSAGE_EMBED: 'TYPE_MESSAGE_EMBED',
    TYPE_ATTACHMENT:    'TYPE_ATTACHMENT',
    TYPE_RECIPIENT:     'TYPE_RECIPIENT',
}


RECIPIENT_SENDER   = 0
RECIPIENT_TO       = 1
RECIPIENT_CC       = 2
RECIPIENT_BCC      = 3
RECIPIENT_DICT = {
    RECIPIENT_SENDER:   'RECIPIENT_SENDER',
    RECIPIENT_TO:       'RECIPIENT_TO',
    RECIPIENT_CC:       'RECIPIENT_CC',
    RECIPIENT_BCC:      'RECIPIENT_BCC',
}

# Define pre-compiled structs to make unpacking slightly faster
ST1                = struct.Struct('<8x4I')
ST2                = struct.Struct('<H2xI8s')
ST3                = struct.Struct('<Q')
STI16              = struct.Struct('<h6x')
STI32              = struct.Struct('<i4x')
STI64              = struct.Struct('<q')
STF32              = struct.Struct('<f4x')
STF64              = struct.Struct('<d')

PTYPES = {
    0x0000: 'PtypUnspecified',
    0x0001: 'PtypNull',
    0x0002: 'PtypInteger16', # Signed short
    0x0003: 'PtypInteger32', # Signed int
    0x0004: 'PtypFloating32', # Float
    0x0005: 'PtypFloating64', # Double
    0x0006: 'PtypCurrency',
    0x0007: 'PtypFloatingTime',
    0x000A: 'PtypErrorCode',
    0x000B: 'PtypBoolean',
    0x000D: 'PtypObject/PtypEmbeddedTable',
    0x0014: 'PtypInteger64', # Signed longlong
    0x001E: 'PtypString8',
    0x001F: 'PtypString',
    0x0040: 'PtypTime', # Use msgEpoch to convert to unix time stamp
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
