"""
Struct constants.
"""

__all__ = [
    'ST_BC_FIELD_INFO',
    'ST_BC_HEAD',
    'ST_BE_F32',
    'ST_BE_F64',
    'ST_BE_I16',
    'ST_BE_I32',
    'ST_BE_I64',
    'ST_BE_I8',
    'ST_BE_UI16',
    'ST_BE_UI32',
    'ST_BE_UI64',
    'ST_BE_UI8',
    'ST_CF_DIR_ENTRY',
    'ST_GUID',
    'ST_LE_F32',
    'ST_LE_F64',
    'ST_LE_I16',
    'ST_LE_I32',
    'ST_LE_I64',
    'ST_LE_I8',
    'ST_LE_UI16',
    'ST_LE_UI32',
    'ST_LE_UI64',
    'ST_LE_UI8',
    'ST_NP_ENT',
    'ST_PEID',
    'ST_PROP_BASE',
    'ST_PROP_VAR',
    'ST_PROPSTORE_HEADER',
    'ST_RGB',
    'ST_SBO_I8',
    'ST_SBO_I16',
    'ST_SBO_I32',
    'ST_SBO_I64',
    'ST_SBO_UI8',
    'ST_SBO_UI16',
    'ST_SBO_UI32',
    'ST_SBO_UI64',
    'ST_SYSTEMTIME',
    'ST_TZ',
]


import struct

from typing import Final


# Define pre-compiled structs to make unpacking slightly faster.
# General structs.
ST_PROPSTORE_HEADER: Final[struct.Struct] = struct.Struct('<8x4I')
ST_PROP_BASE: Final[struct.Struct] = struct.Struct('<2HI')
# Struct used for unpacking a system time.
ST_SYSTEMTIME: Final[struct.Struct] = struct.Struct('<8H')
# Struct used for unpacking a GUID from bytes.
ST_GUID: Final[struct.Struct] = struct.Struct('<IHH8s')
# Struct for unpacking a TimeZoneStruct from bytes.
ST_TZ: Final[struct.Struct] = struct.Struct('<iiiH16sH16s')
# Struct for packing a compount file directory entry.
ST_CF_DIR_ENTRY: Final[struct.Struct] = struct.Struct('<64sHBBIII16sIQQIQ')
# Struct used for unpacking the entries in the entry stream
ST_NP_ENT: Final[struct.Struct] = struct.Struct('<IHH')
# Structs used by prop.py
ST_PROP_VAR: Final[struct.Struct] = struct.Struct('<2I')
# PermanentEntryID parsing struct
ST_PEID: Final[struct.Struct] = struct.Struct('<B3x16s4xI')
# Struct for unpacking the first part of the BusinessCardDisplayDefinition
# structure.
ST_BC_HEAD: Final[struct.Struct] = struct.Struct('BBBBBBBBBBBxB')
# Struct for completely unpacking the FieldInfo structure.
ST_BC_FIELD_INFO: Final[struct.Struct] = struct.Struct('HBBBxHBBBxBBBx')
# Structs for parsing basic types.
ST_LE_I8: Final[struct.Struct] = struct.Struct('<b')
ST_LE_I16: Final[struct.Struct] = struct.Struct('<h')
ST_LE_I32: Final[struct.Struct] = struct.Struct('<i')
ST_LE_I64: Final[struct.Struct] = struct.Struct('<q')
ST_LE_UI8: Final[struct.Struct] = struct.Struct('<B')
ST_LE_UI16: Final[struct.Struct] = struct.Struct('<H')
ST_LE_UI32: Final[struct.Struct] = struct.Struct('<I')
ST_LE_UI64: Final[struct.Struct] = struct.Struct('<Q')
ST_LE_F32: Final[struct.Struct] = struct.Struct('<f')
ST_LE_F64: Final[struct.Struct] = struct.Struct('<d')
ST_BE_I8: Final[struct.Struct] = struct.Struct('>b')
ST_BE_I16: Final[struct.Struct] = struct.Struct('>h')
ST_BE_I32: Final[struct.Struct] = struct.Struct('>i')
ST_BE_I64: Final[struct.Struct] = struct.Struct('>q')
ST_BE_UI8: Final[struct.Struct] = struct.Struct('>B')
ST_BE_UI16: Final[struct.Struct] = struct.Struct('>H')
ST_BE_UI32: Final[struct.Struct] = struct.Struct('>I')
ST_BE_UI64: Final[struct.Struct] = struct.Struct('>Q')
ST_BE_F32: Final[struct.Struct]  = struct.Struct('>f')
ST_BE_F64: Final[struct.Struct] = struct.Struct('>d')
# Structs that use the system byte order, where consistency on a single system
# is all that matters. Mainly used for quick casts between signed and unsigned.
ST_SBO_I8: Final[struct.Struct] = struct.Struct('@b')
ST_SBO_I16: Final[struct.Struct] = struct.Struct('@h')
ST_SBO_I32: Final[struct.Struct] = struct.Struct('@i')
ST_SBO_I64: Final[struct.Struct] = struct.Struct('@q')
ST_SBO_UI8: Final[struct.Struct] = struct.Struct('@B')
ST_SBO_UI16: Final[struct.Struct] = struct.Struct('@H')
ST_SBO_UI32: Final[struct.Struct] = struct.Struct('@I')
ST_SBO_UI64: Final[struct.Struct] = struct.Struct('@Q')
# Struct for an RGB value that, in little endian, would be an int written in hex
# as 0x00BBGGRR.
ST_RGB: Final[struct.Struct] = struct.Struct('<BBBx')