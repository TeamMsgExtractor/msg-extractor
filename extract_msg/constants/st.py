"""
Struct constants.
"""

__all__ = [
    'ST1',
    'ST2',
    'STF32',
    'STF64',
    'STFIX',
    'STI16',
    'STI32',
    'STI64',
    'STI8',
    'STMF32',
    'STMF64',
    'STMI16',
    'STMI32',
    'STMI64',
    'STNP_ENT',
    'STNP_NAM',
    'STPEID',
    'STUI32',
    'STVAR',
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
    'ST_DATA_UI16',
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
    'ST_SYSTEMTIME',
    'ST_TZ',
]


import struct

from typing import Final


# Define pre-compiled structs to make unpacking slightly faster.
# General structs.
ST1 : Final[struct.Struct] = struct.Struct('<8x4I')
ST2 : Final[struct.Struct] = struct.Struct('<H2xI8x')
# Struct used for unpacking a system time.
ST_SYSTEMTIME : Final[struct.Struct] = struct.Struct('<8H')
# Struct used for unpacking a GUID from bytes.
ST_GUID : Final[struct.Struct] = struct.Struct('<IHH8s')
# Struct for unpacking a TimeZoneStruct from bytes.
ST_TZ : Final[struct.Struct] = struct.Struct('<iiiH16sH16s')
# Struct for packing a compount file directory entry.
ST_CF_DIR_ENTRY : Final[struct.Struct] = struct.Struct('<64sHBBIII16sIQQIQ')
# Structs used by data.py
ST_DATA_UI32 : Final[struct.Struct] = struct.Struct('<I')
ST_DATA_UI16 : Final[struct.Struct] = struct.Struct('<H')
# Structs used by named.py
STNP_NAM : Final[struct.Struct] = struct.Struct('<i')
# Struct used for unpacking the entries in the entry stream
STNP_ENT : Final[struct.Struct] = struct.Struct('<IHH')
# Structs used by prop.py
STFIX : Final[struct.Struct] = struct.Struct('<8x8s')
STVAR : Final[struct.Struct] = struct.Struct('<8xi4s')
# Structs to help with email type to python type conversions
STI8 : Final[struct.Struct] = struct.Struct('<b');
STI16 : Final[struct.Struct] = struct.Struct('<h6x')
STI32 : Final[struct.Struct] = struct.Struct('<I4x')
STI64 : Final[struct.Struct] = struct.Struct('<q')
STF32 : Final[struct.Struct] = struct.Struct('<f4x')
STF64 : Final[struct.Struct] = struct.Struct('<d')
STUI32 : Final[struct.Struct] = struct.Struct('<I4x')
STMI16 : Final[struct.Struct] = struct.Struct('<h')
STMI32 : Final[struct.Struct] = struct.Struct('<i')
STMI64 : Final[struct.Struct] = struct.Struct('<q')
STMF32 : Final[struct.Struct] = struct.Struct('<f')
STMF64 : Final[struct.Struct] = struct.Struct('<d')
# PermanentEntryID parsing struct
STPEID : Final[struct.Struct] = struct.Struct('<B3x16s4xI')
# Struct for unpacking the first part of the BusinessCardDisplayDefinition
# structure.
ST_BC_HEAD : Final[struct.Struct] = struct.Struct('BBBBBBBBIB')
# Struct for completely unpacking the FieldInfo structure.
ST_BC_FIELD_INFO : Final[struct.Struct] = struct.Struct('HBBBxHII')
# Structs for reading from a BytesReader. Some are just aliases for existing
# structs, used for clarity and consistency in the code.
ST_LE_I8 : Final[struct.Struct] = STI8
ST_LE_I16 : Final[struct.Struct] = STMI16
ST_LE_I32 : Final[struct.Struct] = STMI32
ST_LE_I64 : Final[struct.Struct] = STMI64
ST_LE_UI8 : Final[struct.Struct] = struct.Struct('<B')
ST_LE_UI16 : Final[struct.Struct] = struct.Struct('<H')
ST_LE_UI32 : Final[struct.Struct] = struct.Struct('<I')
ST_LE_UI64 : Final[struct.Struct] = struct.Struct('<Q')
ST_LE_F32 : Final[struct.Struct] = STF32
ST_LE_F64 : Final[struct.Struct] = STF64
ST_BE_I8 : Final[struct.Struct] = struct.Struct('>b')
ST_BE_I16 : Final[struct.Struct] = struct.Struct('>h')
ST_BE_I32 : Final[struct.Struct] = struct.Struct('>i')
ST_BE_I64 : Final[struct.Struct] = struct.Struct('>q')
ST_BE_UI8 : Final[struct.Struct] = struct.Struct('>B')
ST_BE_UI16 : Final[struct.Struct] = struct.Struct('>H')
ST_BE_UI32 : Final[struct.Struct] = struct.Struct('>I')
ST_BE_UI64 : Final[struct.Struct] = struct.Struct('>Q')
ST_BE_F32: Final[struct.Struct]  = struct.Struct('>f')
ST_BE_F64 : Final[struct.Struct] = struct.Struct('>d')