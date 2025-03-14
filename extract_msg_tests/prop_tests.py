__all__ = [
    'PropTests',
]


import datetime
import decimal
import enum
import typing
import unittest

from extract_msg.constants import (
        FIXED_LENGTH_PROPS_STRING, NULL_DATE, PYTPFLOATINGTIME_START,
        VARIABLE_LENGTH_PROPS_STRING
    )
from extract_msg.enums import ErrorCodeType, PropertyFlags
from extract_msg.properties.prop import (
        createNewProp, createProp, FixedLengthProp, VariableLengthProp
    )
from extract_msg.utils import fromTimeStamp


# List of tuples that contain a property string with data to check for. This
# allows less lines needed for all the tests. The items are, in this order:
# name, source bytes, output class, output name, output type, output flags,
# extras. If  the class is FixedLengthProp, extras will be a single item for
# the value, otherwise it will be the values for size and reserved flags.
_propChecks = [
    # Fixed Length Props.
    (
        'Unspecified',
        b'\x00\x00\x01\x02\x01\x00\x00\x00\x01\x23\x45\x67\x89\xAB\xCD\xEF',
        b'\x00\x00\x01\x02\x01\x00\x00\x00\x01\x23\x45\x67\x89\xAB\xCD\xEF',
        FixedLengthProp,
        '02010000',
        0x0000,
        PropertyFlags.MANDATORY,
        b'\x01\x23\x45\x67\x89\xAB\xCD\xEF'
    ),
    (
        'Null',
        b'\x01\x00\x01\x02\x02\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
        b'\x01\x00\x01\x02\x02\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
        FixedLengthProp,
        '02010001',
        0x0001,
        PropertyFlags.READABLE,
        None
    ),
    (
        'Int16',
        b'\x02\x00\x01\x02\x04\x00\x00\x00\x01\x23\x45\x67\x89\xAB\xCD\xEF',
        b'\x02\x00\x01\x02\x04\x00\x00\x00\x01\x23\x00\x00\x00\x00\x00\x00',
        FixedLengthProp,
        '02010002',
        0x0002,
        PropertyFlags.WRITABLE,
        0x2301
    ),
    (
        'Int32',
        b'\x03\x00\x01\x02\x01\x00\x00\x00\x01\x23\x45\x67\x89\xAB\xCD\xEF',
        b'\x03\x00\x01\x02\x01\x00\x00\x00\x01\x23\x45\x67\x00\x00\x00\x00',
        FixedLengthProp,
        '02010003',
        0x0003,
        PropertyFlags.MANDATORY,
        0x67452301
    ),
    (
        'Float',
        b'\x04\x00\x01\x02\x01\x00\x00\x00\x5B\xD3\xFC\x3D\x89\xAB\xCD\xEF',
        None, # Float won't necessarily end up exactly the same.
        FixedLengthProp,
        '02010004',
        0x0004,
        PropertyFlags.MANDATORY,
        0.12345
    ),
    (
        'Double',
        b'\x05\x00\x01\x02\x01\x00\x00\x00\x7C\xF2\xB0\x50\x6B\x9A\xBF\x3F',
        None, # Double won't necessarily end up exactly the same.
        FixedLengthProp,
        '02010005',
        0x0005,
        PropertyFlags.MANDATORY,
        0.12345
    ),
    (
        'Currency',
        b'\x06\x00\x01\x02\x01\x00\x00\x00\x00\x01\x00\x00\x00\x00\x00\x00',
        b'\x06\x00\x01\x02\x01\x00\x00\x00\x00\x01\x00\x00\x00\x00\x00\x00',
        FixedLengthProp,
        '02010006',
        0x0006,
        PropertyFlags.MANDATORY,
        decimal.Decimal(0.0256)
    ),
    (
        'Floating Time',
        b'\x07\x00\x01\x02\x01\x00\x00\x00\x00\x00\x00\x00\x00\x60\x40\x40',
        None, # Floating time won't necessarily end up the same.
        FixedLengthProp,
        '02010007',
        0x0007,
        PropertyFlags.MANDATORY,
        PYTPFLOATINGTIME_START + datetime.timedelta(days = 32.75)
    ),
    (
        'Error Code',
        b'\x0A\x00\x01\x02\x01\x00\x00\x00\xD7\x04\x00\x00\x00\x00\x00\x00',
        b'\x0A\x00\x01\x02\x01\x00\x00\x00\xD7\x04\x00\x00\x00\x00\x00\x00',
        FixedLengthProp,
        '0201000A',
        0x000A,
        PropertyFlags.MANDATORY,
        ErrorCodeType.DELETE_MESSAGE
    ),
    (
        'Boolean 0x0000',
        b'\x0B\x00\x01\x02\x01\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
        b'\x0B\x00\x01\x02\x01\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
        FixedLengthProp,
        '0201000B',
        0x000B,
        PropertyFlags.MANDATORY,
        False
    ),
    (
        'Boolean 0x0001',
        b'\x0B\x00\x01\x02\x01\x00\x00\x00\x01\x00\x00\x00\x00\x00\x00\x00',
        b'\x0B\x00\x01\x02\x01\x00\x00\x00\x01\x00\x00\x00\x00\x00\x00\x00',
        FixedLengthProp,
        '0201000B',
        0x000B,
        PropertyFlags.MANDATORY,
        True
    ),
    (
        'Boolean 0x0002',
        b'\x0B\x00\x01\x02\x01\x00\x00\x00\x02\x00\x00\x00\x00\x00\x00\x00',
        b'\x0B\x00\x01\x02\x01\x00\x00\x00\x01\x00\x00\x00\x00\x00\x00\x00',
        FixedLengthProp,
        '0201000B',
        0x000B,
        PropertyFlags.MANDATORY,
        True
    ),
    (
        'Boolean 0x10000',
        b'\x0B\x00\x01\x02\x01\x00\x00\x00\x00\x00\x01\x00\x00\x00\x00\x00',
        b'\x0B\x00\x01\x02\x01\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
        FixedLengthProp,
        '0201000B',
        0x000B,
        PropertyFlags.MANDATORY,
        False
    ),
    (
        'Int64',
        b'\x14\x00\x01\x02\x01\x00\x00\x00\x01\x02\x03\x04\x05\x06\x07\x08',
        b'\x14\x00\x01\x02\x01\x00\x00\x00\x01\x02\x03\x04\x05\x06\x07\x08',
        FixedLengthProp,
        '02010014',
        0x0014,
        PropertyFlags.MANDATORY,
        0x807060504030201
    ),
    (
        'Time',
        b'\x40\x00\x01\x02\x01\x00\x00\x00\x00\x80\x3E\xD5\xDE\xB1\x9D\x01',
        b'\x40\x00\x01\x02\x01\x00\x00\x00\x00\x80\x3E\xD5\xDE\xB1\x9D\x01',
        FixedLengthProp,
        '02010040',
        0x0040,
        PropertyFlags.MANDATORY,
        fromTimeStamp(0)
    ),
    (
        'Null Time 1',
        b'\x40\x00\x01\x02\x01\x00\x00\x00\x00\x52\x13\xF6\xB3\xE5\xB2\x0C',
        b'\x40\x00\x01\x02\x01\x00\x00\x00\x00\x52\x13\xF6\xB3\xE5\xB2\x0C',
        FixedLengthProp,
        '02010040',
        0x0040,
        PropertyFlags.MANDATORY,
        NULL_DATE
    ),
    (
        'Null Time 2',
        b'\x40\x00\x01\x02\x01\x00\x00\x00\x00\x40\xDD\xA3\x57\x45\xB3\x0C',
        b'\x40\x00\x01\x02\x01\x00\x00\x00\x00\x40\xDD\xA3\x57\x45\xB3\x0C',
        FixedLengthProp,
        '02010040',
        0x0040,
        PropertyFlags.MANDATORY,
        NULL_DATE
    ),
    (
        'Null Time 3',
        b'\x40\x00\x01\x02\x01\x00\x00\x00\x00\x3F\xDD\xA3\x57\x45\xB3\x0C',
        b'\x40\x00\x01\x02\x01\x00\x00\x00\x00\x3F\xDD\xA3\x57\x45\xB3\x0C',
        FixedLengthProp,
        '02010040',
        0x0040,
        PropertyFlags.MANDATORY,
        NULL_DATE
    ),
    (
        'Null Time 4',
        b'\x40\x00\x1C\x30\x06\x00\x00\x00\xff\xff\xff\xff\xff\xff\xff\x7f',
        b'\x40\x00\x1C\x30\x06\x00\x00\x00\xff\xff\xff\xff\xff\xff\xff\x7f',
        FixedLengthProp,
        '301C0040',
        0x0040,
        PropertyFlags.READABLE | PropertyFlags.WRITABLE,
        NULL_DATE
    ),
    # Variable Length Props.
    (
        'Object',
        b'\x0D\x00\x01\x02\x01\x00\x00\x00\xFF\xFF\xFF\xFF\x01\x23\x45\x67',
        b'\x0D\x00\x01\x02\x01\x00\x00\x00\xFF\xFF\xFF\xFF\x01\x23\x45\x67',
        VariableLengthProp,
        '0201000D',
        0x000D,
        PropertyFlags.MANDATORY,
        0xFFFFFFFF,
        0x67452301
    ),
    (
        'Object Warning',
        b'\x0D\x00\x01\x02\x01\x00\x00\x00\x00\x00\x00\x00\x01\x23\x45\x67',
        b'\x0D\x00\x01\x02\x01\x00\x00\x00\xFF\xFF\xFF\xFF\x01\x23\x45\x67',
        VariableLengthProp,
        '0201000D',
        0x000D,
        PropertyFlags.MANDATORY,
        0xFFFFFFFF,
        0x67452301
    ),
    (
        'String8',
        b'\x1E\x00\x01\x02\x01\x00\x00\x00\x08\x00\x00\x00\x01\x23\x45\x67',
        b'\x1E\x00\x01\x02\x01\x00\x00\x00\x08\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '0201001E',
        0x001E,
        PropertyFlags.MANDATORY,
        7,
        0x67452301
    ),
    (
        'String',
        b'\x1F\x00\x01\x02\x01\x00\x00\x00\x08\x00\x00\x00\x01\x23\x45\x67',
        b'\x1F\x00\x01\x02\x01\x00\x00\x00\x08\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '0201001F',
        0x001F,
        PropertyFlags.MANDATORY,
        3,
        0x67452301
    ),
    (
        'GUID',
        b'\x48\x00\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        b'\x48\x00\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '02010048',
        0x0048,
        PropertyFlags.MANDATORY,
        16,
        0x67452301
    ),
    (
        'GUID Warning',
        b'\x48\x00\x01\x02\x01\x00\x00\x00\x08\x00\x00\x00\x01\x23\x45\x67',
        b'\x48\x00\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '02010048',
        0x0048,
        PropertyFlags.MANDATORY,
        16,
        0x67452301
    ),
    (
        'Multiple Int16',
        b'\x02\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        b'\x02\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '02011002',
        0x1002,
        PropertyFlags.MANDATORY,
        8,
        0x67452301
    ),
    (
        'Multiple Int32',
        b'\x03\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        b'\x03\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '02011003',
        0x1003,
        PropertyFlags.MANDATORY,
        4,
        0x67452301
    ),
    (
        'Multiple Float',
        b'\x04\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        b'\x04\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '02011004',
        0x1004,
        PropertyFlags.MANDATORY,
        4,
        0x67452301
    ),
    (
        'Multiple Double',
        b'\x05\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        b'\x05\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '02011005',
        0x1005,
        PropertyFlags.MANDATORY,
        2,
        0x67452301
    ),
    (
        'Multiple Currency',
        b'\x06\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        b'\x06\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '02011006',
        0x1006,
        PropertyFlags.MANDATORY,
        2,
        0x67452301
    ),
    (
        'Multiple Floating Time',
        b'\x07\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        b'\x07\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '02011007',
        0x1007,
        PropertyFlags.MANDATORY,
        2,
        0x67452301
    ),
    (
        'Multiple Int64',
        b'\x14\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        b'\x14\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '02011014',
        0x1014,
        PropertyFlags.MANDATORY,
        2,
        0x67452301
    ),
    (
        'Multiple String8',
        b'\x1E\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        b'\x1E\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '0201101E',
        0x101E,
        PropertyFlags.MANDATORY,
        4,
        0x67452301
    ),
    (
        'Multiple String',
        b'\x1F\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        b'\x1F\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '0201101F',
        0x101F,
        PropertyFlags.MANDATORY,
        4,
        0x67452301
    ),
    (
        'Multiple Time',
        b'\x40\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        b'\x40\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '02011040',
        0x1040,
        PropertyFlags.MANDATORY,
        2,
        0x67452301
    ),
    (
        'Multiple GUID',
        b'\x48\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        b'\x48\x10\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '02011048',
        0x1048,
        PropertyFlags.MANDATORY,
        1,
        0x67452301
    ),
    (
        'Multiple Binary',
        b'\x02\x11\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        b'\x02\x11\x01\x02\x01\x00\x00\x00\x10\x00\x00\x00\x01\x23\x45\x67',
        VariableLengthProp,
        '02011102',
        0x1102,
        PropertyFlags.MANDATORY,
        2,
        0x67452301
    ),
]


class PropTests(unittest.TestCase):
    def testCreateNewProps(self):
        for type_ in FIXED_LENGTH_PROPS_STRING + VARIABLE_LENGTH_PROPS_STRING:
            with self.subTest(f'Create New Prop {type_}.'):
                createNewProp('0100' + type_)

    def testCreateProps(self):
        for entry in _propChecks:
            with self.subTest(f'Prop Test {entry[0]}.'):
                prop = createProp(entry[1])

                # Common assertions.
                self.assertIsInstance(prop, entry[3])
                self.assertEqual(prop.name, entry[4])
                self.assertEqual(prop.type, entry[5])
                self.assertIs(prop.flags, entry[6])

                # Check which class the prop is to know what checks to use.
                if isinstance(prop, FixedLengthProp):
                    if isinstance(prop.value, (float, decimal.Decimal)):
                        self.assertAlmostEqual(prop.value, entry[7])
                    elif isinstance(prop.value, enum.Enum) or prop.type == 1:
                        self.assertIs(prop.value, entry[7])
                    else:
                        self.assertEqual(prop.value, entry[7])
                else:
                    prop = typing.cast(VariableLengthProp, prop)
                    self.assertEqual(prop.size, entry[7])
                    self.assertEqual(prop.reservedFlags, entry[8])

                # Ensure the output value is as expected if `entry[2]` is not
                # None.
                if entry[2]:
                    self.assertEqual(bytes(prop), entry[2])