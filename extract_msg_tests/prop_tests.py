__all__ = [
    'PropTests',
]


import datetime
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
# the value, otherwise it will be the values for length, real length, and
# reserved flags.
_propChecks = [
    (
        'Unspecified',
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
        FixedLengthProp,
        '02010001',
        0x0001,
        PropertyFlags.READABLE,
        None
    ),
    (
        'Int16',
        b'\x02\x00\x01\x02\x04\x00\x00\x00\x01\x23\x45\x67\x89\xAB\xCD\xEF',
        FixedLengthProp,
        '02010002',
        0x0002,
        PropertyFlags.WRITABLE,
        0x2301
    ),
    (
        'Int32',
        b'\x03\x00\x01\x02\x01\x00\x00\x00\x01\x23\x45\x67\x89\xAB\xCD\xEF',
        FixedLengthProp,
        '02010003',
        0x0003,
        PropertyFlags.MANDATORY,
        0x67452301
    ),
    (
        'Float',
        b'\x04\x00\x01\x02\x01\x00\x00\x00\x5B\xD3\xFC\x3D\x89\xAB\xCD\xEF',
        FixedLengthProp,
        '02010004',
        0x0004,
        PropertyFlags.MANDATORY,
        0.12345
    ),
    (
        'Double',
        b'\x05\x00\x01\x02\x01\x00\x00\x00\x7C\xF2\xB0\x50\x6B\x9A\xBF\x3F',
        FixedLengthProp,
        '02010005',
        0x0005,
        PropertyFlags.MANDATORY,
        0.12345
    ),
    (
        'Currency',
        b'\x06\x00\x01\x02\x01\x00\x00\x00\x00\x01\x00\x00\x00\x00\x00\x00',
        FixedLengthProp,
        '02010006',
        0x0006,
        PropertyFlags.MANDATORY,
        0.0256
    ),
    (
        'Floating Time',
        b'\x07\x00\x01\x02\x01\x00\x00\x00\x00\x00\x00\x00\x00\x60\x40\x40',
        FixedLengthProp,
        '02010007',
        0x0007,
        PropertyFlags.MANDATORY,
        PYTPFLOATINGTIME_START + datetime.timedelta(days = 32.75)
    ),
    (
        'Error Code',
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
        FixedLengthProp,
        '0201000B',
        0x000B,
        PropertyFlags.MANDATORY,
        False
    ),
    (
        'Boolean 0x0001',
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
        FixedLengthProp,
        '0201000B',
        0x000B,
        PropertyFlags.MANDATORY,
        True
    ),
    (
        'Boolean 0x10000',
        b'\x0B\x00\x01\x02\x01\x00\x00\x00\x00\x00\x01\x00\x00\x00\x00\x00',
        FixedLengthProp,
        '0201000B',
        0x000B,
        PropertyFlags.MANDATORY,
        False
    ),
    (
        'Int64',
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
        FixedLengthProp,
        '02010040',
        0x0040,
        PropertyFlags.MANDATORY,
        fromTimeStamp(0)
    ),
    (
        'Null Time 1',
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
        FixedLengthProp,
        '02010040',
        0x0040,
        PropertyFlags.MANDATORY,
        NULL_DATE
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
                self.assertIsInstance(prop, entry[2])
                self.assertEqual(prop.name, entry[3])
                self.assertEqual(prop.type, entry[4])
                self.assertIs(prop.flags, entry[5])

                # Check which class the prop is to know what checks to use.
                if isinstance(prop, FixedLengthProp):
                    if isinstance(prop.value, float):
                        self.assertAlmostEqual(prop.value, entry[6])
                    elif isinstance(prop.value, enum.Enum) or prop.type == 1:
                        self.assertIs(prop.value, entry[6])
                    else:
                        self.assertEqual(prop.value, entry[6])
                else:
                    prop = typing.cast(VariableLengthProp, prop)
                    self.assertEqual(prop.length, entry[6])
                    self.assertEqual(prop.realLength, entry[7])
                    self.assertEqual(prop.reservedFlags, entry[8])