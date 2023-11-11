__all__ = [
    'PropTests',
]


import enum
import typing
import unittest

from extract_msg.enums import PropertyFlags
from extract_msg.properties.prop import (
        createProp, FixedLengthProp, VariableLengthProp
    )


# List of tuples that contain a property string with data to check for. This
# allows less lines needed for all the tests. The items are, in this order:
# source bytes, output class, output name, output type, output flags, extras. If
# the class is FixedLengthProp, extras will be a single item for the value,
# otherwise it will be the values for length, real length, and reserved flags.
_propChecks = [
    (
        b'\x00\x00\x01\x02\x01\x00\x00\x00\x01\x23\x45\x67\x89\xAB\xCD\xEF', FixedLengthProp,
        '02010000',
        0,
        PropertyFlags.MANDATORY,
        b'\x01\x23\x45\x67\x89\xAB\xCD\xEF'
    ),
    (
        b'\x01\x00\x01\x02\x02\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00', FixedLengthProp,
        '02010001',
        1,
        PropertyFlags.READABLE,
        None
    ),
    (
        b'\x02\x00\x01\x02\x04\x00\x00\x00\x01\x23\x45\x67\x89\xAB\xCD\xEF', FixedLengthProp,
        '02010002',
        2,
        PropertyFlags.WRITABLE,
        0x2301
    ),
    (
        b'\x03\x00\x01\x02\x01\x00\x00\x00\x01\x23\x45\x67\x89\xAB\xCD\xEF', FixedLengthProp,
        '02010003',
        3,
        PropertyFlags.MANDATORY,
        0x67452301
    ),
]


class PropTests(unittest.TestCase):
    def testCreateProps(self):
        for index, entry in enumerate(_propChecks):
            errMsg = f'Prop Test {index}'
            prop = createProp(entry[0])

            # Common assertions.
            self.assertIsInstance(prop, entry[1], errMsg)
            self.assertEqual(prop.name, entry[2], errMsg)
            self.assertEqual(prop.type, entry[3], errMsg)
            self.assertIs(prop.flags, entry[4], errMsg)

            # Check which class the prop is to know what checks to use.
            if isinstance(prop, FixedLengthProp):
                if isinstance(prop.value, float):
                    self.assertAlmostEqual(prop.value, entry[5], errMsg)
                elif isinstance(prop.value, enum.Enum) or prop.type == 1:
                    self.assertIs(prop.value, entry[5], errMsg)
                else:
                    self.assertEqual(prop.value, entry[5], errMsg)
            else:
                prop = typing.cast(VariableLengthProp, prop)
                self.assertEqual(prop.length, entry[5], errMsg)
                self.assertEqual(prop.realLength, entry[6], errMsg)
                self.assertEqual(prop.reservedFlags, entry[7], errMsg)