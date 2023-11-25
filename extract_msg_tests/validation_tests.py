__all__ = [
    'ValidationTests',
]


import enum
import re
import struct
import unittest


class ValidationTests(unittest.TestCase):
    """
    Tests to perform basic validation on parts of the module.
    """

    def testConstants(self):
        """
        Tests to validate the constants submodule.
        """
        from extract_msg import constants

        # Validate the structs section.
        # Check for exports that are not structs.
        for name in constants.st.__all__:
            with self.subTest(f'Check that {name} is an instance of Struct.'):
                self.assertIsInstance(getattr(constants.st, name), struct.Struct)

        # Check for structs that didn't get exported.
        for name in dir(constants.st):
            if isinstance(getattr(constants.st, name), struct.Struct):
                with self.subTest(f'Ensure {name} is exported.'):
                    self.assertIn(name, constants.st.__all__)

    def testEnums(self):
        """
        Tests to validate the enums submodule.
        """
        # First test, make sure everything in enums is actually an enum. Only
        # test the actual exports.
        from extract_msg import enums
        for name in enums.__all__:
            with self.subTest(f'Check that {name} is a subclass of Enum.'):
                self.assertTrue(issubclass(getattr(enums, name), enum.Enum))

        # Check for enums that didn't get exported.
        for name in dir(enums):
            if isinstance(getattr(enums, name), enum.Enum):
                with self.subTest(f'Ensure {name} is exported.'):
                    self.assertIn(name, enums.__all__)