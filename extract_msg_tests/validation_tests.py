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

        # Ensure names are all uppercase.
        for name in constants.st.__all__:
            with self.subTest(f'Ensure {name} is a valid constant name.'):
                self.assertTrue(name.isupper())

        # Validate the regular expressions section.
        # Check for exports that are not regular expressions.
        for name in constants.re.__all__:
            with self.subTest(f'Check that {name} is an instance of Pattern.'):
                self.assertIsInstance(getattr(constants.re, name), re.Pattern)

        # Check for regular expressions that didn't get exported.
        for name in dir(constants.re):
            if isinstance(getattr(constants.re, name), re.Pattern):
                with self.subTest(f'Ensure {name} is exported.'):
                    self.assertIn(name, constants.re.__all__)

        # Ensure names are all uppercase.
        for name in constants.re.__all__:
            with self.subTest(f'Ensure {name} is a valid constant name.'):
                self.assertTrue(name.isupper())

        # The PropertiesSet section.
        # Ensure names are all uppercase.
        for name in constants.ps.__all__:
            with self.subTest(f'Ensure {name} is a valid constant name.'):
                self.assertTrue(name.isupper())

        # Basic validation of the normal constants section. We have to do
        # things a little differently since the submodule are lowercase.
        # Ensure names are all uppercase.
        for name in constants.__all__:
            with self.subTest(f'Ensure {name} is a valid constant name.'):
                if not isinstance(getattr(constants, name), type(constants)):
                    self.assertTrue(name.isupper())

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