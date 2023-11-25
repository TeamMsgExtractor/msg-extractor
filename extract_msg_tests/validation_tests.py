__all__ = [
    'ValidationTests',
]


import enum
import unittest


class ValidationTests(unittest.TestCase):
    """
    Tests to perform basic validation on parts of the module.
    """

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