"""
The tests for various parts of extract_msg. Tests can be run on user files by
defining the EXTRACT_MSG_TEST_DIR environment variable and setting it to the
folder with your own tests files. It must match the folder structure for the
specific test for that test to run.
"""
import unittest

from extract_msg_tests import *


unittest.main(verbosity=2)
