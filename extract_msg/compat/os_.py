"""
Compatibility module to ensure that certain functions exist across python versions
"""

from os import *
import sys

if sys.version_info[0] >= 3:
    getcwdu = getcwd
