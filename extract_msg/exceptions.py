# -*- coding: utf-8 -*-

import logging

"""
extract_msg.exceptions
~~~~~~~~~~~~~~~~~~~
This module contains the set of extract_msg exceptions.
"""

# add logger bus
logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class InvalidFileFormatError(OSError):
    """
    An Invalid File Format Error occurred.
    """
    pass

class MissingEncodingError(Exception):
    """
    MSG file is missing an encodng.
    """
    pass

class UnrecognizedMSGTypeError(TypeError):
    """
    An exception that is raised when the module cannot determine how to properly
    open a specific class of msg file.
    """
    pass
