# -*- coding: utf-8 -*-

import logging

"""
extract_msg.exceptions
~~~~~~~~~~~~~~~~~~~
This module contains the set of extract_msg exceptions.
"""

# Add logger bus.
logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class BadHtmlError(ValueError):
    """
    HTML failed to pass validation.
    """
    pass

class ConversionError(Exception):
    """
    An error occured during type conversion.
    """
    pass

class DataNotFoundError(Exception):
    """
    Requested stream type was unavailable.
    """
    pass

class DeencapMalformedData(Exception):
    """
    Data to deencapsulate was malformed in some way.
    """

class DeencapNotEncapsulated(Exception):
    """
    Data to deencapsulate did not contain any encapsulated data.
    """

class ExecutableNotFound(Exception):
    """
    Could not find the specified executable.
    """
    pass

class IncompatibleOptionsError(Exception):
    """
    Provided options are incompatible with each other.
    """

class InvalidFileFormatError(OSError):
    """
    An Invalid File Format Error occurred.
    """
    pass

class InvaildPropertyIdError(Exception):
    """
    The provided property ID was invalid.
    """
    pass

class InvalidVersionError(Exception):
    """
    The version specified is invalid.
    """
    pass

class StandardViolationError(Exception):
    """
    A critical violation of the MSG standards was detected and could not be
    recovered from. Recoverable violations will result in log messages instead.
    """
    pass

class UnknownCodepageError(Exception):
    """
    The codepage provided was not one we know of.
    """
    pass

class UnknownTypeError(Exception):
    """
    The type specified is not one that is recognized.
    """
    pass

class UnsupportedMSGTypeError(NotImplementedError):
    """
    An exception that is raised when an MSG class is recognized by not
    supported.
    """

class UnrecognizedMSGTypeError(TypeError):
    """
    An exception that is raised when the module cannot determine how to properly
    open a specific class of msg file.
    """
    pass

class WKError(RuntimeError):
    """
    An error occured while running wkhtmltopdf.
    """
    pass
