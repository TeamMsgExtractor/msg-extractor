# -*- coding: utf-8 -*-

"""
extract_msg.exceptions
~~~~~~~~~~~~~~~~~~~
This module contains the set of extract_msg exceptions.
"""

__all__ = [
    'BadHtmlError',
    'ConversionError',
    'DataNotFoundError',
    'DeencapMalformedData',
    'DeencapNotEncapsulated',
    'ExecutableNotFound',
    'IncompatibleOptionsError',
    'InvalidFileFormatError',
    'InvaildPropertyIdError',
    'InvalidVersionError',
    'StandardViolationError',
    'TZError',
    'UnknownCodepageError',
    'UnsupportedEncodingError',
    'UnknownTypeError',
    'UnsupportedMSGTypeError',
    'UnrecognizedMSGTypeError',
    'WKError',
]


import logging


# Add logger bus.
logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class BadHtmlError(ValueError):
    """
    HTML failed to pass validation.
    """

class ConversionError(Exception):
    """
    An error occured during type conversion.
    """

class DataNotFoundError(Exception):
    """
    Requested stream type was unavailable.
    """

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

class IncompatibleOptionsError(Exception):
    """
    Provided options are incompatible with each other.
    """

class InvalidFileFormatError(OSError):
    """
    An Invalid File Format Error occurred.
    """

class InvaildPropertyIdError(Exception):
    """
    The provided property ID was invalid.
    """

class InvalidVersionError(Exception):
    """
    The version specified is invalid.
    """

class StandardViolationError(InvalidFileFormatError):
    """
    A critical violation of the MSG standards was detected and could not be
    recovered from. Recoverable violations will result in log messages instead.

    Any that could reasonably be skipped, although are likely to still cause 
    errors down the line, can be suppressed.
    """

class TZError(Exception):
    """
    Specifically not an OSError to avoid being caught by parts of the module.
    This error represents a fatal error in the datetime parsing as it usually
    means your installation of tzlocal or tzdata are broken. If you have
    received this error after using PyInstaller, you must include the resource
    files for tzdata for it to work properly. See TeamMsgExtractor#272 and
    TeamMsgExtractor#169 for information on why you are getting this error.
    """

class UnknownCodepageError(Exception):
    """
    The codepage provided was not one we know of.
    """

class UnsupportedEncodingError(NotImplementedError):
    """
    The codepage provided is known but is not supported.
    """

class UnknownTypeError(Exception):
    """
    The type specified is not one that is recognized.
    """

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

class WKError(RuntimeError):
    """
    An error occured while running wkhtmltopdf.
    """
