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


class InvalidFileFormat(OSError):
    """
    A Invalid File Format Error occurred
    """
    pass
