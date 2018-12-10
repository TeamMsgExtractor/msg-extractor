# -*- coding: utf-8 -*-

import logging

"""
requests.exceptions
~~~~~~~~~~~~~~~~~~~
This module contains the set of msg_extractor exceptions.
"""

# add logger bus
logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class InvalidFileFormat(OSError):
    """A Invalid File Format Error occurred"""
    pass
