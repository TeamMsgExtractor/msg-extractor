"""
Debug variable used throughout the entire module.
Turns on debugging information.
"""

import logging



logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

_debug = False

def toggle_debug():
	global _debug
	_debug = bool(_debug ^ 1)
