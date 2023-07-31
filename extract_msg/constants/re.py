"""
Regular expression constants.
"""

__all__ = [
    'HTML_BODY_START',
    'HTML_SAN_SPACE',
    'INVALID_FILENAME_CHARS',
    'INVALID_OLE_PATH',
    'RTF_ENC_BODY_START',
]


import re


# Characters that are invalid in a filename.
INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')
# Regular expression to find sections of spaces for htmlSanitize.
HTML_SAN_SPACE = re.compile('  +')
# Regular expression to find the start of the html body.
HTML_BODY_START = re.compile(b'<body[^>]*>')
# Regular expression to find the start of the html body in encapsulated RTF.
# This is used for one of the pattern types that makes life easy.
RTF_ENC_BODY_START = re.compile(br'\{\\\*\\htmltag[0-9]* ?<body[^>]*>\}')
# Used in the vaildation of OLE paths. Any of these characters in a name make it
# invalid.
INVALID_OLE_PATH = re.compile(r'[:/\\!]')

