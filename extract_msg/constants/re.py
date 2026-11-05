"""
Regular expression constants.
"""


__all__ = [
    'CSS_IMPORT_EXTERNAL',
    'CSS_URL',
    'EXTERNAL_URL_IGNORED_CHARS',
    'HTML_BODY_START',
    'HTML_SAN_SPACE',
    'INVALID_FILENAME_CHARS',
    'INVALID_OLE_PATH',
    'RTF_BODY_STRIP_INIT',
    'RTF_BODY_STRIP_PRE_CLOSE',
    'RTF_BODY_STRIP_PRE_OPEN',
    'RTF_ENC_BODY_START',
]


import re

from typing import Final


# Allow better typing in versions above 3.8.
import sys
if sys.version_info >= (3, 9):
    _RE_STR_TYPE = re.Pattern[str]
    _RE_BYTES_TYPE = re.Pattern[bytes]
else:
    _RE_STR_TYPE = re.Pattern
    _RE_BYTES_TYPE = re.Pattern


# Characters that are invalid in a filename.
INVALID_FILENAME_CHARS: Final[_RE_STR_TYPE] = re.compile(r'[\\/:*?"<>|]')
# Regular expression to find sections of spaces for htmlSanitize.
HTML_SAN_SPACE: Final[_RE_STR_TYPE] = re.compile('  +')
# Regular expression to find the start of the html body.
HTML_BODY_START: Final[_RE_BYTES_TYPE] = re.compile(b'<body[^>]*>')
# Regular expression to find the start of the html body in encapsulated RTF.
# This is used for one of the pattern types that makes life easy.
RTF_ENC_BODY_START: Final[_RE_BYTES_TYPE] = re.compile(br'\{\\\*\\htmltag[0-9]* ?<body[^>]*>\}')
# Used in the vaildation of OLE paths. Any of these characters in a name make it
# invalid.
INVALID_OLE_PATH: Final[_RE_STR_TYPE] = re.compile(r'[:/\\!]')

# Used as the initial step in stripping RTF files for deencapsulation. Finds
# ignored sections that do not contrain groups *and* finds HTML tag sections
# that are entirely empty. It also then finds sections of data that can be
# merged together without affecting the results
RTF_BODY_STRIP_INIT: Final[_RE_BYTES_TYPE] = re.compile(rb'(\\htmlrtf[^0{}][^{}]*?\\htmlrtf0 ?)|(\{\\\*\\htmltag[0-9]+\})|(\\htmlrtf0 ?\\htmlrtf1? ?)|(\\htmlrtf1? ?\{\}\\htmlrtf0 ?)|(\\htmlrtf1? ?\\\'[a-fA-F0-9]{2}\\htmlrtf0 ?)')

# Preprocessing steps to simplify the RTF.
RTF_BODY_STRIP_PRE_CLOSE: Final[_RE_BYTES_TYPE] = re.compile(rb'(\\htmlrtf1? ?}\\htmlrtf0 ?)|(\\htmlrtf1? ?[^0{}][^{}]*?} ?\\htmlrtf0 ?)')
RTF_BODY_STRIP_PRE_OPEN: Final[_RE_BYTES_TYPE] = re.compile(rb'\\htmlrtf1? ?{[^{}]*?\\htmlrtf0 ?')

# Used for stripping urls from CSS.
CSS_URL: Final[_RE_STR_TYPE] = re.compile(r'url\(\s*["\']?(https?:\/\/[^"\')\s]+)["\']?\s*\)', re.I)
# Used for stripping external @import rules from CSS. An @import that pulls in
# an external http(s) resource causes wkhtmltopdf to fetch it, so the whole
# at-rule is removed. This matches both the string form (@import "http://...";)
# and the url() form (@import url(http://...);), including any media-query list,
# up to the terminating semicolon. CSS_URL alone does not catch the string form.
CSS_IMPORT_EXTERNAL: Final[_RE_STR_TYPE] = re.compile(r'@import\b[^;]*?https?:\/\/[^;]*;?', re.I)
# Characters that HTML URL parsing ignores: leading C0 controls and spaces, and
# interior tab/CR/LF. These are removed before testing an attribute's scheme so
# a value like " http://..." or "ht\ttp://..." cannot smuggle a URL past the
# external-file check.
EXTERNAL_URL_IGNORED_CHARS: Final[_RE_STR_TYPE] = re.compile(r'^[\x00-\x20]+|[\t\r\n]')
