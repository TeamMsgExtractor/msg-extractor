"""
File for handling specialized encoding tasks or information.
"""

__all__ = [
    'lookupCodePage',
]


# This adds additional encodings to python.
import ebcdic as _
import codecs

from ..exceptions import UnknownCodepageError, UnsupportedEncodingError


# This is a dictionary matching the code page number to it's encoding name.
# The list used to make this can be found here:
# https://docs.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
### TODO:
# Many of these code pages are not supported by Python. As such, we should
# really implement them ourselves to make sure that if someone wants to use an
# MSG file with one of those encodings, they are able to. Perhaps we should
# create a seperate module for that?
# Code pages that currently don't have a supported encoding will be preceded by
# `# UNSUPPORTED`.
# For some of these, it is also possible that the name we are trying to find
# them with is not known to Python. I have already confirmed this for a few of
# them, and adjusted their names to ones that python would recognize. It is
# Possible I missed a few.
_CODE_PAGES = {
    37: 'IBM037', # IBM EBCDIC US-Canada
    437: 'IBM437', # OEM United States
    500: 'IBM500', # IBM EBCDIC International
    708: 'ASMO-708', # Arabic (ASMO 708)
    # UNSUPPORTED.
    709: '', # Arabic (ASMO-449+, BCON V4)
    # UNSUPPORTED.
    710: '', # Arabic - Transparent Arabic
    # UNSUPPORTED.
    720: 'DOS-720', # Arabic (Transparent ASMO); Arabic (DOS)
    737: 'cp737', # OEM Greek (formerly 437G); Greek (DOS)
    775: 'ibm775', # OEM Baltic; Baltic (DOS)
    850: 'ibm850', # OEM Multilingual Latin 1; Western European (DOS)
    852: 'ibm852', # OEM Latin 2; Central European (DOS)
    855: 'IBM855', # OEM Cyrillic (primarily Russian)
    857: 'ibm857', # OEM Turkish; Turkish (DOS)
    858: 'cp858', # OEM Multilingual Latin 1 + Euro symbol
    860: 'IBM860', # OEM Portuguese; Portuguese (DOS)
    861: 'ibm861', # OEM Icelandic; Icelandic (DOS)
    862: 'cp862', # OEM Hebrew; Hebrew (DOS)
    863: 'IBM863', # OEM French Canadian; French Canadian (DOS)
    864: 'IBM864', # OEM Arabic; Arabic (864)
    865: 'IBM865', # OEM Nordic; Nordic (DOS)
    866: 'cp866', # OEM Russian; Cyrillic (DOS)
    869: 'ibm869', # OEM Modern Greek; Greek, Modern (DOS)
    870: 'cp870', # IBM870 # IBM EBCDIC Multilingual/ROECE (Latin 2); IBM EBCDIC Multilingual Latin 2
    874: 'windows-874', # ANSI/OEM Thai (ISO 8859-11); Thai (Windows)
    875: 'cp875', # IBM EBCDIC Greek Modern
    932: 'shift_jis', # ANSI/OEM Japanese; Japanese (Shift-JIS)
    936: 'gb2312', # ANSI/OEM Simplified Chinese (PRC, Singapore); Chinese Simplified (GB2312)
    949: 'ks_c_5601-1987', # ANSI/OEM Korean (Unified Hangul Code)
    # We *must* use a custom encoding because the Python implementation differs
    # from the Microsoft implementation.
    950: 'windows-950', # ANSI/OEM Traditional Chinese (Taiwan; Hong Kong SAR, PRC); Chinese Traditional (Big5)
    1026: 'IBM1026', # IBM EBCDIC Turkish (Latin 5)
    1047: 'cp1047', # IBM EBCDIC Latin 1/Open System
    1140: 'cp1140', # IBM EBCDIC US-Canada (037 + Euro symbol); IBM EBCDIC (US-Canada-Euro)
    1141: 'cp1141', # IBM EBCDIC Germany (20273 + Euro symbol); IBM EBCDIC (Germany-Euro)
    1142: 'cp1142', # IBM EBCDIC Denmark-Norway (20277 + Euro symbol); IBM EBCDIC (Denmark-Norway-Euro)
    1143: 'cp1143', # IBM EBCDIC Finland-Sweden (20278 + Euro symbol); IBM EBCDIC (Finland-Sweden-Euro)
    1144: 'cp1144', # IBM EBCDIC Italy (20280 + Euro symbol); IBM EBCDIC (Italy-Euro)
    1145: 'cp1145', # IBM EBCDIC Latin America-Spain (20284 + Euro symbol); IBM EBCDIC (Spain-Euro)
    1146: 'cp1146', # IBM EBCDIC United Kingdom (20285 + Euro symbol); IBM EBCDIC (UK-Euro)
    1147: 'cp1147', # IBM EBCDIC France (20297 + Euro symbol); IBM EBCDIC (France-Euro)
    1148: 'cp1148ms', # IBM EBCDIC International (500 + Euro symbol); IBM EBCDIC (International-Euro)
    1149: 'cp1149', # IBM EBCDIC Icelandic (20871 + Euro symbol); IBM EBCDIC (Icelandic-Euro)
    1200: 'utf-16-le', # Unicode UTF-16, little endian byte order (BMP of ISO 10646);
    1201: 'utf-16-be', # Unicode UTF-16, big endian byte order;
    1250: 'windows-1250', # ANSI Central European; Central European (Windows)
    1251: 'windows-1251', # ANSI Cyrillic; Cyrillic (Windows)
    1252: 'windows-1252', # ANSI Latin 1; Western European (Windows)
    1253: 'windows-1253', # ANSI Greek; Greek (Windows)
    1254: 'windows-1254', # ANSI Turkish; Turkish (Windows)
    1255: 'windows-1255', # ANSI Hebrew; Hebrew (Windows)
    1256: 'windows-1256', # ANSI Arabic; Arabic (Windows)
    1257: 'windows-1257', # ANSI Baltic; Baltic (Windows)
    1258: 'windows-1258', # ANSI/OEM Vietnamese; Vietnamese (Windows)
    1361: 'Johab', # Korean (Johab)
    10000: 'macintosh', # MAC Roman; Western European (Mac)
    10001: 'x-mac-japanese', # Japanese (Mac)
    # UNSUPPORTED.
    10002: 'x-mac-chinesetrad', # MAC Traditional Chinese (Big5); Chinese Traditional (Mac)
    10003: 'x-mac-korean', # Korean (Mac)
    # UNSUPPORTED.
    10004: 'x-mac-arabic', # Arabic (Mac)
    # UNSUPPORTED.
    10005: 'x-mac-hebrew', # Hebrew (Mac)
    10006: 'x-mac-greek', # Greek (Mac)
    10007: 'x-mac-cyrillic', # Cyrillic (Mac)
    # UNSUPPORTED.
    10008: 'x-mac-chinesesimp', # MAC Simplified Chinese (GB 2312); Chinese Simplified (Mac)
    # UNSUPPORTED.
    10010: 'x-mac-romanian', # Romanian (Mac)
    # UNSUPPORTED.
    10017: 'x-mac-ukrainian', # Ukrainian (Mac)
    # UNSUPPORTED.
    10021: 'x-mac-thai', # Thai (Mac)
    10029: 'x-mac-ce', # MAC Latin 2; Central European (Mac)
    10079: 'x-mac-icelandic', # Icelandic (Mac)
    10081: 'x-mac-turkish', # Turkish (Mac)
    # UNSUPPORTED.
    10082: 'x-mac-croatian', # Croatian (Mac)
    12000: 'utf-32', # Unicode UTF-32, little endian byte order
    12001: 'utf-32BE', # Unicode UTF-32, big endian byte order
    # UNSUPPORTED.
    20000: 'x-Chinese_CNS', # CNS Taiwan; Chinese Traditional (CNS)
    # UNSUPPORTED.
    20001: 'x-cp20001', # TCA Taiwan
    # UNSUPPORTED.
    20002: 'x_Chinese-Eten', # Eten Taiwan; Chinese Traditional (Eten)
    # UNSUPPORTED.
    20003: 'x-cp20003', # IBM5550 Taiwan
    # UNSUPPORTED.
    20004: 'x-cp20004', # TeleText Taiwan
    # UNSUPPORTED.
    20005: 'x-cp20005', # Wang Taiwan
    # UNSUPPORTED.
    20105: 'x-IA5', # IA5 (IRV International Alphabet No. 5, 7-bit); Western European (IA5)
    # UNSUPPORTED.
    20106: 'x-IA5-German', # IA5 German (7-bit)
    # UNSUPPORTED.
    20107: 'x-IA5-Swedish', # IA5 Swedish (7-bit)
    # UNSUPPORTED.
    20108: 'x-IA5-Norwegian', # IA5 Norwegian (7-bit)
    20127: 'us-ascii', # US-ASCII (7-bit)
    # UNSUPPORTED.
    20261: 'x-cp20261', # T.61
    # UNSUPPORTED.
    20269: 'x-cp20269', # ISO 6937 Non-Spacing Accent
    20273: 'IBM273', # IBM EBCDIC Germany
    20277: 'cp277', # IBM EBCDIC Denmark-Norway
    20278: 'cp278', # IBM EBCDIC Finland-Sweden
    20280: 'cp280', # IBM EBCDIC Italy
    20284: 'cp284', # IBM EBCDIC Latin America-Spain
    20285: 'cp285', # IBM EBCDIC United Kingdom
    20290: 'cp290', # IBM EBCDIC Japanese Katakana Extended
    20297: 'cp297', # IBM EBCDIC France
    20420: 'cp420', # IBM EBCDIC Arabic
    # UNSUPPORTED.
    20423: 'IBM423', # IBM EBCDIC Greek
    20424: 'IBM424', # IBM EBCDIC Hebrew
    20833: 'cp833', # IBM EBCDIC Korean Extended
    20838: 'cp838', # IBM EBCDIC Thai
    20866: 'koi8-r', # Russian (KOI8-R); Cyrillic (KOI8-R)
    20871: 'cp871', # IBM EBCDIC Icelandic
    # UNSUPPORTED.
    20880: 'IBM880', # IBM EBCDIC Cyrillic Russian
    # UNSUPPORTED.
    20905: 'IBM905', # IBM EBCDIC Turkish
    # UNSUPPORTED.
    20924: 'IBM00924', # IBM EBCDIC Latin 1/Open System (1047 + Euro symbol)
    20932: 'EUC-JP', # Japanese (JIS 0208-1990 and 0212-1990)
    # UNSUPPORTED.
    20936: 'x-cp20936', # Simplified Chinese (GB2312); Chinese Simplified (GB2312-80)
    # UNSUPPORTED.
    20949: 'x-cp20949', # Korean Wansung
    21025: 'cp1025', # IBM EBCDIC Cyrillic Serbian-Bulgarian
    # UNSUPPORTED.
    21027: '', # (deprecated)
    21866: 'koi8-u', # Ukrainian (KOI8-U); Cyrillic (KOI8-U)
    28591: 'iso-8859-1', # ISO 8859-1 Latin 1; Western European (ISO)
    28592: 'iso-8859-2', # ISO 8859-2 Central European; Central European (ISO)
    28593: 'iso-8859-3', # ISO 8859-3 Latin 3
    28594: 'iso-8859-4', # ISO 8859-4 Baltic
    28595: 'iso-8859-5', # ISO 8859-5 Cyrillic
    28596: 'iso-8859-6', # ISO 8859-6 Arabic
    28597: 'iso-8859-7', # ISO 8859-7 Greek
    28598: 'iso-8859-8', # ISO 8859-8 Hebrew; Hebrew (ISO-Visual)
    28599: 'iso-8859-9', # ISO 8859-9 Turkish
    28603: 'iso-8859-13', # ISO 8859-13 Estonian
    28605: 'iso-8859-15', # ISO 8859-15 Latin 9
    # UNSUPPORTED.
    29001: 'x-Europa', # Europa 3
    # UNSUPPORTED.
    38598: 'iso-8859-8-i', # ISO 8859-8 Hebrew; Hebrew (ISO-Logical)
    50220: 'iso-2022-jp', # ISO 2022 Japanese with no halfwidth Katakana; Japanese (JIS)
    50221: 'csISO2022JP', # ISO 2022 Japanese with halfwidth Katakana; Japanese (JIS-Allow 1 byte Kana)
    50222: 'iso-2022-jp', # ISO 2022 Japanese JIS X 0201-1989; Japanese (JIS-Allow 1 byte Kana - SO/SI)
    50225: 'iso-2022-kr', # ISO 2022 Korean
    # UNSUPPORTED.
    50227: 'x-cp50227', # ISO 2022 Simplified Chinese; Chinese Simplified (ISO 2022)
    # UNSUPPORTED.
    50229: '', # ISO 2022 Traditional Chinese
    # UNSUPPORTED.
    50930: '', # EBCDIC Japanese (Katakana) Extended
    # UNSUPPORTED.
    50931: '', # EBCDIC US-Canada and Japanese
    # UNSUPPORTED.
    50933: '', # EBCDIC Korean Extended and Korean
    # UNSUPPORTED.
    50935: '', # EBCDIC Simplified Chinese Extended and Simplified Chinese
    # UNSUPPORTED.
    50936: '', # EBCDIC Simplified Chinese
    # UNSUPPORTED.
    50937: '', # EBCDIC US-Canada and Traditional Chinese
    # UNSUPPORTED.
    50939: '', # EBCDIC Japanese (Latin) Extended and Japanese
    51932: 'euc-jp', # EUC Japanese
    51936: 'EUC-CN', # EUC Simplified Chinese; Chinese Simplified (EUC)
    51949: 'euc-kr', # EUC Korean
    # UNSUPPORTED.
    51950: '', # EUC Traditional Chinese
    52936: 'hz-gb-2312', # HZ-GB2312 Simplified Chinese; Chinese Simplified (HZ)
    54936: 'GB18030', # Windows XP and later: GB18030 Simplified Chinese (4 byte); Chinese Simplified (GB18030)
    # UNSUPPORTED.
    57002: 'x-iscii-de', # ISCII Devanagari
    # UNSUPPORTED.
    57003: 'x-iscii-be', # ISCII Bangla
    # UNSUPPORTED.
    57004: 'x-iscii-ta', # ISCII Tamil
    # UNSUPPORTED.
    57005: 'x-iscii-te', # ISCII Telugu
    # UNSUPPORTED.
    57006: 'x-iscii-as', # ISCII Assamese
    # UNSUPPORTED.
    57007: 'x-iscii-or', # ISCII Odia
    # UNSUPPORTED.
    57008: 'x-iscii-ka', # ISCII Kannada
    # UNSUPPORTED.
    57009: 'x-iscii-ma', # ISCII Malayalam
    # UNSUPPORTED.
    57010: 'x-iscii-gu', # ISCII Gujarati
    # UNSUPPORTED.
    57011: 'x-iscii-pa', # ISCII Punjabi
    65000: 'utf-7', # Unicode (UTF-7)
    65001: 'utf-8', # Unicode (UTF-8)
}

# Register new encodings.

def lookupCodePage(id_: int) -> str:
    """
    Converts an encoding id into it's name.

    :raises UnknownCodepageError: The code page was not recognized.
    :raises UnsupportedEncodingError: The code page was recognized, but no
        encoding exists in the environment with support for it.
    """
    if id_ in _CODE_PAGES:
        if (page := _CODE_PAGES[id_]):
            return page
        else:
            raise UnsupportedEncodingError(f'Code page {id_} is unsupported.')
    else:
        raise UnknownCodepageError(f'Unknown code page {id_}.')

def _lookupEncoding(name):
    return _codecsInfo.get(name)

from .utils import createSBEncoding as _sb, createVBEncoding as _vb
from ._dt import (
        _mac_ce, _mac_cyrillic, _mac_greek, _mac_iceland, _mac_turkish,
        _win874_dec, _win950_dec
    )

_codecsInfo = {
    'x_mac_ce': _sb('x-mac-ce', _mac_ce.decodingTable),
    'x_mac_cyrillic': _sb('x-mac-cyrillic', _mac_cyrillic.decodingTable),
    'x_mac_greek': _sb('x-mac-greek', _mac_greek.decodingTable),
    'x_mac_icelandic': _sb('x-mac-icelandic', _mac_iceland.decodingTable),
    'x_mac_turkish': _sb('x-mac-turkish', _mac_turkish.decodingTable),
    'windows_950': _vb('windows-950', _win950_dec.decodingTable),
    'windows_874': _sb('windows-874', _win874_dec.decodingTable),
}

codecs.register(_lookupEncoding)