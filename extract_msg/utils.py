from __future__ import annotations


"""
Utility functions of extract_msg.
"""


__all__ = [
    'addNumToDir',
    'addNumToZipDir',
    'bitwiseAdjust',
    'bitwiseAdjustedAnd',
    'bytesToGuid',
    'ceilDiv',
    'cloneOleFile',
    'createZipOpen',
    'decodeRfc2047',
    'dictGetCasedKey',
    'divide',
    'filetimeToDatetime',
    'filetimeToUtc',
    'findWk',
    'fromTimeStamp',
    'getCommandArgs',
    'guessEncoding',
    'htmlSanitize',
    'inputToBytes',
    'inputToMsgPath',
    'inputToString',
    'isEncapsulatedRtf',
    'makeWeakRef',
    'msgPathToString',
    'parseType',
    'prepareFilename',
    'roundUp',
    'rtfSanitizeHtml',
    'rtfSanitizePlain',
    'setupLogging',
    'stripRtf',
    'tryGetMimetype',
    'unsignedToSignedInt',
    'unwrapMsg',
    'unwrapMultipart',
    'validateHtml',
    'verifyPropertyId',
    'verifyType',
]


import argparse
import collections
import copy
import datetime
import decimal
import email.header
import email.message
import email.policy
import glob
import json
import logging
import logging.config
import os
import pathlib
import re
import shutil
import struct
import sys
import weakref
import zipfile

import bs4
import olefile
import tzlocal

from html import escape as htmlEscape
from typing import (
        Any, AnyStr, Callable, Dict, Iterable, List, Optional, Sequence,
        SupportsBytes, TypeVar, TYPE_CHECKING, Union
    )

from . import constants
from .enums import AttachmentType
from .exceptions import (
        ConversionError, DependencyError, ExecutableNotFound,
        IncompatibleOptionsError, InvalidPropertyIdError, TZError,
        UnknownTypeError
    )


# Allow for nice type checking.
if TYPE_CHECKING:
    from .msg_classes.msg import MSGFile
    from .attachments import AttachmentBase

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())
logging.addLevelName(5, 'DEVELOPER')

_T = TypeVar('_T')


def addNumToDir(dirName: pathlib.Path) -> Optional[pathlib.Path]:
    """
    Attempt to create the directory with a '(n)' appended.
    """
    for i in range(2, 100):
        try:
            newDirName = dirName.with_name(dirName.name + f' ({i})')
            os.makedirs(newDirName)
            return newDirName
        except Exception as e:
            pass
    return None


def addNumToZipDir(dirName: pathlib.Path, _zip) -> Optional[pathlib.Path]:
    """
    Attempt to create the directory with a '(n)' appended.
    """
    for i in range(2, 100):
        newDirName = dirName.with_name(dirName.name + f' ({i})')
        pathCompare = str(newDirName).rstrip('/') + '/'
        if not any(x.startswith(pathCompare) for x in _zip.namelist()):
            return newDirName
    return None


def bitwiseAdjust(inp: int, mask: int) -> int:
    """
    Uses a given mask to adjust the location of bits after an operation like
    bitwise AND.

    This is useful for things like flags where you are trying to get a small
    portion of a larger number. Say for example, you had the number ``0xED``
    (``0b11101101``) and you needed the adjusted result of the AND operation
    with ``0x70`` (``0b01110000``). The result of the AND operation
    (``0b01100000``) and the mask used to get it (``0x70``) are given to this
    function and the adjustment will be done automatically.

    :param mask: MUST be greater than 0.

    :raises ValueError: The mask is not greater than 0.
    """
    if mask < 1:
        raise ValueError('Mask MUST be greater than 0')
    return inp >> bin(mask)[::-1].index('1')


def bitwiseAdjustedAnd(inp: int, mask: int) -> int:
    """
    Preforms the bitwise AND operation between :param inp: and :param mask: and
    adjusts the results based on the rules of :func:`bitwiseAdjust`.

    :raises ValueError: The mask is not greater than 0.
    """
    if mask < 1:
        raise ValueError('Mask MUST be greater than 0')
    return (inp & mask) >> bin(mask)[::-1].index('1')


def bytesToGuid(bytesInput: bytes) -> str:
    """
    Converts a bytes instance to a GUID.
    """
    guidVals = constants.st.ST_GUID.unpack(bytesInput)
    return f'{{{guidVals[0]:08X}-{guidVals[1]:04X}-{guidVals[2]:04X}-{guidVals[3][:2].hex().upper()}-{guidVals[3][2:].hex().upper()}}}'


def ceilDiv(n: int, d: int) -> int:
    """
    Returns the ``int`` from the ceiling division of n / d.

    ONLY use ``int``\\s as inputs to this function.

    For ``int``\\s, this is faster and more accurate for numbers outside the
    precision range of ``float``.
    """
    return -(n // -d)


def cloneOleFile(sourcePath, outputPath) -> None:
    """
    Uses the ``OleWriter`` class to clone the specified OLE file into a new
    location.

    Mainly designed for testing.
    """
    from .ole_writer import OleWriter

    with olefile.OleFileIO(sourcePath) as f:
        writer = OleWriter()
        writer.fromOleFile(f)

    writer.write(outputPath)


def createZipOpen(func) -> Callable:
    """
    Creates a wrapper for the open function of a ZipFile that will automatically
    set the current date as the modified time to the current time.
    """
    def _open(name, mode = 'r', *args, **kwargs):
        if mode == 'w':
            name = zipfile.ZipInfo(name, datetime.datetime.now().timetuple()[:6])

        return func(name, mode, *args, **kwargs)

    return _open


def decodeRfc2047(encoded: str) -> str:
    """
    Decodes text encoded using the method specified in RFC 2047.
    """
    # Fix an issue with folded header fields.
    encoded = encoded.replace('\r\n', '')

    # This returns a list of tuples containing the bytes and the encoding they
    # are using, so we decode each one and join them together.
    #
    # decode_header header will return a string instead of bytes for the first
    # object if the input is not encoded, something that is frustrating.
    return ''.join(
        x[0].decode(x[1] or 'raw-unicode-escape') if isinstance(x[0], bytes) else x[0]
        for x in email.header.decode_header(encoded)
    )


def dictGetCasedKey(_dict: Dict[str, Any], key: str) -> str:
    """
    Retrieves the key from the dictionary with the proper casing using a
    caseless key.
    """
    try:
        return next((x for x in _dict.keys() if x.lower() == key.lower()))
    except StopIteration:
        # If we couldn't find the key, raise a KeyError.
        raise KeyError(key)


def divide(string: AnyStr, length: int) -> List[AnyStr]:
    """
    Divides a string into multiple substrings of equal length.

    If there is not enough for the last substring to be equal, it will simply
    use the rest of the string. Can also be used for things like lists and
    tuples.

    :param string: The string to be divided.
    :param length: The length of each division.
    :returns: list containing the divided strings.

    Example:

    .. code-block:: python

        >>> a = divide('Hello World!', 2)
        >>> print(a)
        ['He', 'll', 'o ', 'Wo', 'rl', 'd!']
        >>> a = divide('Hello World!', 5)
        >>> print(a)
        ['Hello', ' Worl', 'd!']
    """
    return [string[length * x:length * (x + 1)] for x in range(ceilDiv(len(string), length))]


def filetimeToDatetime(rawTime: int) -> datetime.datetime:
    """
    Converts a filetime into a ``datetime``.

    Some values have specialized meanings, listed below:

    * ``915151392000000000``: December 31, 4500, representing a null time.
      Returns an instance of extract_msg.null_date.NullDate.
    * ``915046235400000000``: 23:59 on August 31, 4500, representing a null
      time. Returns extract_msg.constants.NULL_DATE.
    """
    try:
        if rawTime < 116444736000000000:
            # We can't properly parse this with our current setup, so
            # we will rely on olefile to handle this one.
            return olefile.olefile.filetime2datetime(rawTime)
        elif rawTime == 915151392000000000:
            # So this is actually a different null date, specifically
            # supposed to be December 31, 4500, but it's weird that the same
            # spec has 2 different ones. It's "the last valid date." Checking
            # the value of this though, it looks like it's actually one minute
            # further in the future, according to the datetime module.
            from .null_date import NullDate
            date = NullDate(4500, 12, 31, 23, 59)
            date.filetime = rawTime
            return date
        elif rawTime == 915046235400000000:
            return constants.NULL_DATE
        elif rawTime > 915000000000000000:
            # Just make null dates from all of these time stamps.
            from .null_date import NullDate
            date = NullDate(1970, 1, 1, 1)
            try:
                date += datetime.timedelta(seconds = filetimeToUtc(rawTime))
            except OverflowError:
                # Time value is so large we physically can't represent it, so
                # let's just modify the date to it's highest possible value and
                # call it a day.
                m = date.max
                date = NullDate(m.year, m.month, m.day, m.hour, m.minute, m.second, m.microsecond)
            date.filetime = rawTime

            return date
        else:
            return fromTimeStamp(filetimeToUtc(rawTime))
    except TZError:
        # For TZError we just raise it again. It is a fatal error.
        raise
    except Exception:
        raise ValueError(f'Timestamp value of {filetimeToUtc(rawTime)} (raw: {rawTime}) caused an exception. This was probably caused by the time stamp being too far in the future.')


def filetimeToUtc(inp: int) -> float:
    """
    Converts a FILETIME into a unix timestamp.
    """
    return (inp - 116444736000000000) / 10000000.0


def findWk(path = None):
    """
    Attempt to find the path of the wkhtmltopdf executable.

    :param path: If provided, the function will verify that it is executable
        and returns the path if it is.

    :raises ExecutableNotFound: A valid executable could not be found.
    """
    if path:
        if os.path.isfile(path):
            # Check if executable.
            if os.access(path, os.X_OK):
                return path
            else:
                raise ExecutableNotFound('Path provided was not a valid executable (execution bit not set).')
        else:
            raise ExecutableNotFound('Path provided was not a valid executable (not a file).')

    candidate = shutil.which('wkhtmltopdf')
    if candidate:
        return candidate

    raise ExecutableNotFound('Could not find wkhtmltopdf.')


def fromTimeStamp(stamp: float) -> datetime.datetime:
    """
    Returns a ``datetime`` from the UTC timestamp given the current timezone.
    """
    try:
        tz = tzlocal.get_localzone()
    except Exception:
        # I know "generalized exception catching is bad" but if *any* exception
        # happens here that is a subclass of Exception then something has gone
        # wrong with tzlocal.
        raise TZError(f'Error occured using tzlocal. If you are seeing this, this is likely a problem with your installation ot tzlocal or tzdata.')
    return datetime.datetime.fromtimestamp(stamp, tz)


def getCommandArgs(args: Sequence[str]) -> argparse.Namespace:
    """
    Parse command-line arguments.

    :raises IncompatibleOptionsError: Some options were provided that are
        incompatible.
    :raises ValueError: Something about the options was invalid. This could mean
        an option was specified that requires another option or it could mean
        that an option was looking for data that was not found.
    """
    parser = argparse.ArgumentParser(description = constants.MAINDOC, prog = 'extract_msg')
    outFormat = parser.add_mutually_exclusive_group()
    inputFormat = parser.add_mutually_exclusive_group()
    inputType = parser.add_mutually_exclusive_group(required = True)
    # --use-content-id, --cid
    parser.add_argument('--use-content-id', '--cid', dest='cid', action='store_true',
                        help='Save attachments by their Content ID, if they have one. Useful when working with the HTML body.')
    # --json
    outFormat.add_argument('--json', dest='json', action='store_true',
                        help='Changes to write output files as json.')
    # --file-logging
    parser.add_argument('--file-logging', dest='fileLogging', action='store_true',
                        help='Enables file logging. Implies --verbose level 1.')
    # -v, --verbose
    parser.add_argument('-v', '--verbose', dest='verbose', action='count', default=0,
                        help='Turns on console logging. Specify more than once for higher verbosity.')
    # --log PATH
    parser.add_argument('--log', dest='log',
                        help='Set the path to write the file log to.')
    # --config PATH
    parser.add_argument('--config', dest='configPath',
                        help='Set the path to load the logging config from.')
    # --out PATH
    parser.add_argument('--out', dest='outPath',
                        help='Set the folder to use for the program output. (Default: Current directory)')
    # --use-filename
    parser.add_argument('--use-filename', dest='useFilename', action='store_true',
                        help='Sets whether the name of each output is based on the MSG filename.')
    # --dump-stdout
    parser.add_argument('--dump-stdout', dest='dumpStdout', action='store_true',
                        help='Tells the program to dump the message body (plain text) to stdout. Overrides saving arguments.')
    # --html
    outFormat.add_argument('--html', dest='html', action='store_true',
                        help='Sets whether the output should be HTML. If this is not possible, will error.')
    # --pdf
    outFormat.add_argument('--pdf', dest='pdf', action='store_true',
                           help='Saves the body as a PDF. If this is not possible, will error.')
    # --wk-path PATH
    parser.add_argument('--wk-path', dest='wkPath',
                        help='Overrides the path for finding wkhtmltopdf.')
    # --wk-options OPTIONS
    parser.add_argument('--wk-options', dest='wkOptions', nargs='*',
                        help='Sets additional options to be used in wkhtmltopdf. Should be a series of options and values, replacing the - or -- in the beginning with + or ++, respectively. For example: --wk-options "+O Landscape"')
    # --prepared-html
    parser.add_argument('--prepared-html', dest='preparedHtml', action='store_true',
                        help='When used in conjunction with --html, sets whether the HTML output should be prepared for embedded attachments.')
    # --charset
    parser.add_argument('--charset', dest='charset', default='utf-8',
                        help='Character set to use for the prepared HTML in the added tag. (Default: utf-8)')
    # --raw
    outFormat.add_argument('--raw', dest='raw', action='store_true',
                           help='Sets whether the output should be raw. If this is not possible, will error.')
    # --rtf
    outFormat.add_argument('--rtf', dest='rtf', action='store_true',
                           help='Sets whether the output should be RTF. If this is not possible, will error.')
    # --allow-fallback
    parser.add_argument('--allow-fallback', dest='allowFallback', action='store_true',
                        help='Tells the program to fallback to a different save type if the selected one is not possible.')
    # --skip-body-not-found
    parser.add_argument('--skip-body-not-found', dest='skipBodyNotFound', action='store_true',
                        help='Skips saving the body if the body cannot be found, rather than throwing an error.')
    # --zip
    parser.add_argument('--zip', dest='zip',
                        help='Path to use for saving to a zip file.')
    # --save-header
    parser.add_argument('--save-header', dest='saveHeader', action='store_true',
                        help='Store the header in a separate file.')
    # --attachments-only
    outFormat.add_argument('--attachments-only', dest='attachmentsOnly', action='store_true',
                           help='Specify to only save attachments from an MSG file.')
    # --skip-hidden
    parser.add_argument('--skip-hidden', dest='skipHidden', action='store_true',
                        help='Skips any attachment marked as hidden (usually ones embedded in the body).')
    # --no-folders
    parser.add_argument('--no-folders', dest='noFolders', action='store_true',
                        help='Stores everything in the location specified by --out. Requires --attachments-only and is incompatible with --out-name.')
    # --skip-embedded
    parser.add_argument('--skip-embedded', dest='skipEmbedded', action='store_true',
                        help='Skips all embedded MSG files when saving attachments.')
    # --extract-embedded
    parser.add_argument('--extract-embedded', dest='extractEmbedded', action='store_true',
                        help='Extracts the embedded MSG files as MSG files instead of running their save functions.')
    # --overwrite-existing
    parser.add_argument('--overwrite-existing', dest='overwriteExisting', action='store_true',
                        help='Disables filename conflict resolution code for attachments when saving a file, causing files to be overwriten if two attachments with the same filename are on an MSG file.')
    # --skip-not-implemented
    parser.add_argument('--skip-not-implemented', '--skip-ni', dest='skipNotImplemented', action='store_true',
                        help='Skips any attachments that are not implemented, allowing saving of the rest of the message.')
    # --out-name NAME
    inputFormat.add_argument('--out-name', dest='outName',
                        help='Name to be used with saving the file output. Cannot be used if you are saving more than one file.')
    # --glob
    inputFormat.add_argument('--glob', '--wildcard', dest='glob', action='store_true',
                        help='Interpret all paths as having wildcards. Incompatible with --out-name.')
    # --ignore-rtfde
    parser.add_argument('--ignore-rtfde', dest='ignoreRtfDeErrors', action='store_true',
                        help='Ignores all errors thrown from RTFDE when trying to save. Useful for allowing fallback to continue when an exception happens.')
    # --progress
    parser.add_argument('--progress', dest='progress', action='store_true',
                        help='Shows what file the program is currently working on during it\'s progress.')
    # -s, --stdout
    inputType.add_argument('-s', '--stdin', dest='stdin', action='store_true',
                        help='Read file from stdin (only works with one file at a time).')
    # [MSG files]
    inputType.add_argument('msgs', metavar='msg', nargs='*', default=[],
                        help='An MSG file to be parsed.')

    options = parser.parse_args(args)

    if options.stdin:
        # Read the MSG file from stdin and shove it into the msgs list.
        options.msgs.append(sys.stdin.buffer.read())

    if options.outName and options.noFolders:
        raise IncompatibleOptionsError('--out-name is not compatible with --no-folders.')

    if options.fileLogging:
        options.verbose = options.verbose or 1

    # Handle the wkOptions if they exist.
    if options.wkOptions:
        wkOptions = []
        for option in options.wkOptions:
            if option.startswith('++'):
                option = '--' + option[2:]
            elif option.startswith('+'):
                option = '-' + option[1:]

            # Now that we have corrected to the correct start, split the argument if
            # necessary.
            split = option.split(' ')
            if len(split) == 1:
                # No spaces means we just pass that directly.
                wkOptions.append(option)
            else:
                wkOptions.append(split[0])
                wkOptions.append(' '.join(split[1:]))

        options.wkOptions = wkOptions

    # If dump_stdout is True, we need to unset all arguments used in files.
    # Technically we actually only *need* to unset `out_path`, but that may
    # change in the future, so let's be thorough.
    if options.dumpStdout:
        options.outPath = None
        options.json = False
        options.rtf = False
        options.html = False
        options.useFilename = False
        options.cid = False

    if options.glob:
        if options.outName:
            raise IncompatibleOptionsError('--out-name is not supported when using wildcards.')
        if options.stdin:
            raise IncompatibleOptionsError('--stdin is not supported with using wildcards.')
        fileLists = []
        for path in options.msgs:
            fileLists += glob.glob(path)

        if len(fileLists) == 0:
            raise ValueError('Could not find any MSG files using the specified wildcards.')
        options.msgs = fileLists

    # Make it so outName can only be used on single files.
    if options.outName and len(options.msgs) > 1:
        raise IncompatibleOptionsError('--out-name is not supported when saving multiple MSG files.')

    # Handle the verbosity level.
    if options.verbose == 0:
        options.logLevel = logging.ERROR
    elif options.verbose == 1:
        options.logLevel = logging.WARNING
    elif options.verbose == 2:
        options.logLevel = logging.INFO
    else:
        options.logLevel = 5

    # If --no-folders is turned on but --attachments-only is not, error.
    if options.noFolders and not options.attachmentsOnly:
        raise ValueError('--no-folders requires the --attachments-only option.')

    return options


def guessEncoding(msg: MSGFile) -> Optional[str]:
    """
    Analyzes the strings on an MSG file and attempts to form a consensus about the encoding based on the top-level strings.

    Returns ``None`` if no consensus could be formed.

    :raises DependencyError: ``chardet`` is not installed or could not be used
        properly.
    """
    try:
        import chardet
    except ImportError:
        raise DependencyError('Cannot guess the encoding of an MSG file if chardet is not installed.')

    data = b''
    for name in (x[0] for x in msg.listDir(True, False, False) if len(x) == 1):
        if name.lower().endswith('001f'):
            # This is a guarentee.
            return 'utf-16-le'
        elif name.lower().endswith('001e'):
            data += msg.getStream(name) + b'\n'

    try:
        if not data or (result := chardet.detect(data))['confidence'] < 0.5:
            return None

        return result['encoding']
    except Exception as e:
        raise DependencyError(f'Failed to detect encoding: {e}')


def htmlSanitize(inp: str) -> str:
    """
    Santizes the input for injection into an HTML string.

    Converts charactersinto forms that will not be misinterpreted, if
    necessary.
    """
    # First step, do a basic escape of the HTML.
    inp = htmlEscape(inp)

    # Change newlines to <br/> to they won't be ignored.
    inp = inp.replace('\r\n', '\n').replace('\n', '<br/>')

    # Escape long sections of spaces to ensure they won't be ignored.
    inp = constants.re.HTML_SAN_SPACE.sub((lambda spaces: '&nbsp;' * len(spaces.group(0))),inp)

    return inp


def inputToBytes(obj: Union[bytes, None, str, SupportsBytes], encoding: str) -> bytes:
    """
    Converts the input into bytes.

    :raises ConversionError: The input cannot be converted.
    :raises UnicodeEncodeError: The input was a str but the encoding was not
        valid.
    :raises TypeError: The input has a __bytes__ method, but it failed.
    :raises ValueError: Same as above.
    """
    if isinstance(obj, bytes):
        return obj
    if isinstance(obj, str):
        return obj.encode(encoding)
    if obj is None:
        return b''
    if hasattr(obj, '__bytes__'):
        return bytes(obj)

    raise ConversionError('Cannot convert to bytes.')


def inputToMsgPath(inp: constants.MSG_PATH) -> List[str]:
    """
    Converts the input into an msg path.

    :raises ValueError: The path contains an illegal character.
    """
    if isinstance(inp, (list, tuple)):
        inp = '/'.join(inp)

    inp = inputToString(inp, 'utf-8')

    # Validate the path is okay. Normally we would check for '/' and '\', but
    # we are expecting a string or similar which will use those as path
    # separators, so we will ignore that for now.
    if ':' in inp or '!' in inp:
        raise ValueError('Illegal character ("!" or ":") found in MSG path.')

    ret = [x for x in inp.replace('\\', '/').split('/') if x]

    # One last thing to check: all path segments can be, at most, 31 characters
    # (32 if you include the null character), so we should verify that.
    if any(len(x) > 31 for x in ret):
        raise ValueError('Path segments must not be greater than 31 characters.')
    return ret


def inputToString(bytesInputVar: Optional[Union[str, bytes]], encoding: str) -> str:
    """
    Converts the input into a string.

    :raises ConversionError: The input cannot be converted.
    """
    if isinstance(bytesInputVar, str):
        return bytesInputVar
    elif isinstance(bytesInputVar, bytes):
        return bytesInputVar.decode(encoding)
    elif bytesInputVar is None:
        return ''
    else:
        raise ConversionError('Cannot convert to str type.')


def isEncapsulatedRtf(inp: bytes) -> bool:
    """
    Checks if the RTF data has encapsulated HTML.

    Currently the detection is made to be *extremly* basic, but this will work
    for now. In the future this will be fixed so that literal text in the body
    of a message won't cause false detection.
    """
    return b'\\fromhtml' in inp


def makeWeakRef(obj: Optional[_T]) -> Optional[weakref.ReferenceType[_T]]:
    """
    Attempts to return a weak reference to the object, returning None if not
    possible.
    """
    if obj is None:
        return None
    else:
        return weakref.ref(obj)


def minutesToDurationStr(minutes: int) -> str:
    """
    Converts the number of minutes into a duration string.
    """
    if minutes == 0:
        return '0 hours'
    elif minutes == 1:
        return '1 minute'
    elif minutes < 60:
        return f'{minutes} minutes'
    elif minutes == 60:
        return '1 hour'
    elif minutes % 60 == 0:
        return f'{minutes // 60} hours'
    elif minutes < 120:
        if minutes == 61:
            return f'1 hour 1 minute'
        else:
            return f'1 hour {minutes - 60} minutes'
    elif minutes % 60 == 1:
        return f'{minutes // 60} hours 1 minute'
    else:
        return f'{minutes // 60} hours {minutes % 60} minutes'


def msgPathToString(inp: Union[str, Iterable[str]]) -> str:
    """
    Converts an MSG path (one of the internal paths inside an MSG file) into a
    string.
    """
    if not isinstance(inp, str):
        inp = '/'.join(inp)
    return inp.replace('\\', '/')


def parseType(_type: int, stream: Union[int, bytes], encoding: str, extras: Sequence[bytes]):
    """
    Converts the data in :param stream: to a much more accurate type, specified
    by :param _type:.

    :param _type: The data's type.
    :param stream: The data to be converted.
    :param encoding: The encoding to be used for regular strings.
    :param extras: Used in the case of types like PtypMultipleString. For that
        example, extras should be a list of the bytes from rest of the streams.

    :raises NotImplementedError: The type has no current support. Most of these
        types have no documentation in [MS-OXMSG].
    """
    # WARNING Not done. Do not try to implement anywhere where it is not already implemented.
    value = stream
    lengthExtras = len(extras)
    if _type == 0x0000:  # PtypUnspecified
        pass
    elif _type == 0x0001:  # PtypNull
        if value != b'\x00\x00\x00\x00\x00\x00\x00\x00':
            # DEBUG
            logger.warning('Property type is PtypNull, but is not equal to 0.')
        return None
    elif _type == 0x0002:  # PtypInteger16
        return constants.st.ST_LE_UI16.unpack(value[:2])[0]
    elif _type == 0x0003:  # PtypInteger32
        return constants.st.ST_LE_UI32.unpack(value[:4])[0]
    elif _type == 0x0004:  # PtypFloating32
        return constants.st.ST_LE_F32.unpack(value[:4])[0]
    elif _type == 0x0005:  # PtypFloating64
        return constants.st.ST_LE_F64.unpack(value)[0]
    elif _type == 0x0006:  # PtypCurrency
        return decimal.Decimal((constants.st.ST_LE_I64.unpack(value))[0]) / 10000
    elif _type == 0x0007:  # PtypFloatingTime
        value = constants.st.ST_LE_F64.unpack(value)[0]
        return constants.PYTPFLOATINGTIME_START + datetime.timedelta(days = value)
    elif _type == 0x000A:  # PtypErrorCode
        from .enums import ErrorCode, ErrorCodeType
        value = constants.st.ST_LE_UI32.unpack(value[:4])[0]
        try:
            value = ErrorCodeType(value)
        except ValueError:
            logger.warning(f'Error type found that was not from Additional Error Codes. Value was {value}. You should report this to the developers.')
            # So here, the value should be from Additional Error Codes, but it
            # wasn't. So we are just returning the int. However, we want to see
            # if it is a normal error code.
            try:
                logger.warning(f'REPORT TO DEVELOPERS: Error type of {ErrorCode(value)} was found.')
            except ValueError:
                pass
        return value
    elif _type == 0x000B:  # PtypBoolean
        return constants.st.ST_LE_UI16.unpack(value[:2])[0] != 0
    elif _type == 0x000D:  # PtypObject/PtypEmbeddedTable
        # TODO parsing for this.
        # Wait, that's the extension for an attachment folder, so parsing this
        # might not be as easy as we would hope. The function may be released
        # without support for this.
        raise NotImplementedError('Current version of extract-msg does not support the parsing of PtypObject/PtypEmbeddedTable in this function.')
    elif _type == 0x0014:  # PtypInteger64
        return constants.st.ST_LE_UI64.unpack(value)[0]
    elif _type == 0x001E:  # PtypString8
        return value.decode(encoding)
    elif _type == 0x001F:  # PtypString
        return value.decode('utf-16-le')
    elif _type == 0x0040:  # PtypTime
        rawTime = constants.st.ST_LE_UI64.unpack(value)[0]
        return filetimeToDatetime(rawTime)
    elif _type == 0x0048:  # PtypGuid
        return bytesToGuid(value)
    elif _type == 0x00FB:  # PtypServerId
        count = constants.st.ST_LE_UI16.unpack(value[:2])[0]
        # If the first byte is a 1 then it uses the ServerID structure.
        if value[3] == 1:
            from .structures.misc_id import ServerID
            return ServerID(value)
        else:
            return (count, value[2:count + 2])
    elif _type == 0x00FD:  # PtypRestriction
        # TODO parsing for this.
        raise NotImplementedError('Parsing for type 0x00FD (PtypRestriction) has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType 0x00FD PtypRestriction".')
    elif _type == 0x00FE:  # PtypRuleAction
        # TODO parsing for this.
        raise NotImplementedError('Parsing for type 0x00FE (PtypRuleAction) has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType 0x00FE PtypRuleAction".')
    elif _type == 0x0102:  # PtypBinary
        return value
    elif _type & 0x1000 == 0x1000:  # PtypMultiple
        # TODO parsing for remaining "multiple" types.
        if _type in (0x101F, 0x101E): # PtypMultipleString/PtypMultipleString8
            ret = [x.decode(encoding)[:-1] for x in extras]
            lengths = struct.unpack(f'<{len(ret)}i', stream)
            lengthLengths = len(lengths)
            if lengthLengths > lengthExtras:
                logger.warning(f'Error while parsing multiple type. Expected {lengthLengths} stream{"s" if lengthLengths != 1 else ""}, got {lengthExtras}. Ignoring.')
            for x, y in enumerate(extras):
                if lengths[x] != len(y):
                    logger.warning(f'Error while parsing multiple type. Expected length {lengths[x]}, got {len(y)}. Ignoring.')
            return ret
        elif _type == 0x1102: # PtypMultipleBinary
            ret = copy.deepcopy(extras)
            lengths = tuple(constants.st.ST_LE_UI32.unpack(stream[pos*8:pos*8+4])[0] for pos in range(len(stream) // 8))
            lengthLengths = len(lengths)
            if lengthLengths > lengthExtras:
                logger.warning(f'Error while parsing multiple type. Expected {lengthLengths} stream{"s" if lengthLengths != 1 else ""}, got {lengthExtras}. Ignoring.')
            for x, y in enumerate(extras):
                if lengths[x] != len(y):
                    logger.warning(f'Error while parsing multiple type. Expected length {lengths[x]}, got {len(y)}. Ignoring.')
            return ret
        elif _type in (0x1002, 0x1003, 0x1004, 0x1005, 0x1007, 0x1014, 0x1040, 0x1048):
            if stream != len(extras):
                logger.warning(f'Error while parsing multiple type. Expected {stream} entr{"y" if stream == 1 else "ies"}, got {len(extras)}. Ignoring.')
            if _type == 0x1002: # PtypMultipleInteger16
                return tuple(constants.st.ST_LE_UI16.unpack(x)[0] for x in extras)
            if _type == 0x1003: # PtypMultipleInteger32
                return tuple(constants.st.ST_LE_UI32.unpack(x)[0] for x in extras)
            if _type == 0x1004: # PtypMultipleFloating32
                return tuple(constants.st.ST_LE_F32.unpack(x)[0] for x in extras)
            if _type == 0x1005: # PtypMultipleFloating64
                return tuple(constants.st.ST_LE_F64.unpack(x)[0] for x in extras)
            if _type == 0x1007: # PtypMultipleFloatingTime
                values = (constants.st.ST_LE_F64.unpack(x)[0] for x in extras)
                return tuple(constants.PYTPFLOATINGTIME_START + datetime.timedelta(days = amount) for amount in values)
            if _type == 0x1014: # PtypMultipleInteger64
                return tuple(constants.st.ST_LE_UI64.unpack(x)[0] for x in extras)
            if _type == 0x1040: # PtypMultipleTime
                return tuple(filetimeToUtc(constants.st.ST_LE_UI64.unpack(x)[0]) for x in extras)
            if _type == 0x1048: # PtypMultipleGuid
                return tuple(bytesToGuid(x) for x in extras)
        else:
            raise NotImplementedError(f'Parsing for type {_type} has not yet been implmented. If you need this type, please create a new issue labeled "NotImplementedError: parseType {_type}".')
    return value


def prepareFilename(filename: str) -> str:
    """
    Adjusts :param filename: so that it can succesfully be used as an actual
    file name.
    """
    # I would use re here, but it tested to be slightly slower than this.
    return ''.join(i for i in filename if i not in r'\/:*?"<>|' + '\x00').strip()


def roundUp(inp: int, mult: int) -> int:
    """
    Rounds :param inp: up to the nearest multiple of :param mult:.
    """
    return inp + (mult - inp) % mult


def rtfSanitizeHtml(inp: str) -> str:
    """
    Sanitizes input to an RTF stream that has encapsulated HTML.
    """
    if not inp:
        return ''
    output = ''
    for char in inp:
        # Check if it is in the right range to be printed directly.
        if 32 <= ord(char) < 128:
            # Quick check for handling the HTML escapes. Will eventually
            # upgrade this code to actually handle all the HTML escapes
            # but this will do for now.
            if char == '<':
                output += r'{\*\htmltag84 &lt;}\htmlrtf <\htmlrtf0 '
            elif char == '>':
                output += r'{\*\htmltag84 &gt;}\htmlrtf >\htmlrtf0'
            else:
                if char in ('\\', '{', '}'):
                    output += '\\'
                output += char
        elif ord(char) < 32 or 128 <= ord(char) <= 255:
            # Otherwise, see if it is just a small escape.
            output += f"\\'{ord(char):02X}"
        else:
            # Handle Unicode characters.
            enc = char.encode('utf-16-le')
            output += ''.join(f'\\u{x}?' for x in struct.unpack(f'<{len(enc) // 2}h', enc))

    return output


def rtfSanitizePlain(inp: str) -> str:
    """
    Sanitizes input to a plain RTF stream.
    """
    if not inp:
        return ''
    output = ''
    for char in inp:
        # Check if it is in the right range to be printed directly.
        if 32 <= ord(char) < 128:
            if char in ('\\', '{', '}'):
                output += '\\'
            output += char
        elif ord(char) < 32 or 128 <= ord(char) <= 255:
            # Otherwise, see if it is just a small escape.
            output += f"\\'{ord(char):02X}"
        else:
            # Handle Unicode characters.
            enc = char.encode('utf-16-le')
            output += ''.join(f'\\u{x}?' for x in struct.unpack(f'<{len(enc) // 2}h', enc))

    return output


def setupLogging(defaultPath = None, defaultLevel = logging.WARN, logfile = None, enableFileLogging: bool = False,
                  env_key = 'EXTRACT_MSG_LOG_CFG') -> bool:
    """
    Setup logging configuration

    :param defaultPath: Default path to use for the logging configuration file.
    :param defaultLevel: Default logging level.
    :param env_key: Environment variable name to search for, for setting logfile
        path.
    :param enableFileLogging: Whether to use a file to log or not.

    :returns: ``True`` if the configuration file was found and applied,
        ``False`` otherwise
    """
    shippedConfig = pathlib.Path(__file__).parent / 'data' / 'logging-config'
    if os.name == 'nt':
        null = 'NUL'
        shippedConfig /= 'logging-nt.json'
    elif os.name == 'posix':
        null = '/dev/null'
        shippedConfig /= 'logging-posix.json'
    # Find logging.json if not provided
    defaultPath = pathlib.Path(defaultPath) if defaultPath else shippedConfig

    paths = [
        defaultPath,
        pathlib.Path('logging.json'),
        pathlib.Path('../logging.json'),
        pathlib.Path('../../logging.json'),
        shippedConfig,
    ]

    path = None

    for configPath in paths:
        if configPath.exists():
            path = configPath
            break

    value = os.getenv(env_key, None)
    if value and os.path.exists(value) and os.path.isfile(value):
        path = pathlib.Path(value)

    if not path:
        print('Unable to find logging.json configuration file')
        print('Make sure a valid logging configuration file is referenced in the defaultPath'
              ' argument, is inside the extract_msg install location, or is available at one '
              'of the following file-paths:')
        print(str(paths[1:]))
        logging.basicConfig(level = defaultLevel)
        logging.warning('The extract_msg logging configuration was not found - using a basic configuration.'
                        f'Please check the extract_msg installation directory for "logging-{os.name}.json".')
        return False

    with open(path, 'rt') as f:
        config = json.load(f)

    for x in config['handlers']:
        if 'filename' in config['handlers'][x]:
            if enableFileLogging:
                config['handlers'][x]['filename'] = tmp = os.path.expanduser(
                    os.path.expandvars(logfile if logfile else config['handlers'][x]['filename']))
                tmp = pathlib.Path(tmp).parent
                if not tmp.exists:
                    os.makedirs(tmp)
            else:
                config['handlers'][x]['filename'] = null

    try:
        logging.config.dictConfig(config)
    except ValueError as e:
        print('Failed to configure the logger. Did your installation get messed up?')
        print(e)

    logging.getLogger().setLevel(defaultLevel)
    return True


def stripRtf(rtfBody: bytes) -> bytes:
    """
    Cleans up RTF before sending it to RTFDE.

    Attempts to find common sections of RTF data that will
    """
    # First, do a pre-strip to try and simplify ignored sections as much as possible.
    rtfBody = constants.re.RTF_BODY_STRIP_PRE_OPEN.sub(_stripRtfOpenHelper, rtfBody)
    rtfBody = constants.re.RTF_BODY_STRIP_PRE_CLOSE.sub(_stripRtfCloseHelper, rtfBody)
    # Second do an initial strip to simplify our data stream.
    rtfBody = constants.re.RTF_BODY_STRIP_INIT.sub(b'', rtfBody)
    # Do it one more time to help with some things that might not have gotten
    # caught the first time, perhaps because something now exists after
    # stripping.
    rtfBody = constants.re.RTF_BODY_STRIP_INIT.sub(b'', rtfBody)

    # TODO: Further processing...

    return rtfBody

def _stripRtfCloseHelper(match: re.Match) -> bytes:
    if (ret := match.expand(b'\\g<0>')).count(b'\\htmlrtf0') > 1:
        return ret

    if b'\\f' in ret:
        return ret

    return b'\\htmlrtf}\\htmlrtf0 '


def _stripRtfOpenHelper(match: re.Match) -> bytes:
    if b'\\f' in (ret := match.expand(b'\\g<0>')):
        return ret

    return b'\\htmlrtf{\\htmlrtf0 '


def _stripRtfHelper(match: re.Match) -> bytes:
    res = match.string

    # If these don't match, don't even try.
    if res.count(b'{') != res.count(b'}') or res.count(b'{') == 0:
        return res

    # If any group markers are prefixed by a backslash, give up.
    if res.find(b'\\{') != -1 or res.find(b'\\}') != -1:
        return res

    # Last little bit of processing to validate everything. We know the {}
    # match, but let's be *absolutely* sure.
    # TODO

    return res




def tryGetMimetype(att: AttachmentBase, mimetype: Union[str, None]) -> Union[str, None]:
    """
    Uses an optional dependency to try and get the mimetype of an attachment.

    If the mimetype has already been found, the optional dependency does not
    exist, or an error occurs in the optional dependency, then the provided
    mimetype is returned.

    :param att: The attachment to use for getting the mimetype.
    :param mimetype: The mimetype acquired directly from an attachment stream.
        If this value evaluates to ``False``, the function will try to
        determine it.
    """
    if mimetype:
        return mimetype

    # We only try anything if the data is bytes.
    if att.dataType is bytes:
        # Try to import our dependency module to use it.
        try:
            import magic # pyright: ignore

            if isinstance(att.data, (str, bytes)):
                return magic.from_buffer(att.data, mime = True)
        except ImportError:
            logger.info('Mimetype not found on attachment, and `mime` dependency not installed. Won\'t try to generate.')

        except Exception:
            logger.exception('Error occured while using python-magic. This error will be ignored.')

    return mimetype


def unsignedToSignedInt(uInt: int) -> int:
    """
    Convert the bits of an unsigned int (32-bit) to a signed int.

    :raises ValueError: The number was not valid.
    """
    if uInt > 0xFFFFFFFF:
        raise ValueError('Value is too large.')
    if uInt < 0:
        raise ValueError('Value is already signed.')
    return constants.st.ST_SBO_I32.unpack(constants.st.ST_SBO_UI32.pack(uInt))[0]


def unwrapMsg(msg: MSGFile) -> Dict[str, List]:
    """
    Takes a recursive message-attachment structure and unwraps it into a linear
    dictionary for easy iteration.

    Dictionary contains 4 keys: "attachments" for main message attachments, not
    including embedded MSG files, "embedded" for attachments representing
    embedded MSG files, "msg" for all MSG files (including the original in the
    first index), and "raw_attachments" for raw attachments from signed
    messages.
    """
    from .msg_classes import MessageSignedBase

    # Here is where we store main attachments.
    attachments = []
    # Here is where we are going to store embedded MSG files.
    msgFiles = [msg]
    # Here is where we store embedded attachments.
    embedded = []
    # Here is where we store raw attachments from signed messages.
    raw = []

    # Normally we would need a recursive function to unwrap a recursive
    # structure like the message-attachment structure. Essentially, a function
    # that calls itself. Here, I have designed code capable of circumventing
    # this to do it in a single function, which is a lot more efficient and
    # safer. That is why we store the `toProcess` and use a while loop
    # surrounding a for loop. The for loop would be the main body of the
    # function, while the append to toProcess would be the recursive call.
    toProcess = collections.deque((msg,))

    while len(toProcess) > 0:
        # Remove the last item from the list of things to process, and store it
        # in `currentItem`. We will be processing it in the for loop.
        currentItem = toProcess.popleft()
        # iterate through the attachments and
        for att in currentItem.attachments:
            # If it is a regular attachment, add it to the list. Otherwise, add
            # it to be processed
            if att.type not in (AttachmentType.MSG, AttachmentType.SIGNED_EMBEDDED):
                attachments.append(att)
            else:
                # Here we do two things. The first is we store it to the output
                # so we can return it. The second is we add it to the processing
                # list. The reason this is two steps is because we need to be
                # able to remove items from the processing list, but can't
                # do that from the output.
                embedded.append(att)
                msgFiles.append(att.data)
                toProcess.append(att.data)
        if isinstance(currentItem, MessageSignedBase):
            raw += currentItem.rawAttachments

    return {
        'attachments': attachments,
        'embedded': embedded,
        'msg': msgFiles,
        'raw_attachments': raw,
    }


def unwrapMultipart(mp: Union[bytes, str, email.message.Message]) -> Dict:
    """
    Unwraps a recursive multipart structure into a dictionary of linear lists.

    Similar to unwrapMsg, but for multipart. The dictionary contains 3 keys:
    "attachments" which contains a list of ``dict``\\s containing processed
    attachment data as well as the Message instance associated with it,
    "plain_body" which contains the plain text body, and "html_body" which
    contains the HTML body.

    For clarification, each instance of processed attachment data is a ``dict``
    with keys identical to the args used for the ``SignedAttachment``
    constructor. This makes it easy to expand for use in constructing a
    ``SignedAttachment``. The only argument missing is "msg" to ensure this function will not require one.

    :param mp: The bytes that make up a multipart, the string that makes up a
        multipart, or a ``Message`` instance from the ``email`` module created
        from the multipart data to unwrap. If providing a ``Message`` instance,
        prefer it to be an instance of ``EmailMessage``. If you are doing so,
        make sure it's policy is default.
    """
    # In the event we are generating it, these are the kwargs to use.
    genKwargs = {
        '_class': email.message.EmailMessage,
        'policy': email.policy.default,
    }
    # Convert our input into something usable.
    if isinstance(mp, email.message.EmailMessage):
        if mp.policy == email.policy.default:
            mpMessage = mp
        else:
            mpMessage = email.message_from_bytes(mp.as_bytes(), **genKwargs)
    elif isinstance(mp, email.message.Message):
        mpMessage = email.message_from_bytes(mp.as_bytes(), **genKwargs)
    elif isinstance(mp, bytes):
        mpMessage = email.message_from_bytes(mp, **genKwargs)
    elif isinstance(mp, str):
        mpMessage = email.message_from_string(mp, **genKwargs)
    else:
        raise TypeError(f'Unsupported type "{type(mp)}" provided to unwrapMultipart.')

    # Okay, now that we have it in a useable form, let's do the most basic
    # unwrapping possible. Once the most basic unwrapping is done, we can
    # actually process the data. For this, we only care if the section is
    # multipart or not. If it is, it get's unwrapped too.
    #
    # In case you are curious, this is effectively doing a breadth first
    # traversal of the tree.
    dataNodes = []

    toProcess = collections.deque((mpMessage,))
    # I do know about the walk method, but it might *also* walk embedded
    # messages which we very much don't want.
    while len(toProcess) > 0:
        currentItem = toProcess.popleft()
        # 'multipart' indicates that it shouldn't contain any data itself, just
        # other nodes to go through.
        if currentItem.get_content_maintype() == 'multipart':
            payload = currentItem.get_payload()
            # For multipart, the payload should be a list, but handle it not
            # being one.
            if isinstance(payload, list):
                toProcess.extend(payload)
            else:
                logging.warning('Found multipart node that did not return a list. Appending as a data node.')
                dataNodes.append(currentItem)
        else:
            # The opposite is *not* true. If it's not multipart, always add as a
            # data node.
            dataNodes.append(currentItem)

    # At this point, all of our nodes should have processed and we should now
    # have data nodes. Now let's process them. For anything that was parsed as
    # a message, we actually want to get it's raw bytes back so it can be saved.
    # If they user wants to process that message in some way, they can do it
    # themself.
    attachments = []
    plainBody = None
    htmlBody = None

    for node in dataNodes:
        # Let's setup our attachment we are going to use.
        attachment = {
            'data': None,
            'name': node.get_filename(),
            'mimetype': node.get_content_type(),
            'node': node,
        }

        # Finally, we need to get the data. As we need to ensure it is bytes,
        # we may have to do some special processing.
        data = node.get_content()
        if isinstance(data, bytes):
            # If the data is bytes, we are perfectly good.
            pass
        elif isinstance(data, email.message.Message):
            # If it is a message, get it's bytes directly.
            data = data.as_bytes()
        elif isinstance(data, str):
            # If it is a string, let's reverse encode it where possible.
            # First thing we want to check is if we can find the encoding type.
            # If we can, use that to reverse the process. Otherwise use utf-8.
            data = data.encode(node.get_content_charset('utf-8'))
        else:
            # We throw an exception to describe the problem if we can't reverse
            # the problem.
            raise TypeError(f'Attempted to get bytes for attachment, but could not convert {type(data)} to bytes.')

        attachment['data'] = data

        # Now for the fun part, figuring out if we actually have an attachment.
        if attachment['name']:
            attachments.append(attachment)
        elif attachment['mimetype'] == 'text/plain':
            if plainBody:
                logger.warning('Found multiple candidates for plain text body.')
            plainBody = data
        elif attachment['mimetype'] == 'text/html':
            if htmlBody:
                logger.warning('Found multiple candidates for HTML body.')
            htmlBody = data

    return {
        'attachments': attachments,
        'plain_body': plainBody,
        'html_body': htmlBody,
    }


def validateHtml(html: bytes, encoding: Optional[str]) -> bool:
    """
    Checks whether the HTML is considered valid.

    To be valid, the HTML must, at minimum, contain an ``<html>`` tag, a
    ``<body>`` tag, and closing tags for each.
    """
    bs = bs4.BeautifulSoup(html, 'html.parser', from_encoding = encoding)
    if not bs.find('html') or not bs.find('body'):
        return False
    return True


def verifyPropertyId(id: str) -> None:
    """
    Determines whether a property ID is valid for certain functions.

    Property IDs MUST be a 4 digit hexadecimal string. Property is valid if no
    exception is raised.

    :raises InvalidPropertyIdError: The ID is not a 4 digit hexadecimal number.
    """
    if not isinstance(id, str):
        raise InvalidPropertyIdError('ID was not a 4 digit hexadecimal string')
    if len(id) != 4:
        raise InvalidPropertyIdError('ID was not a 4 digit hexadecimal string')
    try:
        int(id, 16)
    except ValueError:
        raise InvalidPropertyIdError('ID was not a 4 digit hexadecimal string')


def verifyType(_type: Optional[str]) -> None:
    """
    Verifies that the type is valid.

    Raises an exception if it is not.

    :raises UnknownTypeError: The type is not recognized.
    """
    if _type is not None:
        if (_type not in constants.VARIABLE_LENGTH_PROPS_STRING) and (_type not in constants.FIXED_LENGTH_PROPS_STRING):
            raise UnknownTypeError(f'Unknown type {_type}.')
