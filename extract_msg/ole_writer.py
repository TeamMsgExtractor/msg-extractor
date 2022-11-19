import io
import pathlib

from typing import Union


class OleWritter:
    """
    Takes data to write to a compound binary format file, as specified in
    [MS-CFB].
    """
    def __init__(self):
        raise NotImplementedError('This class is unfinished and not ready to be used.')

    def _writeHeader(self, f):
        """
        Writes the header to the file f.
        """
        # Header signature.
        f.write(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1)
        # Header CLSID.
        f.write(b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00')
        # Minor version.
        f.write(b'\x3E\x00')
        # Major version. For now, we only support version 3.
        f.write(b'\x03\x00')
        # Byte order. Specifies that it is little endian.
        f.write(b'\xFE\xFF')


    def write(self, path) -> None:
        """
        Writes the data to the path specified. If :param path: has a write
        method it will use the object directly.
        """
        opened = False

        # First, let's open the file if it is not a writable object.
        if hasattr(path, 'write') and hasattr(path.write, '__call__'):
            f = path
        else:
            f = open(path, 'wb')
            opened = True

        # Make sure we close the file after everything, especially if there is
        # an error.
        try:
            ### First we need to write the header.
            self._writeHeader(f)

        finally:
            if opened:
                f.close()
