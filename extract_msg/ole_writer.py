import io
import pathlib

from typing import Union

from . import constants
from .utils import ceilDiv


class OleWriter:
    """
    Takes data to write to a compound binary format file, as specified in
    [MS-CFB].
    """
    def __init__(self):
        raise NotImplementedError('This class is unfinished and not ready to be used.')

        self.__numberOfSectors = 0

    def _writeBeginning(self, f):
        """
        Writes the beginning to the file :param f:. This includes the header,
        DIFAT, and FAT blocks.
        """
        # Since we are going to need these multiple times, get them now.
        numFat, numDifat, totalSectors = self._getFatSectors()

        # Header signature.
        f.write(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1')
        # Header CLSID.
        f.write(b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00')
        # Minor version.
        f.write(b'\x3E\x00')
        # Major version. For now, we only support version 3.
        f.write(b'\x03\x00')
        # Byte order. Specifies that it is little endian.
        f.write(b'\xFE\xFF')
        # Sector shift.
        f.write(b'\x09\x00')
        # Mini sector shift.
        f.write(b'\x06\x00')
        # Reserved.
        f.write(b'\x00\x00\x00\x00\x00\x00')
        # Number of directory sectors. Version 3 says this *must* be 0.
        f.write(constants.ST_LE_UI32.pack(0))
        # Number of FAT sectors.
        f.write(constants.ST_LE_UI32.pack(numFat))
        # First directory sector location (Sector for the directory stream).
        # We place that right after the DIFAT and FAT.
        f.write(constants.ST_LE_UI32.pack(numFat + numDifat))
        # Transation signature number.
        f.write(b'\x00\x00\x00\x00')
        # Mini stream cutoff size.
        f.write(b'\x00\x10\x00\x00')
        ### TODO: First mini FAT sector location.
        f.write(b'\x00\x00\x00\x00')
        ### TODO: Number of mini FAT sectors.
        f.write(b'\x00\x00\x00\x00')
        # First DIFAT sector location. If there are none, set to 0xFFFFFFFE (End
        # of chain).
        f.write(constants.ST_LE_UI32.pack(0 if numDifat else 0xFFFFFFFE))
        # Number of DIFAT sectors.
        f.write(constants.ST_LE_UI32.pack(numDifat))

        # To make life easier on me, I'm having the code start with the DIFAT
        # followed by the FAT sectors, as I can write them all at once before
        # writing the actual contents of the file.

        # Write the DIFAT sectors.
        for x in range(numFat):
            # This kind of sucks to code, ngl.
            if x > 109 and (x - 109) % 127 == 0:
                # If we are at the end of a DIFAT sector, write the jump.
                f.write(constants.ST_LE_UI32.pack((x - 109) // 127))
            # Write the next FAT sector location.
            f.write(constants.ST_LE_UI32.pack(x + numDifat))

        # Finally, fill out the last DIFAT sector with null entries.
        if numFat > 109:
            f.write(b'\xFF\xFF\xFF\xFF' * ((numFat - 109) % 127))
        else:
            f.write(b'\xFF\xFF\xFF\xFF' * (109 - numFat))

        ### FAT.

        # First, if we had any DIFAT sectors, write that the previous sectors
        # were all a part of it.
        f.write(b'\xFC\xFF\xFF\xFF' * numDifat)
        # Second write that the next x sectors are all FAT sectors.
        f.write(b'\xFD\xFF\xFF\xFF' * numFat)

        ### TODO: handle the rest of the actual sectors.
        # TEMP!!!
        f.write(b'\x88\x88\x88\x88' * (totalSectors - numDifat - numFat))

        # Finally, fill fat with markers to specify no block exists.
        freeSectors = totalSectors & 0x7F
        if freeSectors:
            f.write(b'\xFF\xFF\xFF\xFF' * (128 - freeSectors))

    def _getFatSectors(self):
        """
        Returns a tuple containing the number of FAT sectors, the number of
        DIFAT sectors, and the total number of sectors the saved file will have.
        """
        # Right now we just use an annoying while loop to get the numbers.
        numDifat = 0
        # All divisions are ceiling divisions, so we leave them as
        numFat = ceilDiv(self.__numberOfSectors, 127) or 1
        newNumFat = 1
        while numFat != newNumFat:
            numFat = newNumFat
            numDifat = ceilDiv(max(numFat - 109, 0), 127)
            newNumFat = ceilDiv(self.__numberOfSectors + numDifat, 127)

        return (numFat, numDifat, self.__numberOfSectors + numDifat + numFat)


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
            self._writeBeginning(f)

        finally:
            if opened:
                f.close()
