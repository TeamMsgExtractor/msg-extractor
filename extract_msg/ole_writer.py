from __future__ import annotations


__all__ = [
    'DirectoryEntry',
    'OleWriter',
]


import copy
import re

from typing import (
        Dict, Iterator, List, Optional, SupportsBytes, Tuple, TYPE_CHECKING,
        Union
    )

from . import constants
from .constants import MSG_PATH
from .enums import Color, DirectoryEntryType
from .exceptions import StandardViolationError, TooManySectorsError
from .utils import ceilDiv, dictGetCasedKey, inputToMsgPath
from olefile.olefile import OleDirectoryEntry, OleFileIO
from red_black_dict_mod import RedBlackTree


# Allow for nice type checking.
if TYPE_CHECKING:
    from .msg_classes import MSGFile


class DirectoryEntry:
    """
    An internal representation of a stream or storage in the OleWriter.

    Originals should be inaccessible outside of the class.
    """
    name: str = ''
    rightChild: Optional[DirectoryEntry] = None
    leftChild: Optional[DirectoryEntry] = None
    childTreeRoot: Optional[DirectoryEntry] = None
    stateBits: int = 0
    creationTime: int = 0
    modifiedTime: int = 0
    type: DirectoryEntryType = DirectoryEntryType.UNALLOCATED

    # These get set after things have been sorted by the red black tree.
    id: int = -1
    # This is the ID for the left child. The terminology in the docs is really
    # annoying.
    leftSiblingID: int = 0xFFFFFFFF
    rightSiblingID: int = 0xFFFFFFFF
    # This is the ID for the root of the child tree, if any.
    childID: int = 0xFFFFFFFF
    startingSectorLocation: int = 0
    color: Color = Color.BLACK

    clsid: bytes = b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00'
    data: bytes = b''

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        """
        Converts the entry to bytes to be writen to a file.
        """
        # First write the name and the name length.
        if len(self.name) > 31:
            raise ValueError('Name is too long for directory entry.')
        if len(self.name) < 1:
            raise ValueError('Directory entry must have a name.')
        if re.search('/\\\\:!', self.name):
            raise ValueError('Directory entry name contains an illegal character.')

        nameBytes = self.name.encode('utf-16-le')

        return constants.st.ST_CF_DIR_ENTRY.pack(
                                              nameBytes,
                                              len(nameBytes) + 2,
                                              self.type,
                                              self.color,
                                              self.leftSiblingID,
                                              self.rightSiblingID,
                                              self.childID,
                                              self.clsid,
                                              self.stateBits,
                                              self.creationTime,
                                              self.modifiedTime,
                                              self.startingSectorLocation,
                                              getattr(self, 'streamSize', len(self.data)),
                                             )



class OleWriter:
    """
    Takes data to write to a compound binary format file, as specified in
    [MS-CFB].
    """
    def __init__(self, rootClsid: bytes = constants.DEFAULT_CLSID):
        self.__rootEntry = DirectoryEntry()
        self.__rootEntry.name = "Root Entry"
        self.__rootEntry.type = DirectoryEntryType.ROOT_STORAGE
        self.__rootEntry.clsid = rootClsid
        # The root entry will always exist, so this must be at least 1.
        self.__dirEntryCount = 1
        self.__dirEntries = {}
        self.__largeEntries: List[DirectoryEntry] = []
        self.__largeEntrySectors = 0
        self.__numMinifatSectors = 0

        # In a future version, this will be setable as an optional argument.
        self.__version = 3

    def __getContainingStorage(self, path: List[str], entryExists: bool = True, create: bool = False) -> Dict:
        """
        Finds the storage ``dict`` internally where the entry specified by
        :param path: would be created.

        :param entryExists: If ``True``, throws an error when the requested
            entry does not yet exist.
        :param create: If ``True``, creates missing storages with default
            settings.

        :raises OSError: If :param create: is ``False`` and the path could not
            be found. Also raised if :param entryExists: is ``True`` and the
            requested entry does not exist.
        :raises ValueError: Tried to access an interal stream or tried to use
            both the create option and the entryExists option is ``True``.

        :returns: The storage ``dict`` that the entry is in.
        """
        if not path:
            raise OSError('Path cannot be empty.')

        # Quick check for incompatability between create and entryExists.
        if create and entryExists:
            raise ValueError(':param create: and :param entryExists: cannot both be True (an entry cannot exist if it is being created).')

        # Check that the path is not an internal entry. Given the validation on
        # paths that most functions should do because of the call to
        # inputToMsgPath, this shouldn't actually be necessary.
        if any(x.startswith('::') for x in path):
            raise ValueError('Found internal name in path.')

        _dir = self.__dirEntries

        for index, name in enumerate(path[:-1]):
            # If no entry in the current stream matches the path, raise an
            # OSError, *unless* the option to create storages is True.
            if name.lower() not in map(str.lower, _dir.keys()):
                if create:
                    self.addEntry(path[:index + 1], storage = True)
                else:
                    raise OSError(f'Entry not found: {name}')
            _dir = _dir[dictGetCasedKey(_dir, name)]

            # If the current item is not a storage and we have more to the path,
            # raise an OSError.
            if not isinstance(_dir, dict):
                raise OSError('Attempted to access children of a stream.')

        if entryExists and path[-1].lower() not in map(str.lower, _dir.keys()):
            raise OSError(f'Entry not found: {path[-1]}')

        return _dir

    def __getEntry(self, path: List[str]) -> DirectoryEntry:
        """
        Finds and returns an existing ``DirectoryEntry`` instance in the writer.

        :raises OSError: If the entry does not exist.
        :raises ValueError: If access to an internal item is attempted.
        """
        _dir = self.__getContainingStorage(path)
        item = _dir[dictGetCasedKey(_dir, path[-1])]
        if isinstance(item, dict):
            return item['::DirectoryEntry']
        else:
            return item

    def __modifyEntry(self, entry: DirectoryEntry, **kwargs):
        """
        Edits the DirectoryEntry with the data provided.

        Common code used for :meth:`addEntry` and :meth:`editEntry`.

        :raises TypeError: Attempted to modify the data of a storage.
        :raises ValueError: Some part of the data given to modify the various
            properties was invalid. See the the listed methods for details.
        """
        # Extract the arguments.
        data = kwargs.get('data')
        clsid = kwargs.get('clsid')
        creationTime = kwargs.get('creationTime')
        modifiedTime = kwargs.get('modifiedTime')
        stateBits = kwargs.get('stateBits')

        # I don't like that I have repeated if statements for checking each of
        # the arguments, but I need to make sure nothing changes if something is
        # invalid.
        if data is not None:
            if entry.type is not DirectoryEntryType.STREAM:
                raise TypeError('Cannot set the data of a storage object.')
            if not isinstance(data, bytes):
                try:
                    data = bytes(data)
                except Exception:
                    raise ValueError('Data must be a bytes instance or convertable to bytes if set.')
            # Check the length of data. In future versions, this may be a
            # different check which is done when swapping between version 3 and
            # 4 of the compound file binary file format.
            if len(data) > 0x80000000:
                raise ValueError('Current version of extract_msg does not support streams greater than 2 GB in OLE files.')

        if clsid is not None:
            if not isinstance(clsid, bytes):
                raise ValueError('CLSID must be bytes.')
            if len(clsid) != 16:
                raise ValueError('CLSID must be 16 bytes.')

        if creationTime is not None:
            if entry.type is DirectoryEntryType.STREAM:
                raise ValueError('Modification of creation time cannot be done on a stream.')
            if not isinstance(creationTime, int) or creationTime < 0 or creationTime > 0xFFFFFFFFFFFFFFFF:
                raise ValueError('Creation time must be a positive 8 byte int.')

        if modifiedTime is not None:
            if entry.type is DirectoryEntryType.STREAM:
                raise ValueError('Modification of modified time cannot be done on a stream.')
            if not isinstance(modifiedTime, int) or modifiedTime < 0 or modifiedTime > 0xFFFFFFFFFFFFFFFF:
                raise ValueError('Modified time must be a positive 8 byte int.')

        if stateBits is not None:
            if not isinstance(stateBits, int) or stateBits < 0 or stateBits > 0xFFFFFFFF:
                raise ValueError('State bits must be a positive 4 byte int.')

        # Now that all our checks have passed, let's set our data.
        if data is not None:
            entry.data = data
        if clsid is not None:
            entry.clsid = clsid
        if creationTime is not None:
            entry.creationTime = creationTime
        if modifiedTime is not None:
            entry.modifiedTime = modifiedTime
        if stateBits is not None:
            entry.stateBits = stateBits

    def __recalculateSectors(self) -> None:
        """
        Recalculates several of the internal variables used for saving that
        specify the number of sectors and where things should go.
        """
        self.__dirEntryCount = 0
        self.__numMinifatSectors = 0
        self.__largeEntries.clear()
        self.__largeEntrySectors = 0

        for entry in self.__walkEntries():
            self.__dirEntryCount += 1
            if entry.type == DirectoryEntryType.STREAM:
                if len(entry.data) < 4096:
                    self.__numMinifatSectors += ceilDiv(len(entry.data), 64)
                else:
                    self.__largeEntries.append(entry)
                    self.__largeEntrySectors += ceilDiv(len(entry.data), self.__sectorSize)

    def __walkEntries(self) -> Iterator[DirectoryEntry]:
        """
        Returns a generator that will walk the entires recursively.

        Each item returned by it will be a DirectoryEntry instance.
        """
        toProcess = [self.__dirEntries]
        yield self.__rootEntry

        while len(toProcess) > 0:
            for name, item in toProcess.pop(0).items():
                if not name.startswith('::'):
                    if isinstance(item, dict):
                        yield item['::DirectoryEntry']
                        toProcess.append(item)
                    else:
                        yield item

    @property
    def __dirEntsPerSector(self) -> int:
        """
        The number of Directory Entries that can fit in a sector.
        """
        return self.__sectorSize // 128

    @property
    def __linksPerSector(self) -> int:
        """
        The number of links per FAT/DIFAT sector.
        """
        return self.__sectorSize // 4

    @property
    def __miniSectorsPerSector(self) -> int:
        """
        The number of mini sectors that a regular sector will hold.
        """
        return self.__sectorSize // 64

    @property
    def __numberOfSectors(self) -> int:
        # Most of this should be pretty self evident, but line by line the
        # calculation is as such:
        # 1. How many sectors are needed for the directory entries.
        # 2. How many FAT sectors are needed for the MiniStream.
        # 3. How many sectors are needed for the MiniFat (ceil divide #2 by 16).
        # 4. The number of FAT sectors needed to store the larger data.
        return ceilDiv(self.__dirEntryCount, 4) + \
               self.__numMinifat + \
               ceilDiv(self.__numMinifat, 16) + \
               self.__largeEntrySectors

    @property
    def __numMinifat(self) -> int:
        """
        The number of FAT sectors needed to store the mini stream.
        """
        return ceilDiv(64 * self.__numMinifatSectors, self.__sectorSize)

    @property
    def __sectorSize(self) -> int:
        """
        The size of each sector, in bytes.
        """
        return 512 if self.__version == 3 else 4096

    def _cleanupEntries(self) -> None:
        """
        Cleans up the node connections by walking the tree and removing
        references that were added during writing.
        """
        self.__largeEntries.clear()
        for entry in self.__walkEntries():
            entry.id = -1
            entry.leftChild = None
            entry.rightChild = None
            entry.childTreeRoot = None
            entry.leftSiblingID = 0xFFFFFFFF
            entry.rightSiblingID = 0xFFFFFFFF
            entry.childID = 0xFFFFFFFF

    def _getFatSectors(self) -> Tuple[int, int, int]:
        """
        Returns a tuple containing the number of FAT sectors, the number of
        DIFAT sectors, and the total number of sectors the saved file will have.
        """
        # Right now we just use an annoying while loop to get the numbers.
        numDifat = 0
        # All divisions are ceiling divisions,.
        numFat = ceilDiv(self.__numberOfSectors or 1, self.__linksPerSector - 1)
        newNumFat = 1
        while numFat != newNumFat:
            numFat = newNumFat
            numDifat = ceilDiv(max(numFat - 109, 0), self.__linksPerSector - 1)
            newNumFat = ceilDiv(self.__numberOfSectors + numDifat, self.__linksPerSector - 1)

        return (numFat, numDifat, self.__numberOfSectors + numDifat + numFat)

    def _treeSort(self, startingSector: int) -> List[DirectoryEntry]:
        """
        Uses red-black trees to sort the internal data in preparation for
        writing the file, returning a list, in order, of the entries to write.
        """
        # First, create the root entry.
        root = copy.copy(self.__rootEntry)

        # Add the location of the start of the mini stream.
        root.startingSectorLocation = (startingSector + ceilDiv(self.__dirEntryCount, 4) + ceilDiv(self.__numMinifatSectors, self.__linksPerSector)) if self.__numMinifat > 0 else 0xFFFFFFFE
        root.streamSize = self.__numMinifatSectors * 64
        root.childTreeRoot = None
        root.childID = 0xFFFFFFFF
        entries = [root]

        toProcess = [(root, self.__dirEntries)]
        # Continue looping while there is more to process.
        while toProcess:
            entry, currentItem = toProcess.pop()
            if not currentItem:
                continue
            # If the current item *only* has the directory's entry and no stream
            # entries, we are actually done.
            # Create a tree and add all the items to it. We add it with a key
            # that is a tuple of the length (as shorter is *always* less than
            # longer) and the uppercase name, and the value is the actual entry.
            tree = RedBlackTree()
            for name in currentItem:
                if not name.startswith('::'):
                    val = currentItem[name]
                    # If we find a directory entry, then we need to add it to
                    # the processing list.
                    if isinstance(val, dict):
                        toProcess.append((val['::DirectoryEntry'], val))
                        val = val['::DirectoryEntry']

                    entries.append(val)

                    # Add the data to the tree.
                    tree.add((len(name), name.upper()), val)

            # Now that everything is added, we need to take our root and add it
            # as the child of the current entry.
            entry.childTreeRoot = tree.value

            # Now we need to go through each node and set it's data based on
            # it's sort position.
            for node in tree.in_order():
                item = node.value
                # Set the color immediately.
                item.color = Color.BLACK if node.is_black else Color.RED

                if node.left:
                    item.leftChild = node.left.value
                else:
                    item.leftChild = None

                if node.right:
                    item.rightChild = node.right.value
                else:
                    item.rightChild = None

        # Now that everything is connected, we loop over the entries list a few
        # times and set the data values.
        for _id, entry in enumerate(entries):
            entry.id = _id

        for entry in entries:
            entry.leftSiblingID = entry.leftChild.id if entry.leftChild else 0xFFFFFFFF
            entry.childID = entry.childTreeRoot.id if entry.childTreeRoot else 0xFFFFFFFF
            entry.rightSiblingID = entry.rightChild.id if entry.rightChild else 0xFFFFFFFF

        # Finally, let's figure out the sector IDs to be used for the mini data.
        # We only need to do this for streams with a size less than 4096.

        # Use this to track where the next thing goes in the mini FAT.
        miniFATLocation = 0

        for entry in entries:
            if len(entry.data) == 0 and entry != entries[0]:
                # If there is no data, just set the starting location to none.
                entry.startingSectorLocation = 0xFFFFFFFE
            elif entry.type == DirectoryEntryType.STREAM and len(entry.data) < 4096:
                entry.startingSectorLocation = miniFATLocation
                miniFATLocation += ceilDiv(len(entry.data), 64)

        return entries

    def _writeBeginning(self, f) -> int:
        """
        Writes the beginning to the file :param f:.

        This includes the header, DIFAT, and FAT blocks.

        :returns: The current sector number after all the data is written.

        :raises TooMuchDataError: The number of sectors required for the file is
            too large.
        """
        # Recalculate some things needed for saving.
        self.__recalculateSectors()
        # Since we are going to need these multiple times, get them now.
        numFat, numDifat, totalSectors = self._getFatSectors()

        # Check to make sure there isn't too much data to write.
        if totalSectors > 0xFFFFFFFB:
            raise TooManySectorsError('Data in OleWriter requires too many sectors to write to a version 3 file.')

        # The ministream *cannot* be greater than 2 GB, so check that before
        # writing anything. A minifat sector is 64 bytes, so the maximum amount
        # of them is 0x2000000.
        if self.__numMinifatSectors > 0x2000000:
            raise TooManySectorsError('Data is OleWriter requires too many MiniFAT sectors.')

        # Header signature.
        f.write(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1')
        # Header CLSID.
        f.write(b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00')
        # Minor version.
        f.write(b'\x3E\x00')
        # Major version. For now, we only support version 3, but support for
        # version 4 is planned.
        f.write(b'\x03\x00' if self.__version == 3 else b'\x04\x00')
        # Byte order. Specifies that it is little endian.
        f.write(b'\xFE\xFF')
        # Sector shift.
        f.write(b'\x09\x00' if self.__version == 3 else b'\x0C\x00')
        # Mini sector shift.
        f.write(b'\x06\x00')
        # Reserved.
        f.write(b'\x00\x00\x00\x00\x00\x00')
        # Number of directory sectors. Version 3 says this *must* be 0.
        f.write(constants.st.ST_LE_UI32.pack(0))
        # Number of FAT sectors.
        f.write(constants.st.ST_LE_UI32.pack(numFat))
        # First directory sector location (Sector for the directory stream).
        # We place that right after the DIFAT and FAT.
        f.write(constants.st.ST_LE_UI32.pack(numFat + numDifat))
        # Transation signature number.
        f.write(b'\x00\x00\x00\x00')
        # Mini stream cutoff size.
        f.write(b'\x00\x10\x00\x00')
        # First mini FAT sector location.
        f.write(constants.st.ST_LE_UI32.pack((numFat + numDifat + ceilDiv(self.__dirEntryCount, 4)) if self.__numMinifat > 0 else 0xFFFFFFFE))
        # Number of mini FAT sectors.
        f.write(constants.st.ST_LE_UI32.pack(ceilDiv(self.__numMinifatSectors, self.__linksPerSector)))
        # First DIFAT sector location. If there are none, set to 0xFFFFFFFE (End
        # of chain).
        f.write(constants.st.ST_LE_UI32.pack(0 if numDifat else 0xFFFFFFFE))
        # Number of DIFAT sectors.
        f.write(constants.st.ST_LE_UI32.pack(numDifat))

        # To make life easier on me, I'm having the code start with the DIFAT
        # followed by the FAT sectors, as I can write them all at once before
        # writing the actual contents of the file.

        # Write the DIFAT sectors.
        for x in range(numFat):
            # Quickly check if we have hit 109. If we have, and we are writing
            # a version 4 file, we need to pad a bunch of null bytes.
            if x == 109 and self.__version == 4:
                f.write(b'\x00' * 3584)
            # This kind of sucks to code, ngl.
            if x > 109 and (x - 109) % (self.__linksPerSector - 1) == 0:
                # If we are at the end of a DIFAT sector, write the jump.
                f.write(constants.st.ST_LE_UI32.pack((x - 109) // (self.__linksPerSector - 1)))
            # Write the next FAT sector location.
            f.write(constants.st.ST_LE_UI32.pack(x + numDifat))

        # Finally, fill out the last DIFAT sector with null entries.
        if numFat > 109:
            f.write(b'\xFF\xFF\xFF\xFF' * ((self.__linksPerSector - 1) - ((numFat - 109) % (self.__linksPerSector - 1))))
            # Finally, make sure to write the end of chain marker for the DIFAT.
            f.write(b'\xFE\xFF\xFF\xFF')
        else:
            f.write(b'\xFF\xFF\xFF\xFF' * (109 - numFat))

        ### FAT.

        # First, if we had any DIFAT sectors, write that the previous sectors
        # were all a part of it.
        f.write(b'\xFC\xFF\xFF\xFF' * numDifat)
        # Second write that the next x sectors are all FAT sectors.
        f.write(b'\xFD\xFF\xFF\xFF' * numFat)

        offset = numDifat + numFat

        # Fill in the values for the directory stream.
        for x in range(offset + 1, offset + ceilDiv(self.__dirEntryCount, self.__dirEntsPerSector)):
            f.write(constants.st.ST_LE_UI32.pack(x))

        # Write the end of chain marker.
        f.write(b'\xFE\xFF\xFF\xFF')

        offset += ceilDiv(self.__dirEntryCount, self.__dirEntsPerSector)

        # Check if we have minifat *at all* first.
        if self.__numMinifatSectors > 0:
            # Mini FAT chain.
            for x in range(offset + 1, offset + ceilDiv(self.__numMinifat, 16)):
                f.write(constants.st.ST_LE_UI32.pack(x))

            # Write the end of chain marker.
            f.write(b'\xFE\xFF\xFF\xFF')

            offset += ceilDiv(self.__numMinifat, 16)

            # The mini stream sectors.
            for x in range(offset + 1, offset + self.__numMinifat):
                f.write(constants.st.ST_LE_UI32.pack(x))

            # Write the end of chain marker.
            f.write(b'\xFE\xFF\xFF\xFF')

            offset += self.__numMinifat

        # Regular stream chains. These are the most complex to handle. We handle
        # them by checking a list that was make of entries which were only added
        # to that list if the size was more than 4096. The order in the list is
        # how they will eventually be stored into the file correctly.
        for entry in self.__largeEntries:
            size = ceilDiv(len(entry.data), self.__sectorSize)
            entry.startingSectorLocation = offset
            for x in range(offset + 1, offset + size):
                f.write(constants.st.ST_LE_UI32.pack(x))

            # Write the end of chain marker.
            f.write(b'\xFE\xFF\xFF\xFF')

            offset += size

        # Finally, fill fat with markers to specify no block exists.
        freeSectors = totalSectors & (self.__linksPerSector - 1)
        if freeSectors:
            f.write(b'\xFF\xFF\xFF\xFF' * (self.__linksPerSector - freeSectors))

        # Finally, return the current sector index for use in other places.
        return numDifat + numFat

    def _writeDirectoryEntries(self, f, startingSector: int) -> List[DirectoryEntry]:
        """
        Writes out all the directory entries. Returns the list generated.
        """
        entries = self._treeSort(startingSector)
        for x in entries:
            self._writeDirectoryEntry(f, x)
        if len(entries) & 3:
            f.write(((b'\x00\x00' * 34) + (b'\xFF\xFF' * 6) + (b'\x00\x00' * 24)) * (4 - (len(entries) & 3)))

        return entries

    def _writeDirectoryEntry(self, f, entry: DirectoryEntry) -> None:
        """
        Writes the directory entry to the file f.
        """
        f.write(bytes(entry))

    def _writeFinal(self, f) -> None:
        """
        Writes the final sectors of the file, consisting of the streams too
        large for the mini FAT.
        """
        for x in self.__largeEntries:
            f.write(x.data)
            if len(x.data) & (self.__sectorSize - 1):
                f.write(b'\x00' * (self.__sectorSize - (len(x.data) & (self.__sectorSize - 1))))

    def _writeMini(self, f, entries: List[DirectoryEntry]) -> None:
        """
        Writes the mini FAT followed by the full mini stream.
        """
        # For each of the entires that are streams and less than 4096.
        currentSector = 0
        for x in entries:
            if x.type == DirectoryEntryType.STREAM and len(x.data) < 4096:
                size = ceilDiv(len(x.data), 64)
                for x in range(currentSector + 1, currentSector + size):
                    f.write(constants.st.ST_LE_UI32.pack(x))
                if size > 0:
                    f.write(b'\xFE\xFF\xFF\xFF')
                currentSector += size

        # Finally, write the remaining slots.
        if currentSector & (self.__linksPerSector - 1):
            f.write(b'\xFF\xFF\xFF\xFF' * (self.__linksPerSector - (currentSector & (self.__linksPerSector - 1))))

        # Write the mini stream.
        for x in entries:
            if len(x.data) > 0 and len(x.data) < 4096:
                f.write(x.data)
                if len(x.data) & 63:
                    f.write(b'\x00' * (64 - (len(x.data) & 63)))

        # Pad the final mini stream block.
        if self.__numMinifatSectors & (self.__miniSectorsPerSector - 1):
            f.write((b'\x00' * 64) * (self.__miniSectorsPerSector - (self.__numMinifatSectors & (self.__miniSectorsPerSector - 1))))

    def addEntry(self, path: MSG_PATH, data: Optional[Union[bytes, SupportsBytes]] = None, storage: bool = False, **kwargs) -> None:
        """
        Adds an entry to the OleWriter instance at the path specified, adding
        storages with default settings where necessary. If the entry is not a
        storage, :param data: *must* be set.

        :param path: The path to add the entry at. Must not contain a path part
            that is an already added stream.
        :param data: The bytes for a stream or an object with a ``__bytes__``
            method.
        :param storage: If ``True``, the entry to add is a storage. Otherwise,
            the entry is a stream.
        :param clsid: The CLSID for the stream/storage. Must a a bytes instance
            that is 16 bytes long.
        :param creationTime: An 8 byte filetime int. Sets the creation time of
            the entry. Not applicable to streams.
        :param modifiedTime: An 8 byte filetime int. Sets the modification time
            of the entry. Not applicable to streams.
        :param stateBits: A 4 byte int. Sets the state bits, user-defined flags,
            of the entry. For a stream, this *SHOULD* be unset.

        :raises OSError: A stream was found on the path before the end or an
            entry with the same name already exists.
        :raises ValueError: Attempts to access an internal item.
        :raises ValueError: The data provided is too large.
        """
        path = inputToMsgPath(path)
        # First, find the current place in our dict to add the item.
        _dir = self.__getContainingStorage(path, False, True)
        # Now, check that the item *is not* already in our dict, as that would
        # cause problems.
        if path[-1].lower() in map(str.lower, _dir.keys()):
            raise OSError('Cannot add an entry that already exists.')

        # Create a new entry with basic data and insert it.
        entry = DirectoryEntry()
        entry.type = DirectoryEntryType.STORAGE if storage else DirectoryEntryType.STREAM
        entry.name = path[-1]
        self.__modifyEntry(entry, data = data, **kwargs)
        if storage:
            _dir[path[-1]] = {'::DirectoryEntry': entry}
        else:
            _dir[path[-1]] = entry

    def addOleEntry(self, path: MSG_PATH, entry: OleDirectoryEntry, data: Optional[Union[bytes, SupportsBytes]] = None) -> None:
        """
        Uses the entry provided to add the data to the writer.

        :raises OSError: Tried to add an entry to a path that has not yet
            been added, tried to add as a child of a stream, or tried to add an
            entry where one already exists under the same name.
        :raises ValueError: The data provided is too large.
        """
        path = inputToMsgPath(path)
        # First, find the current place in our dict to add the item.
        _dir = self.__getContainingStorage(path, False)
        # Now, check that the item *is not* already in our dict, as that would
        # cause problems.
        if path[-1].lower() in map(str.lower, _dir.keys()):
            raise OSError('Cannot add an entry that already exists.')

        # Now that we are in the right place, add our data.
        newEntry = DirectoryEntry()
        if entry.entry_type == DirectoryEntryType.STORAGE:
            # Handle a storage entry.
            # First, setup the values for the storage.
            newEntry.name = entry.name
            newEntry.type = DirectoryEntryType.STORAGE
            newEntry.clsid = _unClsid(entry.clsid)
            newEntry.stateBits = entry.dwUserFlags
            newEntry.creationTime = entry.createTime
            newEntry.modifiedTime = entry.modifyTime

            # Finally add the dict to our tree of items.
            _dir[path[-1]] = {'::DirectoryEntry': newEntry}
        else:
            # Handle a stream entry.
            # First, setup the values for the stream.
            newEntry.name = entry.name
            newEntry.type = DirectoryEntryType.STREAM
            newEntry.clsid = _unClsid(entry.clsid)
            newEntry.stateBits = entry.dwUserFlags

            # Next, handle the data.
            data = data or b''
            newEntry.data = bytes(data)
            if len(newEntry.data) > 0x80000000:
                raise ValueError('Current version of extract_msg does not support streams greater than 2 GB in OLE files.')

            # Finally add the entry to out dict of entries.
            _dir[path[-1]] = newEntry

        self.__dirEntryCount += 1

    def deleteEntry(self, path) -> None:
        """
        Deletes the entry specified by :param path:, including all children.

        :raises OSError: If the entry does not exist or a part of the path that
            is not the last was a stream.
        :raises ValueError: Attempted to delete an internal data stream.
        """
        path = inputToMsgPath(path)
        # Get the containing storage for the entry.
        _dir = self.__getContainingStorage(path)

        # The garbage collector will take care of all the loose items, so just
        # remove the entry. Also, once again we deal with the case insensitive
        # nature of the path. Even though comparisons are case insensitive, the
        # path does remember the case used.
        del _dir[dictGetCasedKey(_dir, path[-1])]

    def editEntry(self, path: MSG_PATH, **kwargs) -> None:
        """
        Used to edit values of an entry by setting the specific kwargs. Set a
        value to something other than None to set it.

        :param data: The data of a stream. Will error if used for something
            other than a stream. Must be bytes or convertable to bytes.
        :param clsid: The CLSID for the stream/storage. Must a a bytes instance
            that is 16 bytes long.
        :param creationTime: An 8 byte filetime int. Sets the creation time of
            the entry. Not applicable to streams.
        :param modifiedTime: An 8 byte filetime int. Sets the modification time
            of the entry. Not applicable to streams.
        :param stateBits: A 4 byte int. Sets the state bits, user-defined flags,
            of the entry. For a stream, this *SHOULD* be unset.

        To convert a 32 character hexadecial CLSID into the bytes for this
        function, the _unClsid function in the ole_writer submodule can be used.

        :raises OSError: The entry does not exist in the file.
        :raises TypeError: Attempted to modify the bytes of a storage.
        :raises ValueError: The type of a parameter was wrong, or the data of a
            parameter was invalid.
        """
        # First, find our entry to edit.
        entry = self.__getEntry(inputToMsgPath(path))

        # Send it to be modified using the arguments given.
        self.__modifyEntry(entry, **kwargs)

    def fromMsg(self, msg: MSGFile, allowBadEmbed: bool = False) -> None:
        """
        Copies the streams and stream information necessary from the MSG file.

        :param allowBadEmbed: If True, attempts to skip steps that will fail if 
            the embedded MSG file violates standards. It will also attempt to repair the data to try to ensure it can open in Outlook.

        :raises StandardViolationError: Something about the embedded data has a
            fundemental issue that violates the standard.
        """
        # Get the root OLE entry's CLSID.
        self.__rootEntry.clsid = _unClsid(msg._getOleEntry('/').clsid)

        # List both storages and directories, but sort them by shortest length
        # first to prevent errors.
        entries = msg.listDir(True, True, False)
        entries.sort(key = len)

        for x in entries:
            entry = msg._getOleEntry(x)
            data = msg.getStream(x) if entry.entry_type == DirectoryEntryType.STREAM else None
            # THe properties stream on embedded messages actualy needs to be
            # transformed a little (*why* it is like that is a mystery to me).
            # Basically we just need to add a "reserved" section to it in a
            # specific place. So let's check if we are doing the properties
            # stream and then if we are embedded.
            if x[0] == '__properties_version1.0' and msg.prefixLen > 0:
                if len(data) % 16 != 0:
                    data = data[:24] + b'\x00\x00\x00\x00\x00\x00\x00\x00' + data[24:]
                elif not allowBadEmbed:
                    # If we are not allowing bad data, throw an error.
                    raise StandardViolationError('Embedded msg file attempted to be extracted that contains a top level properties stream.')
                if allowBadEmbed:
                    # See if we need to fix the properties stream at all.
                    if msg.getPropertyVal('340D0003') is None:
                        if msg.areStringsUnicode:
                            # We need to add a property to allow this file to open:
                                data += b'\x03\x00\x0D\x34\x02\x00\x00\x00\x00\x00\x04\x00\x00\x00\x00\x00'
            self.addOleEntry(x, entry, data)

        # Now check if it is an embedded file. If so, we need to copy the named
        # properties streams (the metadata, not the values).
        if msg.prefixLen > 0:
            # Get the entry for the named properties directory and add it
            # immediately if it exists. If it doesn't exist, this whole
            # section will be skipped.
            try:
                self.addOleEntry('__nameid_version1.0', msg._getOleEntry('__nameid_version1.0', False), None)
            except OSError as e:
                if str(e).startswith('Cannot add an entry'):
                    if allowBadEmbed:
                        return
                    raise StandardViolationError('Embedded msg file attempted to be extracted that contains it\'s own named streams.')
                raise

            # Now that we know it exists, grab all the file inside and copy
            # them to our root.
            # Create our generator.
            gen = (x for x in msg._oleListDir() if len(x) > 1 and x[0] == '__nameid_version1.0')
            for x in gen:
                self.addOleEntry(x, msg._getOleEntry(x, prefix = False), msg.getStream(x, prefix = False))

    def fromOleFile(self, ole: OleFileIO, rootPath: MSG_PATH = []) -> None:
        """
        Copies all the streams from the proided OLE file into this writer.

        NOTE: This method does *not* handle any special rule that may be
        required by a format that uses the compound binary file format as a base
        when extracting an embedded directory. For example, MSG files require
        modification of an embedded properties stream when extracting an
        embedded MSG file.

        :param rootPath: A path (accepted by ``olefile.OleFileIO``) to the
            directory to use as the root of the file. If not provided, the file
            root will be used.

        :raises OSError: If :param rootPath: does not exist in the file.
        """
        rootPath = inputToMsgPath(rootPath)

        # Check if the root path is simply the top of the file.
        if rootPath == []:
            # Copy the clsid of the root entry.
            self.__rootEntry.clsid = _unClsid(ole.direntries[0].clsid)
            paths = {tuple(x): (x, ole.direntries[ole._find(x)]) for x in ole.listdir(True, True)}
        else:
            # If it is not the top of the file, we need to do some filtering.
            # First get the CLSID from the entry the path points to.
            try:
                entry = ole.direntries[ole._find(rootPath)]
                self.__rootEntry.clsid = _unClsid(entry.clsid)

            except OSError as e:
                if str(e) == 'file not found':
                    # Get the cause/context for the original exception and use
                    # it for the new exception. This hides the exception from
                    # OleFileIO.
                    context = e.__cause__ or e.__context__
                    raise OSError('Root path was not found in the OLE file.') from context
                else:
                    raise

            paths = {tuple(x[len(rootPath):]): (x, ole.direntries[ole._find(x)])
                     for x in ole.listdir(True, True) if len(x) > len(rootPath)}


        # Copy all of the other entries. Ensure that directories come before
        # their streams by sorting the paths.
        for x in sorted(paths.keys()):
            fullPath, entry = paths[x]

            if entry.entry_type == DirectoryEntryType.STREAM:
                with ole.openstream(fullPath) as f:
                    data = f.read()
            else:
                data = None

            self.addOleEntry(x, entry, data)

    def getEntry(self, path: MSG_PATH) -> DirectoryEntry:
        """
        Finds and returns a copy of an existing `DirectoryEntry` instance in the
        writer. Use this method to check the internal status of an entry.

        :raises OSError: If the entry does not exist.
        :raises ValueError: If access to an internal item is attempted.
        """
        return copy.copy(self.__getEntry(inputToMsgPath(path)))

    def listItems(self, streams: bool = True, storages: bool = False) -> List[List[str]]:
        """
        Returns a list of the specified items currently in the writter.

        :param streams: If ``True``, includes the path for each stream in the
            list.
        :param storages: If ``True``, includes the path for each storage in the
            list.
        """
        # We are actually abusing the walk function a bit here to life much
        # easier. The way we do this is to look at the current directory that
        # the walk function is giving information about and then deciding what
        # parts of it we want to use. Once we have all the paths created, we
        # will then sort and return it to give an output similar, if not
        # identical, to OleFileIO.listdir. The mentioned method sorts keeping
        # case in mind.
        if not streams and not storages:
            return []

        paths = []
        for currentDir, stor, stre in self.walk():
            if storages:
                for name in stor:
                    paths.append(currentDir + [name])
            if streams:
                for name in stre:
                    paths.append(currentDir + [name])

        paths.sort()
        return paths

    def renameEntry(self, path: MSG_PATH, newName: str) -> None:
        """
        Changes the name of an entry, leaving it in it's current position.

        :raises OSError: If the entry does not exist or an entry with the new
            name already exists,
        :raises ValueError: If access to an internal item is attempted or the
            new name provided is invalid.
        """
        # First, validate the new name.
        if not isinstance(newName, str):
            raise ValueError('New name must be a string.')
        if constants.re.INVALID_OLE_PATH.search(newName):
            raise ValueError('Invalid character(s) in new name. Must not contain the following characters: \\//!:')
        if len(newName) > 31:
            raise ValueError('New name must be less than 32 characters.')

        # Get the storage for our entry. Entry *must* exist.
        _dir = self.__getContainingStorage(inputToMsgPath(path))

        # See if an item in the storage already has that new name.
        if newName.lower() in map(str.lower, _dir.keys()):
            raise OSError('An entry with the new name already exists.')

        # Get the original name.
        originalName = dictGetCasedKey(_dir, path[-1])

        # Get the entry to change.
        entry = _dir[originalName]
        if isinstance(entry, dict):
            dirData = entry
            entry = entry['::DirectoryEntry']
        else:
            dirData = None

        # Change the name on the entry first.
        entry.name = newName

        # Now, we need to remove the item from the current storage and add it
        # back with the new name.
        del _dir[originalName]

        if dirData is None:
            _dir[newName] = entry
        else:
            _dir[newName] = dirData

    def walk(self) -> Iterator[Tuple[List[str], List[str], List[str]]]:
        """
        Functional equivelent to ``os.walk``, but for going over the file
        structure of the OLE file to be written. Unlike ``os.walk``, it takes
        no arguments.

        :returns: A tuple of three lists. The first is the path, as a list of
            strings, for the directory (or an empty list for the root), the
            second is a list of the storages in the current directory, and the
            last is a list of the streams. Streams and storages are sorted
            caselessly.
        """
        toProcess = [([], self.__dirEntries)]

        # Go through the toProcess list, removing the last item every time to
        # mimic the behavior of os.walk.
        while toProcess:
            currentDir, dirDict = toProcess.pop()
            storages = []
            streams = []
            for name in sorted(dirDict.keys(), key = str.lower):
                if not name.startswith('::'):
                    if isinstance(dirDict[name], dict):
                        storages.append(name)
                        toProcess.append((currentDir + [name], dirDict[name]))
                    else:
                        streams.append(name)

            yield (currentDir, storages, streams)

    def write(self, path) -> None:
        """
        Writes the data to the path specified.

        If :param path: has a ``write`` method, the object will be used
        directly.

        If a failure occurs, the file or IO device may have been modified.

        :raises TooManySectorsError: The number of sectors requires for a part
            of writing is too large.
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
            # Write each section, transferring data between functions where
            # necessary.
            offset = self._writeBeginning(f)
            entries = self._writeDirectoryEntries(f, offset)
            self._writeMini(f, entries)
            self._writeFinal(f)
        finally:
            self._cleanupEntries()

            if opened:
                f.close()



def _unClsid(clsid: str) -> bytes:
    """
    Converts the clsid from ``olefile.olefile._clsid`` back to bytes.
    """
    if not clsid:
        return b''
    clsid = clsid.replace('-', '')
    try:
        return bytes((
            int(clsid[6:8], 16),
            int(clsid[4:6], 16),
            int(clsid[2:4], 16),
            int(clsid[0:2], 16),
            int(clsid[10:12], 16),
            int(clsid[8:10], 16),
            int(clsid[14:16], 16),
            int(clsid[12:14], 16),
            int(clsid[16:18], 16),
            int(clsid[18:20], 16),
            int(clsid[20:22], 16),
            int(clsid[22:24], 16),
            int(clsid[24:26], 16),
            int(clsid[26:28], 16),
            int(clsid[28:30], 16),
            int(clsid[30:32], 16),
        ))
    except Exception:
        raise
