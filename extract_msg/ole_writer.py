import io
import pathlib
import re

from typing import List, Optional, Union

from . import constants
from .enums import Color, DirectoryEntryType
from .utils import ceilDiv, inputToMsgPath
from olefile.olefile import OleDirectoryEntry, OleFileIO
from red_black_dict_mod import RedBlackTree


class _DirectoryEntry:
    """
    Hidden class, will probably be modified later.
    """
    name : str = ''
    rightChild : '_DirectoryEntry' = None
    leftChild : '_DirectoryEntry' = None
    childTreeRoot : '_DirectoryEntry' = None
    stateBits : int = 0
    creationTime : int = 0
    modifiedTime : int = 0
    type : DirectoryEntryType = DirectoryEntryType.UNALLOCATED

    # These get set after things have been sorted by the red black tree.
    id : int = -1
    # This is the ID for the left child. The terminology in the docs is really
    # annoying.
    leftSiblingID : int = 0xFFFFFFFF
    rightSiblingID : int = 0xFFFFFFFF
    # This is the ID for the root of the child tree, if any.
    childID : int = 0xFFFFFFFF
    startingSectorLocation : int = 0
    color : Color = Color.BLACK

    clsid : bytes = b''
    data : bytes = b''

    def __init__(self):
        pass

    def toBytes(self):
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

        return constants.ST_CF_DIR_ENTRY.pack(
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
    def __init__(self, rootClsid : bytes = constants.DEFAULT_CLSID):
        self.__rootClsid = rootClsid
        # The root entry will always exist, so this must be at least 1.
        self.__dirEntryCount = 1
        self.__dirEntries = {}
        self.__largeEntries = []
        self.__largeEntrySectors = 0
        self.__numMinifatSectors = 0

    @property
    def __numberOfSectors(self) -> int:
        """
        TODO: finish the calculation needed. For now this just notes how many
        sectors are needed for the directory entries.
        """
        return ceilDiv(self.__dirEntryCount, 4) + \
               self.__numMinifat + \
               ceilDiv(self.__numMinifat, 16) + \
               self.__largeEntrySectors

    @property
    def __numMinifat(self) -> int:
        return ceilDiv(self.__numMinifatSectors, 8)

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

    def _treeSort(self, startingSector : int) -> List[_DirectoryEntry]:
        """
        Uses red-black trees to sort the internal data in preparation for
        writing the file, returning a list, in order, of the entries to write.
        """
        # First, create the root entry.
        root = _DirectoryEntry()
        root.name = "Root Entry"
        root.type = DirectoryEntryType.ROOT_STORAGE
        root.clsid = self.__rootClsid
        # Add the location of the start of the mini stream.
        root.startingSectorLocation = (startingSector + ceilDiv(self.__dirEntryCount, 4) + ceilDiv(self.__numMinifatSectors, 128)) if self.__numMinifat > 0 else 0xFFFFFFFE
        root.streamSize = self.__numMinifatSectors * 64
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
                        entries.append(val['::DirectoryEntry'])
                    else:
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
                if isinstance(item, dict):
                    item = item['::DirectoryEntry']
                item.color = Color.BLACK if node.is_black else Color.RED
                if node.left:
                    val = node.left.value
                    if isinstance(val, _DirectoryEntry):
                        item.leftChild = val
                    else:
                        item.leftChild = val['::DirectoryEntry']
                else:
                    item.leftChild = None
                if node.right:
                    val = node.right.value
                    if isinstance(val, _DirectoryEntry):
                        item.rightChild = val
                    else:
                        item.rightChild = val['::DirectoryEntry']
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
        Writes the beginning to the file :param f:. This includes the header,
        DIFAT, and FAT blocks.

        :returns: The current sector number after all the data is written.
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
        # First mini FAT sector location.
        f.write(constants.ST_LE_UI32.pack((numFat + numDifat + ceilDiv(self.__dirEntryCount, 4)) if self.__numMinifat > 0 else 0xFFFFFFFE))
        # Number of mini FAT sectors.
        f.write(constants.ST_LE_UI32.pack(ceilDiv(self.__numMinifatSectors, 128)))
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
        for x in range(offset + 1, offset + ceilDiv(self.__dirEntryCount, 4)):
            f.write(constants.ST_LE_UI32.pack(x))

        # Write the end of chain marker.
        f.write(b'\xFE\xFF\xFF\xFF')

        offset += ceilDiv(self.__dirEntryCount, 4)

        # Check if we have minifat *at all* first.
        if self.__numMinifatSectors > 0:
            # Mini FAT chain.
            for x in range(offset + 1, offset + ceilDiv(self.__numMinifat, 16)):
                f.write(constants.ST_LE_UI32.pack(x))

            # Write the end of chain marker.
            f.write(b'\xFE\xFF\xFF\xFF')

            offset += ceilDiv(self.__numMinifat, 16)

            # The mini stream sectors.
            for x in range(offset + 1, offset + self.__numMinifat):
                f.write(constants.ST_LE_UI32.pack(x))

            # Write the end of chain marker.
            f.write(b'\xFE\xFF\xFF\xFF')

            offset += self.__numMinifat

        # Regular stream chains. These are the most complex to handle. We handle
        # them by checking a list that was make of entries which were only added
        # to that list if the size was more than 4096. The order in the list is
        # how they will eventually be stored into the file correctly.
        for entry in self.__largeEntries:
            size = ceilDiv(len(entry.data), 512)
            entry.startingSectorLocation = offset
            for x in range(offset + 1, offset + size):
                f.write(constants.ST_LE_UI32.pack(x))

            # Write the end of chain marker.
            f.write(b'\xFE\xFF\xFF\xFF')

            offset += size

        # Finally, fill fat with markers to specify no block exists.
        freeSectors = totalSectors & 0x7F
        if freeSectors:
            f.write(b'\xFF\xFF\xFF\xFF' * (128 - freeSectors))

        # Finally, return the current sector index for use in other places.
        return numDifat + numFat

    def _writeDirectoryEntries(self, f, startingSector : int) -> List[_DirectoryEntry]:
        """
        Writes out all the directory entries. Returns the list generated.
        """
        entries = self._treeSort(startingSector)
        for x in entries:
            self._writeDirectoryEntry(f, x)
        if len(entries) & 3:
            f.write(((b'\x00\x00' * 34) + (b'\xFF\xFF' * 6) + (b'\x00\x00' * 24)) * (4 - (len(entries) & 3)))

        return entries

    def _writeDirectoryEntry(self, f, entry : _DirectoryEntry) -> None:
        """
        Writes the directory entry to the file f.
        """
        f.write(entry.toBytes())

    def _writeFinal(self, f) -> None:
        """
        Writes the final sectors of the file, consisting of the streams too
        large for the mini FAT.
        """
        for x in self.__largeEntries:
            f.write(x.data)
            if len(x.data) & 511:
                f.write(b'\x00' * (512 - (len(x.data) & 511)))

    def _writeMini(self, f, entries : List[_DirectoryEntry]) -> None:
        """
        Writes the mini FAT followed by the full mini stream.
        """
        # For each of the entires that are streams and less than 4096.
        currentSector = 0
        for x in entries:
            if x.type == DirectoryEntryType.STREAM and len(x.data) < 4096:
                size = ceilDiv(len(x.data), 64)
                for x in range(currentSector + 1, currentSector + size):
                    f.write(constants.ST_LE_UI32.pack(x))
                if size > 0:
                    f.write(b'\xFE\xFF\xFF\xFF')
                currentSector += size

        # Finally, write the remaining slots.
        if currentSector & 127:
            f.write(b'\xFF\xFF\xFF\xFF' * (128 - (currentSector & 127)))

        # Write the mini stream.
        for x in entries:
            if len(x.data) > 0 and len(x.data) < 4096:
                f.write(x.data)
                if len(x.data) & 63:
                    f.write(b'\x00' * (64 - (len(x.data) & 63)))

        # Pad the final mini stream block.
        if self.__numMinifatSectors & 7:
            f.write((b'\x00' * 64) * (8 - (self.__numMinifatSectors & 7)))

    def addOleEntry(self, path, entry : OleDirectoryEntry, data : Optional[bytes] = None) -> None:
        """
        Uses the entry provided to add the data to the writer.

        :raises ValueError: Tried to add an entry to a path that has not yet
            been added.
        """
        pathList = inputToMsgPath(path)
        # First, find the current place in our dict to add the item.
        _dir = self.__dirEntries
        while len(pathList) > 1:
            if pathList[0] not in _dir:
                print(pathList[0])
                print(_dir)
                print(self.__dirEntries)
                # If no entry has been provided already for the directory, that
                # is considered a fatal error.
                raise ValueError('Path not found.')
            _dir = _dir[pathList[0]]
            pathList.pop(0)

        # Now that we are in the right place, add our data.
        newEntry = _DirectoryEntry()
        if entry.entry_type == DirectoryEntryType.STORAGE:
            # Handle a storage entry.
            # First add the dict to our tree of items.
            _dir[pathList[0]] = {'::DirectoryEntry': newEntry}

            # Finally, setup the values for the stream.
            newEntry.name = entry.name
            newEntry.type = DirectoryEntryType.STORAGE
            newEntry.clsid = _unClsid(entry.clsid)
            newEntry.stateBits = entry.dwUserFlags
        else:
            # Handle a stream entry.
            # First add the entry to out dict of entries.
            _dir[pathList[0]] = newEntry
            newEntry.name = entry.name
            newEntry.type = DirectoryEntryType.STREAM
            newEntry.clsid = _unClsid(entry.clsid)
            newEntry.stateBits = entry.dwUserFlags

            # Finally, handle the data.
            newEntry.data = data or b''
            if len(newEntry.data) < 4096:
                self.__numMinifatSectors += ceilDiv(len(newEntry.data), 64)
            else:
                self.__largeEntries.append(newEntry)
                self.__largeEntrySectors += ceilDiv(len(newEntry.data), 512)

        self.__dirEntryCount += 1

    def fromMsg(self, msg : 'MSGFile') -> None:
        """
        Copies the streams and stream information necessary from the MSG file.
        """
        # Get the root OLE entry's CLSID.
        self.__rootClsid = _unClsid(msg._getOleEntry('/').clsid)

        # List both storages and directories, but sort them by shortest length
        # first to prevent errors.
        entries = msg.listDir(True, True, False)
        entries.sort(key = len)

        for x in entries:
            entry = msg._getOleEntry(x)
            data = msg._getStream(x) if entry.entry_type == DirectoryEntryType.STREAM else None
            # THe properties stream on embedded messages actualy needs to be
            # transformed a little (*why* it is like that is a mystery to me).
            # Basically we just need to add a "reserved" section to it in a
            # specific place. So let's check if we are doing the properties
            # stream and then if we are embedded.
            if x[0] == '__properties_version1.0' and msg.prefixLen > 0:
                data = data[:24] + b'\x00\x00\x00\x00\x00\x00\x00\x00' + data[24:]
            self.addOleEntry(x, entry, data)

        # Now check if it is an embedded file. If so, we need to copy the named
        # properties streams (the metadata, not the values).
        if msg.prefixLen > 0:
            # Get the entry for the named properties directory and add it
            # immediately if it exists. If it doesn't exist, this whole
            # section will be skipped.
            self.addOleEntry('__nameid_version1.0', msg._getOleEntry('__nameid_version1.0', False), None)

            # Now that we know it exists, grab all the file inside and copy
            # them to our root.
            # Create our generator.
            gen = (x for x in msg._oleListDir() if len(x) > 1 and x[0] == '__nameid_version1.0')
            for x in gen:
                self.addOleEntry(x, msg._getOleEntry(x, prefix = False), msg._getStream(x, prefix = False))

    def fromOleFile(self, ole : OleFileIO, rootPath = []) -> None:
        """
        Copies all the streams from the proided OLE file into this writer.
        """
        # Copy the clsid of the root entry.
        self.__rootClsid = _unClsid(ole.direntries[0].clsid)

        # Copy all of the other entries.
        for x in ole.listdir(True, True):
            entry = ole.direntries[ole._find(x)]

            if entry.entry_type == DirectoryEntryType.STREAM:
                with ole.openstream(x) as f:
                    data = f.read()
            else:
                data = None

            self.addOleEntry(x, entry, data)

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
            offset = self._writeBeginning(f)
            entries = self._writeDirectoryEntries(f, offset)
            self._writeMini(f, entries)
            self._writeFinal(f)
        finally:
            if opened:
                f.close()



def _unClsid(clsid : str) -> bytes:
    """
    Converts the clsid from olefile.olefile._clsid back to bytes.
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
