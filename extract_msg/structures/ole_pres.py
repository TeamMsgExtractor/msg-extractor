from __future__ import annotations


__all__ = [
    'ClipboardFormatOrAnsiString',
    'OLEPresentationStream',
]


import enum
import logging
import struct

from typing import List, Final, Optional, Union

from .. import constants
from ._helpers import BytesReader
from ..enums import ADVF, ClipboardFormat, DVAspect


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class ClipboardFormatOrAnsiString:
    def __init__(self, reader : Union[bytes, BytesReader]):
        if isinstance(reader, bytes):
            reader = BytesReader(reader)

        self.__markerOrLength = reader.readUnsignedInt()
        if self.__markerOrLength > 0xFFFFFFFD:
            self.__ansiString = None
            self.__clipboardFormat = ClipboardFormat(reader.readUnsignedInt())
        elif self.__markerOrLength > 0:
            self.__ansiString = reader.read(self.__markerOrLength)
            self.__clipboardFormat = None
        else:
            self.__ansiString = None
            self.__clipboardFormat = None

    def toBytes(self) -> bytes:
        ret = constants.st.ST_LE_UI32.pack(self.markerOrLength)
        if self.markerOrLength > 0xFFFFFFFD:
            ret += constants.st.ST_LE_UI32.pack(self.clipboardFormat)
        elif self.markerOrLength > 0:
            ret += self.__ansiString
        return ret

    @property
    def ansiString(self) -> Optional[bytes]:
        """
        The null-terminated ANSI string, as bytes, of the name of a registered
        clipboard format. Only set if markerOrLength is not 0x00000000,
        0xFFFFFFFE, or 0xFFFFFFFF.

        Setting this will modify the markerOrLength field automatically.
        """
        return self.__ansiString

    @ansiString.setter
    def _(self, val : bytes) -> None:
        if not val:
            raise ValueError('Cannot set :property ansiString: to None or empty bytes. Set :property markerOrLength: to a value ')

        self.__ansiString = val

    @property
    def clipboardFormat(self) -> Optional[ClipboardFormat]:
        """
        The clipboard format, if any.

        To set this, make sure that :property markerOrLength: is 0xFFFFFFFE or
        0xFFFFFFFF *before* setting.
        """
        return self.__clipboardFormat

    @clipboardFormat.setter
    def _(self, val : ClipboardFormat) -> None:
        if not val:
            raise ValueError('Cannot set clipboard format to None.')
        if self.markerOrLength < 0xFFFFFFFE:
            raise ValueError('Cannot set the clipboard format while the marker or length is not 0xFFFFFFFE or 0xFFFFFFFF')
        self.__clipboardFormat = val

    @property
    def markerOrLength(self) -> int:
        """
        If set the 0x00000000, then neither the format property nor the
        ansiString property will be set. If it is 0xFFFFFFFF or 0xFFFFFFFE, then
        the clipboardFormat property will be set. Otherwise, the ansiString
        property
        will be set.
        """
        return self.__markerOrLength

    @markerOrLength.setter
    def _(self, val : int) -> None:
        if val < 0:
            raise ValueError('markerOrLength must be a positive integer.')
        if val > 0xFFFFFFFF:
            raise ValueError('markerOrLength must be a 4 byte unsigned integer.')

        if val == 0:
            self.__ansiString = None
            self.__clipboardFormat = None
        elif val > 0xFFFFFFFD:
            self.__ansiString = None
            self.__clipboardFormat = ClipboardFormat.CF_BITMAP
        else:
            raise ValueError('Cannot set :property markerOrLength: to a length value. Set :property ansiString: instead.')
        self.__markerOrLength = val



class DevModeA:
    """
    A DEVMODEA structure, as specified in [MS-OLEDS]. For the purposes of
    parsing from bytes, if something goes wrong this will evaluate to False when
    converting to bool. If no data is prodided, the fields are set to default
    values.
    """
    PARSE_STRUCT : Final[struct.Struct] = struct.Struct('<32s32s4HI13h14xI4x4I16x')

    def __init__(self, data : Optional[bytes] = None):
        self.__valid = data is None
        
        # Set default values for fields that may not be initialized.
        self.__orientation = 0
        self.__paperSize = 0
        self.__paperLength = 0
        self.__paperWidth = 0
        self.__scale = 0
        self.__copies = 0
        self.__defaultSource = 0
        self.__printQuality = 0
        self.__color = 0
        self.__duplex = 0
        self.__yResolution = 0
        self.__ttOption = 0
        self.__collate = 0
        self.__nup = 0
        self.__icmMethod = 0
        self.__icmIntent = 0
        self.__mediaType = 0
        self.__ditherType = 0

        if self.__valid:
            self.__deviceName = b'\x00' * 32
            self.__formName = b'\x00' * 32
            self.__specVersion = 0
            self.__driverVersion = 0
            self.__driverExtra = 0
            self.__fields = _DevModeFields[0]
            return

        reader = BytesReader(data)

        try:
            items = reader.readStruct(self.PARSE_STRUCT)
            self.__deviceName = items[0]
            self.__formName = items[1]
            self.__specVersion = items[2]
            self.__driverVersion = items[3]
            if items[4] != self.PARSE_STRUCT.size:
                logger.warn(f'Unexpected `size` field for DevModeA detected ({items[4]})')
            self.__driverExtra = items[5]
            self.__fields = _DevModeFields(items[6])
            # TODO fields specifies if we should read the field or ignore it.
            if _DevModeFields.DM_ORIENTATION in self.__fields:
                self.__orientation = items[7]
            if _DevModeFields.DM_PAPERSIZE in self.__fields:
                self.__paperSize = items[8]
            if _DevModeFields.DM_PAPERLENGTH in self.__fields:
                self.__paperLength = items[9]
            if _DevModeFields.DM_PAPERWIDTH in self.__fields:
                self.__paperWidth = items[10]
            if _DevModeFields.DM_SCALE in self.__fields:
                self.__scale = items[11]
            if _DevModeFields.DM_COPIES in self.__fields:
                self.__copies = items[12]
            if _DevModeFields.DM_DEFAULTSOURCE in self.__fields:
                self.__defaultSource = items[13]
            if _DevModeFields.DM_PRINTQUALITY in self.__fields:
                self.__printQuality = items[14]
            if _DevModeFields.DM_COLOR in self.__fields:
                self.__color = items[15]
            if _DevModeFields.DM_DUPLEX in self.__fields:
                self.__duplex = items[16]
            if _DevModeFields.DM_YRESOLUTION in self.__fields:
                self.__yResolution = items[17]
            if _DevModeFields.DM_TTOPTION in self.__fields:
                self.__ttOption = items[18]
            if _DevModeFields.DM_COLLATE in self.__fields:
                self.__collate = items[19]
            if _DevModeFields.DM_NUP in self.__fields:
                self.__nup = items[20]
            if _DevModeFields.DM_ICMMETHOD in self.__fields:
                self.__icmMethod = items[21]
            if _DevModeFields.DM_ICMINTENT in self.__fields:
                self.__icmIntent = items[22]
            if _DevModeFields.DM_MEDIATYPE in self.__fields:
                self.__mediaType = items[23]
            if _DevModeFields.DM_DITHERTYPE in self.__fields:
                self.__ditherType = items[24]
        except IOError:
            return

        self.__valid = True

    def __bool__(self) -> bool:
        return self.__valid

    def toBytes(self) -> bytes:
        return self.PARSE_STRUCT.pack(
                                     self.__deviceName,
                                     self.__formName,
                                     self.__specVersion,
                                     self.__driverVersion,
                                     self.PARSE_STRUCT.size,
                                     self.__driverExtra,
                                     self.__fields,
                                     self.__orientation,
                                     self.__paperSize,
                                     self.__paperLength,
                                     self.__paperWidth,
                                     self.__scale,
                                     self.__copies,
                                     self.__defaultSource,
                                     self.__printQuality,
                                     self.__color,
                                     self.__duplex,
                                     self.__yResolution,
                                     self.__ttOption,
                                     self.__collate,
                                     self.__nup,
                                     self.__icmMethod,
                                     self.__icmIntent,
                                     self.__mediaType,
                                     self.__ditherType,
                                     )



class _DevModeFields(enum.IntFlag):
    DM_NUP           = 0b00000000000000000000000000000010
    DM_SCALE         = 0b00000000000000000000000000001000
    DM_PAPERWIDTH    = 0b00000000000000000000000000010000
    DM_PAPERLENGTH   = 0b00000000000000000000000000100000
    DM_PAPERSIZE     = 0b00000000000000000000000001000000
    DM_ORIENTATION   = 0b00000000000000000000000010000000
    DM_COLLATE       = 0b00000000000000000000000100000000
    DM_TTOPTION      = 0b00000000000000000000001000000000
    DM_YRESOLUTION   = 0b00000000000000000000010000000000
    DM_DUPLEX        = 0b00000000000000000000100000000000
    DM_COLOR         = 0b00000000000000000001000000000000
    DM_PRINTQUALITY  = 0b00000000000000000010000000000000
    DM_DEFAULTSOURCE = 0b00000000000000000100000000000000
    DM_COPIES        = 0b00000000000000001000000000000000
    DM_ICMMETHOD     = 0b00000000000000010000000000000000
    DM_FORMNAME      = 0b00000000100000000000000000000000
    DM_DITHERTYPE    = 0b00100000000000000000000000000000
    DM_MEDIATYPE     = 0b01000000000000000000000000000000
    DM_ICMINTENT     = 0b10000000000000000000000000000000



class DVTargetDevice:
    """
    Specifies information about a device that renders the presentation data.

    The creator of this data structure MUST NOT assume that it will be
    understood during processing.
    """

    def __init__(self, data : Optional[bytes]):
        self.__driverName = None
        self.__deviceName = None
        self.__portName = None
        self.__extDevMode = None

        if not data:
            return
        reader = BytesReader(data)

        # We have 4 fields to read, and *technically* they may not all even be
        # present, given that this structure can be 4 bytes? Reading all of
        # these is also much more complicated than other structures, as they can
        # technically overlap. We are just going to be *much* more lenient about
        # this structure.
        offset1 = offset2 = offset3 = offset4 = -1
        try:
            offset1 = reader.readUnsignedShort()
            offset2 = reader.readUnsignedShort()
            offset3 = reader.readUnsignedShort()
            offset4 = reader.readUnsignedShort()
        except IOError:
            pass

        if offset1 != -1 and offset1 < len(data):
            reader.seek(offset1)
            try:
                self.__driverName = reader.readByteString()
            except IOError:
                self.__driverName = reader.read()
                if not self.__driverName:
                    self.__driverName = None

        if offset2 != -1 and offset2 < len(data):
            reader.seek(offset2)
            try:
                self.__deviceName = reader.readByteString()
            except IOError:
                self.__deviceName = reader.read()
                if not self.__deviceName:
                    self.__deviceName = None

        if offset3 != -1 and offset3 < len(data):
            reader.seek(offset3)
            try:
                self.__portName = reader.readByteString()
            except IOError:
                self.__portName = reader.read()
                if not self.__portName:
                    self.__portName = None

        if offset4 != -1 and offset4 < len(data):
            reader.seek(offset4)
            try:
                devmode = DevModeA(reader.read(56))
                if devmode:
                    self.__extDevMode = devmode
            except IOError:
                self.__extDevMode = None

    def toBytes(self) -> Optional[bytes]:
        if not (self.driverName or self.deviceName or self.portName or self.extDevMode):
            return None
        currentPosition = 8

        offset1 = 8 if self.__driverName else 0
        if self.__driverName:
            currentPosition += len(self.__driverName) + 1

        offset2 = currentPosition if self.__deviceName else 0
        if self.__deviceName:
            currentPosition += len(self.__deviceName) + 1

        offset3 = currentPosition if self.__portName else 0
        if self.__portName:
            currentPosition += len(self.__portName) + 1

        extDevModeBytes = self.__extDevMode.toBytes() if self.__extDevMode else None
        offset4 = currentPosition if extDevModeBytes else 0

        try:
            ret = struct.pack('<HHHH', offset1, offset2, offset3, offset4)
        except struct.error:
            raise ValueError('DVTargetDevice structure contains too much data.')
        if self.__driverName:
            ret += self.__driverName + b'\x00'
        if self.__deviceName:
            ret += self.__deviceName + b'\x00'
        if self.__portName:
            ret += self.__portName + b'\x00'
        if extDevModeBytes:
            ret += extDevModeBytes

        return ret

    @property
    def driverName(self) -> Optional[bytes]:
        """
        Optional ANSI string that contains a hunt on how to display or print
        presentation data.
        """
        return self.__driverName

    @driverName.setter
    def _(self, data : Optional[bytes]) -> None:
        self.__driverName = None if not data else data

    @property
    def deviceName(self) -> Optional[bytes]:
        """
        Optional ANSI string that contains a hunt on how to display or print
        presentation data.
        """
        return self.__deviceName

    @deviceName.setter
    def _(self, data : Optional[bytes]) -> None:
        self.__deviceName = None if not data else data

    @property
    def portName(self) -> Optional[bytes]:
        """
        Optional ANSI string that contains any arbitrary value.
        """
        return self.__portName

    @portName.setter
    def _(self, data : Optional[bytes]) -> None:
        self.__portName = None if not data else data

    @property
    def extDevMode(self) -> Optional[DevModeA]:
        """
        Optional ANSI string that contains a hunt on how to display or print
        presentation data.
        """
        return self.__extDevMode

    @extDevMode.setter
    def _(self, data : Optional[DevModeA]) -> None:
        self.__extDevMode = None if not data else data



class OLEPresentationStream:
    """
    [MS-OLEDS] OLEPresentationStream.
    """
    ansiClipboardFormat : ClipboardFormatOrAnsiString
    targetDeviceSize : int
    targetDevice : Optional[DVTargetDevice]
    aspect : Union[int, DVAspect]
    lindex : int
    advf : Union[int, ADVF]
    width : int
    height : int
    data : bytes
    reserved2 : Optional[bytes]
    tocSignature : int
    tocEntries : List[TOCEntry]

    def __init__(self, data : bytes):
        reader = BytesReader(data)
        self.ansiClipboardFormat = ClipboardFormatOrAnsiString(reader)

        # Validate the structure based on the documentation.
        if self.ansiClipboardFormat.markerOrLength == 0:
            raise ValueError('Invalid OLEPresentationStream (MarkerOrLength is 0).')
        if self.ansiClipboardFormat.clipboardFormat is ClipboardFormat.CF_BITMAP:
            raise ValueError('Invalid OLEPresentationStream (Format is CF_BITMAP).')
        if 0x201 < self.ansiClipboardFormat.markerOrLength < 0xFFFFFFFE:
            raise ValueError('Invalid OLEPresentationStream (ANSI length was more than 0x201).')

        self.targetDeviceSize = reader.readUnsignedInt()
        if self.targetDeviceSize < 0x4:
            raise ValueError('Invalid OLEPresentationStream (TargetDeviceSize was less than 4).')
        if self.targetDeviceSize > 0x4:
            # Read the TargetDevice field.
            self.targetDevice = DVTargetDevice(reader.read(self.targetDeviceSize))
        else:
            self.targetDevice = None

        self.aspect = reader.readUnsignedInt()
        self.lindex = reader.readUnsignedInt()
        self.advf = reader.readUnsignedInt()

        # Reserved1.
        reader.read(4)

        self.width = reader.readUnsignedInt()
        self.height = reader.readUnsignedInt()
        size = reader.readUnsignedInt()
        self.data = reader.read(size)

        if self.ansiClipboardFormat.clipboardFormat is ClipboardFormat.CF_METAFILEPICT:
            self.reserved2 = reader.read(18)
        else:
            self.reserved2 = None

        self.tocSignature = reader.readUnsignedInt()
        self.tocEntries = []
        if self.tocSignature == 0x494E414E: # b'NANI' in little endian.
            for _ in range(reader.readUnsignedInt()):
                self.tocEntries.append(TOCEntry(reader))



class TOCEntry:
    def __init__(self, reader : Union[bytes, BytesReader]):
        if isinstance(reader, bytes):
            reader = BytesReader(reader)
            self.__clipFormat = ClipboardFormatOrAnsiString(reader)
            targetDeviceSize = reader.readUnsignedInt()
            self.__aspect = reader.readUnsignedInt()
            self.__lindex = reader.readUnsignedInt()
            self.__tymed = reader.readUnsignedInt()
            reader.read(4)
            self.__advf = reader.readUnsignedInt()
            reader.read(4)
            if targetDeviceSize == 0:
                self.__targetDevice = None
            else:
                self.__targetDevice = DVTargetDevice(reader.read(targetDeviceSize))
        # TODO