from __future__ import annotations


__all__ = [
    'OLEPresentationStream',
]


from typing import List, Optional, Union

from ._helpers import BytesReader
from .cfoas import ClipboardFormatOrAnsiString
from .dv_target_device import DVTargetDevice
from ..enums import ADVF, ClipboardFormat, DVAspect
from .toc_entry import TOCEntry


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