__all__ = [
    'DevModeA',
]


import logging
import struct

from typing import Final, Optional

from ._helpers import BytesReader
from ..enums import DevModeFields


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class DevModeA:
    """
    A DEVMODEA structure, as specified in [MS-OLEDS].

    For the purposes of parsing from bytes, if something goes wrong this will
    evaluate to ``False`` when converting to ``bool``. If no data is prodided,
    the fields are set to default values.
    """
    PARSE_STRUCT: Final[struct.Struct] = struct.Struct('<32s32s4HI13h14xI4x4I16x')

    def __init__(self, data: Optional[bytes] = None):
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
            self.__fields = DevModeFields[0]
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
            self.__fields = DevModeFields(items[6])
            # TODO fields specifies if we should read the field or ignore it.
            if DevModeFields.DM_ORIENTATION in self.__fields:
                self.__orientation = items[7]
            if DevModeFields.DM_PAPERSIZE in self.__fields:
                self.__paperSize = items[8]
            if DevModeFields.DM_PAPERLENGTH in self.__fields:
                self.__paperLength = items[9]
            if DevModeFields.DM_PAPERWIDTH in self.__fields:
                self.__paperWidth = items[10]
            if DevModeFields.DM_SCALE in self.__fields:
                self.__scale = items[11]
            if DevModeFields.DM_COPIES in self.__fields:
                self.__copies = items[12]
            if DevModeFields.DM_DEFAULTSOURCE in self.__fields:
                self.__defaultSource = items[13]
            if DevModeFields.DM_PRINTQUALITY in self.__fields:
                self.__printQuality = items[14]
            if DevModeFields.DM_COLOR in self.__fields:
                self.__color = items[15]
            if DevModeFields.DM_DUPLEX in self.__fields:
                self.__duplex = items[16]
            if DevModeFields.DM_YRESOLUTION in self.__fields:
                self.__yResolution = items[17]
            if DevModeFields.DM_TTOPTION in self.__fields:
                self.__ttOption = items[18]
            if DevModeFields.DM_COLLATE in self.__fields:
                self.__collate = items[19]
            if DevModeFields.DM_NUP in self.__fields:
                self.__nup = items[20]
            if DevModeFields.DM_ICMMETHOD in self.__fields:
                self.__icmMethod = items[21]
            if DevModeFields.DM_ICMINTENT in self.__fields:
                self.__icmIntent = items[22]
            if DevModeFields.DM_MEDIATYPE in self.__fields:
                self.__mediaType = items[23]
            if DevModeFields.DM_DITHERTYPE in self.__fields:
                self.__ditherType = items[24]
        except IOError:
            return

        self.__valid = True

    def __bool__(self) -> bool:
        return self.__valid

    def __bytes__(self) -> bytes:
        return self.toBytes()

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

    @property
    def collate(self) -> int:
        return self.__collate

    @collate.setter
    def collate(self, val: Optional[int]) -> None:
        if val is None:
            self.__collate = 0
            if DevModeFields.DM_COLLATE in self.__fields:
                self.__fields ^= DevModeFields.DM_COLLATE
            return
        elif not isinstance(val, int):
            raise TypeError(':property collate: must be an int or None.')

        if val < -0x8000:
            raise ValueError(':property collate: cannot be less than -0x8000.')
        if val > 0x7FFF:
            raise ValueError(':property collate: cannot be greater than 0x7FFF.')

        self.__fields |= DevModeFields.DM_COLLATE
        self.__collate = val

    @property
    def color(self) -> int:
        return self.__color

    @color.setter
    def color(self, val: Optional[int]) -> None:
        if val is None:
            self.__color = 0
            if DevModeFields.DM_COLOR in self.__fields:
                self.__fields ^= DevModeFields.DM_COLOR
            return
        elif not isinstance(val, int):
            raise TypeError(':property color: must be an int or None.')

        if val < -0x8000:
            raise ValueError(':property color: cannot be less than -0x8000.')
        if val > 0x7FFF:
            raise ValueError(':property color: cannot be greater than 0x7FFF.')

        self.__fields |= DevModeFields.DM_COLOR
        self.__color = val

    @property
    def copies(self) -> int:
        return self.__copies

    @copies.setter
    def copies(self, val: Optional[int]) -> None:
        if val is None:
            self.__copies = 0
            if DevModeFields.DM_COPIES in self.__fields:
                self.__fields ^= DevModeFields.DM_COPIES
            return
        elif not isinstance(val, int):
            raise TypeError(':property copies: must be an int or None.')

        if val < -0x8000:
            raise ValueError(':property copies: cannot be less than -0x8000.')
        if val > 0x7FFF:
            raise ValueError(':property copies: cannot be greater than 0x7FFF.')

        self.__fields |= DevModeFields.DM_COPIES
        self.__copies = val

    @property
    def defaultSource(self) -> int:
        return self.__defaultSource

    @defaultSource.setter
    def defaultSource(self, val: Optional[int]) -> None:
        if val is None:
            self.__defaultSource = 0
            if DevModeFields.DM_DEFAULTSOURCE in self.__fields:
                self.__fields ^= DevModeFields.DM_DEFAULTSOURCE
            return
        elif not isinstance(val, int):
            raise TypeError(':property defaultSource: must be an int or None.')

        if val < -0x8000:
            raise ValueError(':property defaultSource: cannot be less than -0x8000.')
        if val > 0x7FFF:
            raise ValueError(':property defaultSource: cannot be greater than 0x7FFF.')

        self.__fields |= DevModeFields.DM_DEFAULTSOURCE
        self.__defaultSource = val

    @property
    def deviceName(self) -> bytes:
        """
        A 32 byte ANSI string.
        """
        return self.__deviceName

    @deviceName.setter
    def deviceName(self, val: bytes) -> None:
        if not isinstance(val, bytes):
            raise TypeError(':property deviceName: must be bytes.')
        if len(val) != 32:
            raise ValueError(':property deviceName: must be exactly 32 bytes.')

        self.__deviceName = val

    @property
    def ditherType(self) -> int:
        return self.__ditherType

    @ditherType.setter
    def ditherType(self, val: Optional[int]) -> None:
        if val is None:
            self.__ditherType = 0
            if DevModeFields.DM_DITHERTYPE in self.__fields:
                self.__fields ^= DevModeFields.DM_DITHERTYPE
            return
        elif not isinstance(val, int):
            raise TypeError(':property ditherType: must be an int or None.')

        if val < 0:
            raise ValueError(':property ditherType: must be positive.')
        if val > 0xFFFFFFFF:
            raise ValueError(':property ditherType: cannot be greater than 0xFFFFFFFF.')

        self.__fields |= DevModeFields.DM_DITHERTYPE
        self.__ditherType = val

    @property
    def driverExtra(self) -> int:
        return self.__driverExtra

    @driverExtra.setter
    def driverExtra(self, val: int) -> None:
        if not isinstance(val, int):
            raise TypeError(':property driverExtra: must be an int.')
        if val < 0:
            raise ValueError(':property driverExtra: must be positive.')
        if val > 0xFFFF:
            raise ValueError(':property driverExtra: cannot be greater than 0xFFFF.')

        self.__driverExtra = val

    @property
    def driverVersion(self) -> int:
        return self.__driverVersion

    @driverVersion.setter
    def driverVersion(self, val: int) -> None:
        if not isinstance(val, int):
            raise TypeError(':property driverVersion: must be an int or None.')
        if val < 0:
            raise ValueError(':property driverVersion: must be positive.')
        if val > 0xFFFF:
            raise ValueError(':property driverVersion: cannot be greater than 0xFFFF.')

        self.__driverVersion = val

    @property
    def duplex(self) -> int:
        return self.__duplex

    @duplex.setter
    def duplex(self, val: Optional[int]) -> None:
        if val is None:
            self.__duplex = 0
            if DevModeFields.DM_DUPLEX in self.__fields:
                self.__fields ^= DevModeFields.DM_DUPLEX
            return
        elif not isinstance(val, int):
            raise TypeError(':property duplex: must be an int or None.')

        if val < -0x8000:
            raise ValueError(':property duplex: cannot be less than -0x8000.')
        if val > 0x7FFF:
            raise ValueError(':property duplex: cannot be greater than 0x7FFF.')

        self.__fields |= DevModeFields.DM_DUPLEX
        self.__duplex = val

    @property
    def formName(self) -> bytes:
        """
        A 32 byte ANSI string.
        """
        return self.__formName

    @formName.setter
    def formName(self, val: bytes) -> None:
        if not isinstance(val, bytes):
            raise TypeError(':property formName: must be bytes.')
        if len(val) != 32:
            raise ValueError(':property formName: must be exactly 32 bytes.')

        self.__formName = val

    @property
    def icmIntent(self) -> int:
        return self.__icmIntent

    @icmIntent.setter
    def icmIntent(self, val: Optional[int]) -> None:
        if val is None:
            self.__icmIntent = 0
            if DevModeFields.DM_ICMINTENT in self.__fields:
                self.__fields ^= DevModeFields.DM_ICMINTENT
            return
        elif not isinstance(val, int):
            raise TypeError(':property icmIntent: must be an int or None.')

        if val < 0:
            raise ValueError(':property icmIntent: must be positive.')
        if val > 0xFFFFFFFF:
            raise ValueError(':property icmIntent: cannot be greater than 0xFFFFFFFF.')

        self.__fields |= DevModeFields.DM_ICMINTENT
        self.__icmIntent = val

    @property
    def icmMethod(self) -> int:
        return self.__icmMethod

    @icmMethod.setter
    def icmMethod(self, val: Optional[int]) -> None:
        if val is None:
            self.__icmMethod = 0
            if DevModeFields.DM_ICMMETHOD in self.__fields:
                self.__fields ^= DevModeFields.DM_ICMMETHOD
            return
        elif not isinstance(val, int):
            raise TypeError(':property icmMethod: must be an int or None.')

        if val < 0:
            raise ValueError(':property icmMethod: must be positive.')
        if val > 0xFFFFFFFF:
            raise ValueError(':property icmMethod: cannot be greater than 0xFFFFFFFF.')

        self.__fields |= DevModeFields.DM_ICMMETHOD
        self.__icmMethod = val

    @property
    def mediaType(self) -> int:
        return self.__mediaType

    @mediaType.setter
    def mediaType(self, val: Optional[int]) -> None:
        if val is None:
            self.__mediaType = 0
            if DevModeFields.DM_MEDIATYPE in self.__fields:
                self.__fields ^= DevModeFields.DM_MEDIATYPE
            return
        elif not isinstance(val, int):
            raise TypeError(':property mediaType: must be an int or None.')

        if val < 0:
            raise ValueError(':property mediaType: must be positive.')
        if val > 0xFFFFFFFF:
            raise ValueError(':property mediaType: cannot be greater than 0xFFFFFFFF.')

        self.__fields |= DevModeFields.DM_MEDIATYPE
        self.__mediaType = val

    @property
    def nup(self) -> int:
        return self.__nup

    @nup.setter
    def nup(self, val: Optional[int]) -> None:
        if val is None:
            self.__nup = 0
            if DevModeFields.DM_NUP in self.__fields:
                self.__fields ^= DevModeFields.DM_NUP
            return
        elif not isinstance(val, int):
            raise TypeError(':property nup: must be an int or None.')

        if val < 0:
            raise ValueError(':property nup: must be positive.')
        if val > 0xFFFFFFFF:
            raise ValueError(':property nup: cannot be greater than 0xFFFFFFFF.')

        self.__fields |= DevModeFields.DM_NUP
        self.__nup = val

    @property
    def orientation(self) -> int:
        return self.__orientation

    @orientation.setter
    def orientation(self, val: Optional[int]) -> None:
        if val is None:
            self.__orientation = 0
            if DevModeFields.DM_ORIENTATION in self.__fields:
                self.__fields ^= DevModeFields.DM_ORIENTATION
            return
        elif not isinstance(val, int):
            raise TypeError(':property orientation: must be an int or None.')

        if val < -0x8000:
            raise ValueError(':property orientation: cannot be less than -0x8000.')
        if val > 0x7FFF:
            raise ValueError(':property orientation: cannot be greater than 0x7FFF.')

        self.__fields |= DevModeFields.DM_ORIENTATION
        self.__orientation = val

    @property
    def paperLength(self) -> int:
        return self.__paperLength

    @paperLength.setter
    def paperLength(self, val: Optional[int]) -> None:
        if val is None:
            self.__paperLength = 0
            if DevModeFields.DM_PAPERLENGTH in self.__fields:
                self.__fields ^= DevModeFields.DM_PAPERLENGTH
            return
        elif not isinstance(val, int):
            raise TypeError(':property paperLength: must be an int or None.')

        if val < -0x8000:
            raise ValueError(':property paperLength: cannot be less than -0x8000.')
        if val > 0x7FFF:
            raise ValueError(':property paperLength: cannot be greater than 0x7FFF.')

        self.__fields |= DevModeFields.DM_PAPERLENGTH
        self.__paperLength = val

    @property
    def paperSize(self) -> int:
        return self.__paperSize

    @paperSize.setter
    def paperSize(self, val: Optional[int]) -> None:
        if val is None:
            self.__paperSize = 0
            if DevModeFields.DM_PAPERSIZE in self.__fields:
                self.__fields ^= DevModeFields.DM_PAPERSIZE
            return
        elif not isinstance(val, int):
            raise TypeError(':property paperSize: must be an int or None.')

        if val < -0x8000:
            raise ValueError(':property paperSize: cannot be less than -0x8000.')
        if val > 0x7FFF:
            raise ValueError(':property paperSize: cannot be greater than 0x7FFF.')

        self.__fields |= DevModeFields.DM_PAPERSIZE
        self.__paperSize = val

    @property
    def paperWidth(self) -> int:
        return self.__paperWidth

    @paperWidth.setter
    def paperWidth(self, val: Optional[int]) -> None:
        if val is None:
            self.__paperWidth = 0
            if DevModeFields.DM_PAPERWIDTH in self.__fields:
                self.__fields ^= DevModeFields.DM_PAPERWIDTH
            return
        elif not isinstance(val, int):
            raise TypeError(':property paperWidth: must be an int or None.')

        if val < -0x8000:
            raise ValueError(':property paperWidth: cannot be less than -0x8000.')
        if val > 0x7FFF:
            raise ValueError(':property paperWidth: cannot be greater than 0x7FFF.')

        self.__fields |= DevModeFields.DM_PAPERWIDTH
        self.__paperWidth = val

    @property
    def printQuality(self) -> int:
        return self.__printQuality

    @printQuality.setter
    def printQuality(self, val: Optional[int]) -> None:
        if val is None:
            self.__printQuality = 0
            if DevModeFields.DM_PRINTQUALITY in self.__fields:
                self.__fields ^= DevModeFields.DM_PRINTQUALITY
            return
        elif not isinstance(val, int):
            raise TypeError(':property printQuality: must be an int or None.')

        if val < -0x8000:
            raise ValueError(':property printQuality: cannot be less than -0x8000.')
        if val > 0x7FFF:
            raise ValueError(':property printQuality: cannot be greater than 0x7FFF.')

        self.__fields |= DevModeFields.DM_PRINTQUALITY
        self.__printQuality = val

    @property
    def scale(self) -> int:
        return self.__scale

    @scale.setter
    def scale(self, val: Optional[int]) -> None:
        if val is None:
            self.__scale = 0
            if DevModeFields.DM_SCALE in self.__fields:
                self.__fields ^= DevModeFields.DM_SCALE
            return
        elif not isinstance(val, int):
            raise TypeError(':property scale: must be an int or None.')

        if val < -0x8000:
            raise ValueError(':property scale: cannot be less than -0x8000.')
        if val > 0x7FFF:
            raise ValueError(':property scale: cannot be greater than 0x7FFF.')

        self.__fields |= DevModeFields.DM_SCALE
        self.__scale = val

    @property
    def specVersion(self) -> int:
        return self.__specVersion

    @specVersion.setter
    def specVersion(self, val: int) -> None:
        if not isinstance(val, int):
            raise TypeError(':property specVersion: must be an int.')
        if val < 0:
            raise ValueError(':property specVersion: must be positive.')
        if val > 0xFFFF:
            raise ValueError(':property specVersion: cannot be greater than 0xFFFF.')

        self.__specVersion = val

    @property
    def ttOption(self) -> int:
        return self.__ttOption

    @ttOption.setter
    def ttOption(self, val: Optional[int]) -> None:
        if val is None:
            self.__ttOption = 0
            if DevModeFields.DM_TTOPTION in self.__fields:
                self.__fields ^= DevModeFields.DM_TTOPTION
            return
        elif not isinstance(val, int):
            raise TypeError(':property ttOption: must be an int or None.')

        if val < -0x8000:
            raise ValueError(':property ttOption: cannot be less than -0x8000.')
        if val > 0x7FFF:
            raise ValueError(':property ttOption: cannot be greater than 0x7FFF.')

        self.__fields |= DevModeFields.DM_TTOPTION
        self.__ttOption = val

    @property
    def yResolution(self) -> int:
        return self.__yResolution

    @yResolution.setter
    def yResolution(self, val: Optional[int]) -> None:
        if val is None:
            self.__yResolution = 0
            if DevModeFields.DM_YRESOLUTION in self.__fields:
                self.__fields ^= DevModeFields.DM_YRESOLUTION
            return
        elif not isinstance(val, int):
            raise TypeError(':property yResolution: must be an int or None.')

        if val < -0x8000:
            raise ValueError(':property yResolution: cannot be less than -0x8000.')
        if val > 0x7FFF:
            raise ValueError(':property yResolution: cannot be greater than 0x7FFF.')

        self.__fields |= DevModeFields.DM_YRESOLUTION
        self.__yResolution = val