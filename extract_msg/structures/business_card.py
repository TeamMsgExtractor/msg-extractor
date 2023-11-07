__all__ = [
    'BusinessCardDisplayDefinition',
    'FieldInfo',
]


from typing import Optional, Tuple

from ._helpers import BytesReader
from .. import constants
from ..enums import BCImageAlignment, BCImageSource, BCLabelFormat, BCTemplateID, BCTextFormat
from ..utils import bitwiseAdjustedAnd


class BusinessCardDisplayDefinition:
    """
    Data structure for PidLidBusinessCardDisplayDefinition. Contains information
    used to contruct a business card for a contact.
    """

    def __init__(self, data : bytes):
        self.__rawData = data
        reader = BytesReader(data)
        unpacked = constants.st.ST_BC_HEAD.unpack(reader.read(13))
        # Because doc says it must be ignored, we don't check the reserved here.
        reader.read(4)
        self.__majorVersion = unpacked[0]
        self.__minorVersion = unpacked[1]
        self.__templateID = BCTemplateID(unpacked[2])
        self.__countOfFields = unpacked[3]
        self.__fieldInfoSize = unpacked[4] # Must be 16.
        self.__extraInfoSize = unpacked[5]
        self.__imageAlignment = BCImageAlignment(unpacked[6])
        self.__imageSource = BCImageSource(min(unpacked[7], 1))
        self.__backgroundColor = (bitwiseAdjustedAnd(unpacked[8], 0xFF),
                                  bitwiseAdjustedAnd(unpacked[8], 0xFF00),
                                  bitwiseAdjustedAnd(unpacked[8], 0xFF0000))
        self.__imageArea = unpacked[9]
        self.__extraInfoField = data[17 + 16 * self.__countOfFields:]
        self.__fields = tuple(FieldInfo(reader.read(16), self.__extraInfoField) for _ in range(self.__countOfFields))

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        return self.__rawData

    @property
    def backgroundColor(self) -> Tuple[int, int, int]:
        """
        A tuple of the RGB value of the color of the background.
        """
        return self.__backgroundColor

    @property
    def countOfFields(self) -> int:
        """
        The number of FieldInfo structures.
        """
        return self.__countOfFields

    @property
    def extraInfoField(self) -> bytes:
        """
        The raw data in the ExtraInfo field.
        """
        return self.__extraInfoField

    @property
    def extraInfoSize(self) -> int:
        """
        The size, in bytes, of the ExtraInfo field.
        """
        return self.__extraInfoSize

    @property
    def fieldInfoSize(self) -> int:
        """
        The size, in bytes, of each FieldInfo structure.
        """
        return self.__fieldInfoSize

    @property
    def fields(self) -> Tuple['FieldInfo', ...]:
        """
        The field info structures
        """
        return self.__fields

    @property
    def imageAlignment(self) -> BCImageAlignment:
        """
        The alighment of the image within the image area. Ignored if card is
        text only.
        """
        return self.__imageAlignment

    @property
    def imageArea(self) -> int:
        """
        An integet that specified the percent of space that the image will
        occupy on the business card. Should be between 4 and 50.
        """
        return self.__imageArea

    @property
    def imageSource(self) -> BCImageSource:
        """
        The source of the image.
        """
        return self.__imageSource

    @property
    def majorVersion(self) -> int:
        """
        An 8-bit value that specified the major version number. Must be 3 or
        greater.
        """
        return self.__majorVersion

    @property
    def minorVersion(self) -> int:
        """
        An 8-bit value that specifies the minor version number. SHOULD be set to
        0.
        """
        return self.__minorVersion

    @property
    def templateID(self) -> BCTemplateID:
        """
        The layout of the business card.
        """
        return self.__templateID



class FieldInfo:
    def __init__(self, data : bytes, extraInfo : bytes):
        self.__rawData = data
        unpacked = constants.st.ST_BC_FIELD_INFO.unpack(data)
        self.__textPropertyID = unpacked[0]
        self.__textFormat = BCTextFormat(unpacked[1])
        self.__labelFormat = BCLabelFormat(unpacked[2])
        self.__fontSize = unpacked[3]
        self.__labelOffset = unpacked[4]
        self.__labelText = None if unpacked[4] == 0xFFFE else BytesReader(extraInfo[unpacked[4]:]).readUtf16String()
        self.__valueFontColor = (bitwiseAdjustedAnd(unpacked[5], 0xFF),
                            bitwiseAdjustedAnd(unpacked[5], 0xFF00),
                            bitwiseAdjustedAnd(unpacked[5], 0xFF0000))

        self.__labelFontColor = (bitwiseAdjustedAnd(unpacked[6], 0xFF),
                            bitwiseAdjustedAnd(unpacked[6], 0xFF00),
                            bitwiseAdjustedAnd(unpacked[6], 0xFF0000))

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        return self.__rawData

    @property
    def fontSize(self) -> int:
        """
        An integer that specifies the font size, in points, of the text field.
        MUST be between 3 and 32, or MUST be 0 if the text field is displayed as
        an empty line.
        """
        return self.__fontSize

    @property
    def labelFontColor(self) -> Tuple[int, int, int]:
        """
        A tuple of the RGB value of the color of the label.
        """
        return self.__labelFontColor

    @property
    def labelFormat(self) -> BCLabelFormat:
        """
        The format to use for the label.
        """
        return self.__labelFormat

    @property
    def labelOffset(self) -> int:
        """
        An integer that specified the byte offset into the ExtraInfo field of
        BusinessCardDisplayDefinition. If the text field does not have a label,
        must be 0xFFFE.
        """
        return self.__labelOffset

    @property
    def labelText(self) -> Optional[str]:
        """
        The text of the label, if it exists.
        """
        return self.__labelText

    @property
    def textFormat(self) -> BCTextFormat:
        """
        An enum value representing the formatting to use for the text.
        """
        return self.__textFormat

    @property
    def textPropertyID(self) -> int:
        """
        The property to be used for the text field. If the value is 0, it
        represents an empty field.
        """
        return self.__textPropertyID

    @property
    def valueFontColor(self) -> Tuple[int, int, int]:
        """
        A tuple of the RGB value of the color of the text field.
        """
        return self.__valueFontColor
