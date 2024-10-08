from __future__ import annotations


__all__ = [
    'BusinessCardDisplayDefinition',
    'FieldInfo',
]


import logging
import struct

from typing import List, Optional, Tuple

from ._helpers import BytesReader
from .. import constants
from ..enums import (
        BCImageAlignment, BCImageSource, BCLabelFormat, BCTemplateID,
        BCTextFormat
    )


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class BusinessCardDisplayDefinition:
    """
    Data structure for PidLidBusinessCardDisplayDefinition.

    Contains information used to contruct a business card for a contact.
    """

    def __init__(self, data: bytes):
        reader = BytesReader(data)
        unpacked = constants.st.ST_BC_HEAD.unpack(reader.read(13))
        # Because doc says it must be ignored, we don't check the reserved here.
        reader.read(4)
        self.__majorVersion = unpacked[0]
        if self.__majorVersion < 3:
            raise ValueError('Major version was less than 3.')
        self.__minorVersion = unpacked[1]
        self.__templateID = BCTemplateID(unpacked[2])
        countOfFields = unpacked[3]
        if unpacked[4] != 16:
            raise ValueError('Value of FieldInfoSize was not 16.')
        extraInfoSize = unpacked[5]
        self.__imageAlignment = BCImageAlignment(unpacked[6])
        self.__imageSource = BCImageSource(min(unpacked[7], 1))
        self.__backgroundColor = unpacked[8:11]
        self.__imageArea = unpacked[11]
        extraInfoField = data[17 + 16 * countOfFields:]
        extraInfoField = data[:extraInfoSize]
        self.__fields = [FieldInfo(reader.read(16), extraInfoField)
                         for _ in range(countOfFields)]

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        # Pregenerate the extraInfo field.
        offsets = []
        extraInfo = b''
        for info in self.__fields:
            if info.labelText:
                offsets.append(len(extraInfo))
                extraInfo += info.labelText.encode('utf-16-le') + b'\x00\x00'
            else:
                offsets.append(0xFFFE)

            if len(extraInfo) > 255:
                raise ValueError('Label data can only be 127 characters total, including null characters.')

        # Now that we have what we need, pack the header.
        ret = constants.st.ST_BC_HEAD.pack(
                                           self.__majorVersion,
                                           self.__minorVersion,
                                           self.__templateID,
                                           len(self.__fields),
                                           16,
                                           len(extraInfo),
                                           self.__imageAlignment,
                                           self.__imageSource,
                                           *self.__backgroundColor,
                                           self.__imageArea
                                          )

        # Add the reserved.
        ret += b'\x00\x00\x00\x00'

        # Add each FieldInfo structure.
        for index, info in enumerate(self.__fields):
            ret += info.toBytes(offsets[index])

        return ret + extraInfo

    @property
    def backgroundColor(self) -> Tuple[int, int, int]:
        """
        A tuple of the RGB value of the color of the background.
        """
        return self.__backgroundColor

    @backgroundColor.setter
    def backgroundColor(self, value: Tuple[int, int, int]) -> None:
        if not isinstance(value, tuple) or len(value) != 3:
            raise TypeError(':property backgroundColor: MUST be a tuple of 3 ints.')
        # Quickly try to pack the ints and raise a value error if that fails.
        try:
            constants.st.ST_RGB(*value)
        except struct.error:
            raise ValueError('Value for :property backgroundColor: not in range.')

        self.__backgroundColor = value

    @property
    def fields(self) -> List[FieldInfo]:
        """
        The field info structures.
        """
        return self.__fields

    @property
    def imageAlignment(self) -> BCImageAlignment:
        """
        The alignment of the image within the image area.

        Ignored if card is text only.
        """
        return self.__imageAlignment

    @imageAlignment.setter
    def imageAlignment(self, value: BCImageAlignment) -> None:
        if not isinstance(value, BCImageAlignment):
            raise TypeError(':property imageAlignment: MUST be an instance of BCImageAlignment.')

        self.__imageAlignment = value

    @property
    def imageArea(self) -> int:
        """
        An integer that specified the percent of space that the image will
        occupy on the business card.

        Should be between 4 and 50.
        """
        return self.__imageArea

    @imageArea.setter
    def imageArea(self, value: int) -> None:
        if not isinstance(value, int):
            raise TypeError(':property imageArea: MUST be an int.')
        if value < 0:
            raise ValueError(':property imageArea: MUST be positive.')
        if value > 0xFF:
            raise ValueError(':property imageArea: MUST be less than 0x100.')

        if value < 4 or value > 50:
            logger.warning(f'Business card image area was set to a value outside of suggested range of [4, 50] (got {value}).')

        self.__imageArea = value

    @property
    def imageSource(self) -> BCImageSource:
        """
        The source of the image.
        """
        return self.__imageSource

    @imageSource.setter
    def imageSource(self, value: BCImageSource) -> None:
        if not isinstance(value, BCImageSource):
            raise TypeError(':property imageSource: MUST be an instance of BCImageSource.')

        self.__imageSource = value

    @property
    def majorVersion(self) -> int:
        """
        An 8-bit value that specified the major version number.

        Must be 3 or greater.
        """
        return self.__majorVersion

    @majorVersion.setter
    def majorVersion(self, value: int) -> None:
        if not isinstance(value, int):
            raise TypeError(':property majorVersion: MUST be an int.')
        if value < 3:
            raise ValueError(':property majorVersion: MUST be 3 or greater.')
        if value > 0xFF:
            raise ValueError(':property majorVersion: MUST be less than 0x100.')

        self.__majorVersion = value

    @property
    def minorVersion(self) -> int:
        """
        An 8-bit value that specifies the minor version number.

        SHOULD be set to 0.
        """
        return self.__minorVersion

    @minorVersion.setter
    def minorVersion(self, value: int) -> None:
        if not isinstance(value, int):
            raise TypeError(':property minorVersion: MUST be an int.')
        if value < 0:
            raise ValueError(':property minorVersion: MUST be positive.')
        if value > 0xFF:
            raise ValueError(':property minorVersion: MUST be less than 0x100.')

        if value != 0:
            logger.warning('Business card minor version set to non-zero value.')

        self.__minorVersion = value

    @property
    def templateID(self) -> BCTemplateID:
        """
        The layout of the business card.
        """
        return self.__templateID

    @templateID.setter
    def templateID(self, value: BCTemplateID):
        if not isinstance(value, BCTemplateID):
            raise TypeError(':property templateID: MUST be an instance of BCTemplateID.')

        self.__templateID = value



class FieldInfo:
    def __init__(self, data: Optional[bytes] = None, extraInfo: Optional[bytes] = None):
        if not data:
            self.__textPropertyID = 0
            self.__textFormat = BCTextFormat.DEFAULT
            self.__labelFormat = BCLabelFormat.NO_LABEL
            self.__fontSize = 0
            self.__labelText = None
            self.__valueFontColor = (0, 0, 0)
            self.__labelFontColor = (0, 0, 0)
            return

        if extraInfo is None:
            raise ValueError(':param extraInfo: MUST NOT be None if data is provided.')

        unpacked = constants.st.ST_BC_FIELD_INFO.unpack(data)
        self.__textPropertyID = unpacked[0]
        self.__textFormat = BCTextFormat(unpacked[1])
        self.__labelFormat = BCLabelFormat(unpacked[2])
        self.__fontSize = unpacked[3]
        self.__labelText = None if unpacked[4] == 0xFFFE else BytesReader(extraInfo[unpacked[4]:]).readUtf16String()
        self.__valueFontColor = unpacked[5:8]
        self.__labelFontColor = unpacked[8:11]

    def toBytes(self, offset: int) -> bytes:
        """
        Converts to bytes using the offset into the ExtraInfo field.

        :raise ValueError: The offset was out of range.
        """
        if offset < 0 or offset > 0xFFFF:
            raise ValueError('Offset is out of range.')

        if not self.__labelText:
            offset = 0xFFFE

        return constants.st.ST_BC_FIELD_INFO.pack(
                                                  self.__textPropertyID,
                                                  self.__textFormat,
                                                  self.__labelFormat,
                                                  self.__fontSize,
                                                  offset,
                                                  *self.__valueFontColor,
                                                  *self.__labelFontColor
                                                 )

    @property
    def fontSize(self) -> int:
        """
        An integer that specifies the font size, in points, of the text field.

        MUST be between 3 and 32, or MUST be 0 if the text field is displayed as
        an empty line.
        """
        return self.__fontSize

    @fontSize.setter
    def fontSize(self, value: int) -> None:
        if not isinstance(value, int):
            raise TypeError(':property fontSize: MUST be an int.')

        # Unsure if this is meant to be inclusive or exclusive.
        if value == 0 or 3 <= value <= 32:
            self.__fontSize = value
        else:
            raise ValueError(':property fontSize: MUST be 0 or between 3 and 32, inclusive.')

    @property
    def labelFontColor(self) -> Tuple[int, int, int]:
        """
        A tuple of the RGB value of the color of the label.

        Each channel is a number in range [0, 256).
        """
        return self.__labelFontColor

    @labelFontColor.setter
    def labelFontColor(self, value: Tuple[int, int, int]) -> None:
        if not isinstance(value, tuple) or len(value) != 3:
            raise TypeError(':property labelFontColor: MUST be a tuple of 3 ints.')
        # Quickly try to pack the ints and raise a value error if that fails.
        try:
            constants.st.ST_RGB(*value)
        except struct.error:
            raise ValueError('Value for :property labelFontColor: not in range.')

        self.__labelFontColor = value

    @property
    def labelFormat(self) -> BCLabelFormat:
        """
        The format to use for the label.
        """
        return self.__labelFormat

    @labelFormat.setter
    def labelFormat(self, value: BCLabelFormat) -> None:
        if not isinstance(value, BCLabelFormat):
            raise TypeError(':property labelFormat: MUST be an instance of BCLabelFormat.')
        # Check mutually exclusive flags.
        if BCLabelFormat.ALIGN_LEFT in value and BCLabelFormat.ALIGN_RIGHT in value:
            raise ValueError('BCLabelFormat.ALIGN_LEFT and BCLabelFormat.ALIGN_RIGHT are mutually exclusive.')

        if value and BCLabelFormat.ALIGN_LEFT not in value and BCLabelFormat.ALIGN_RIGHT not in value:
            raise ValueError(':property labelFormat: MUST have the ALIGN_LEFT or ALIGN_RIGHT bit set if any other bits are set.')

        if BCLabelFormat > 0x110:
            raise ValueError('Unknown bits set for :property labelFormat:.')

        self.__textPropertyID = value

    @property
    def labelText(self) -> Optional[str]:
        """
        The text of the label, if it exists.
        """
        return self.__labelText

    @labelText.setter
    def labelText(self, value: Optional[str]) -> None:
        if not value:
            self.__labelText = None
        elif not isinstance(value, str):
            raise TypeError(':property labelText: MUST be a str or None.')

        self.__labelText = value

    @property
    def textFormat(self) -> BCTextFormat:
        """
        An enum value representing the formatting to use for the text.
        """
        return self.__textFormat

    @textFormat.setter
    def textFormat(self, value: BCTextFormat) -> None:
        if not isinstance(value, BCTextFormat):
            raise TypeError(':property textFormat: MUST be an instance of BCTextFormat.')
        # Check mutually exclusive flags.
        if BCTextFormat.CENTER in value and BCTextFormat.RIGHT in value:
            raise ValueError('BCTextFormat.RIGHT and BCTextFormat.CENTER are mutually exclusive.')

        if BCTextFormat > 0x101111:
            raise ValueError('Unknown bits set for :property textFormat:.')

        self.__textPropertyID = value

    @property
    def textPropertyID(self) -> int:
        """
        The property to be used for the text field.

        If the value is 0, it represents an empty field.
        """
        return self.__textPropertyID

    @textPropertyID.setter
    def textPropertyID(self, value: int) -> None:
        if not isinstance(value, int):
            raise TypeError(':property textPropertyID: MUST be an int.')
        if value < 0:
            raise ValueError(':property textPropertyID: MUST be positive.')
        if value > 0xFFFF:
            raise ValueError(':property textPropertyID: MUST NOT be greater than 0xFFFF.')

        self.__textPropertyID = value

    @property
    def valueFontColor(self) -> Tuple[int, int, int]:
        """
        A tuple of the RGB value of the color of the text field.

        Each channel is a number in range [0, 256).
        """
        return self.__valueFontColor

    @valueFontColor.setter
    def valueFontColor(self, value: Tuple[int, int, int]) -> None:
        if not isinstance(value, tuple) or len(value) != 3:
            raise TypeError(':property valueFontColor: MUST be a tuple of 3 ints.')
        # Quickly try to pack the ints and raise a value error if that fails.
        try:
            constants.st.ST_RGB(*value)
        except struct.error:
            raise ValueError('Value for :property valueFontColor: not in range.')

        self.__valueFontColor = value
