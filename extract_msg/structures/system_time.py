__all__ = [
    'SystemTime',
]


from typing import Any, Optional

from .. import constants


class SystemTime:
    """
    A SYSTEMTIME struct, as defined in [MS-DTYP].
    """
    def __init__(self, data: Optional[bytes] = None):
        data = data or (b'\x00' * 16)
        self.unpack(data)

    def __eq__(self, other: Any) -> bool:
        return isinstance(other, SystemTime) and bytes(self) == bytes(other)

    def __ne__(self, other: Any) -> bool:
        return not self.__eq__(other)

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        """
        Packs the current data into bytes.
        """
        return constants.st.ST_SYSTEMTIME.pack(self.__year,
                                               self.__month,
                                               self.__dayOfWeek,
                                               self.__day,
                                               self.__hour,
                                               self.__minute,
                                               self.__second,
                                               self.__milliseconds)

    def unpack(self, data: bytes) -> None:
        """
        Fills out the fields of this instance by unpacking the bytes.
        """
        unpacked = constants.st.ST_SYSTEMTIME.unpack(data)
        self.__year = unpacked[0]
        self.__month = unpacked[1]
        self.__dayOfWeek = unpacked[2]
        self.__day = unpacked[3]
        self.__hour = unpacked[4]
        self.__minute = unpacked[5]
        self.__second = unpacked[6]
        self.__milliseconds = unpacked[7]

    @property
    def day(self) -> int:
        return self.__day

    @day.setter
    def day(self, value: int) -> None:
        if value < 0:
            raise ValueError('Day must be positive.')
        if value > 0xFFFF:
            raise ValueError('Day cannot be greater than 0xFFFF.')

        self.__day = value

    @property
    def dayOfWeek(self) -> int:
        return self.__dayOfWeek

    @dayOfWeek.setter
    def dayOfWeek(self, value: int) -> None:
        if value < 0:
            raise ValueError('Day of week must be positive.')
        if value > 0xFFFF:
            raise ValueError('Day of week cannot be greater than 0xFFFF.')

        self.__dayOfWeek = value

    @property
    def hour(self) -> int:
        return self.__hour

    @hour.setter
    def hour(self, value: int) -> None:
        if value < 0:
            raise ValueError('Hour must be positive.')
        if value > 0xFFFF:
            raise ValueError('Hour cannot be greater than 0xFFFF.')

        self.__hour = value

    @property
    def milliseconds(self) -> int:
        return self.__milliseconds

    @milliseconds.setter
    def milliseconds(self, value: int) -> None:
        if value < 0:
            raise ValueError('Milliseconds must be positive.')
        if value > 0xFFFF:
            raise ValueError('Milliseconds cannot be greater than 0xFFFF.')

        self.__milliseconds = value

    @property
    def minute(self) -> int:
        return self.__minute

    @minute.setter
    def minute(self, value: int) -> None:
        if value < 0:
            raise ValueError('Minute must be positive.')
        if value > 0xFFFF:
            raise ValueError('Minute cannot be greater than 0xFFFF.')

        self.__minute = value

    @property
    def month(self) -> int:
        return self.__month

    @month.setter
    def month(self, value: int) -> None:
        if value < 0:
            raise ValueError('Month must be positive.')
        if value > 0xFFFF:
            raise ValueError('Month cannot be greater than 0xFFFF.')

        self.__month = value

    @property
    def second(self) -> int:
        return self.__second

    @second.setter
    def second(self, value: int) -> None:
        if value < 0:
            raise ValueError('Second must be positive.')
        if value > 0xFFFF:
            raise ValueError('Second cannot be greater than 0xFFFF.')

        self.__second = value

    @property
    def year(self) -> int:
        return self.__year

    @year.setter
    def year(self, value: int) -> None:
        if value < 0:
            raise ValueError('Year must be positive.')
        if value > 0xFFFF:
            raise ValueError('Year cannot be greater than 0xFFFF.')

        self.__year = value

