__all__ = [
    'RecurrencePattern',
]


from typing import Any, Tuple

from ..enums import RecurCalendarType, RecurDOW, RecurEndType, RecurFrequency, RecurMonthNthWeek, RecurPatternType, RecurPatternTypeSpecificWeekday
from ._helpers import BytesReader


class RecurrencePattern:
    """
    A RecurrencePattern structure, as specified in [MS-OXOCAL].
    """

    def __init__(self, data: bytes):
        self.__rawData = data
        reader = BytesReader(data)
        self.__readerVersion = reader.readUnsignedShort()
        self.__writerVersion = reader.readUnsignedShort()
        if not (self.__readerVersion == self.__writerVersion == 0x3004):
            raise ValueError('Reader version or writer version was not set to 0x3004.')

        self.__recurFrequency = RecurFrequency(reader.readUnsignedShort())
        self.__patternType = RecurPatternType(reader.readUnsignedShort())
        self.__calendarType = RecurCalendarType(reader.readUnsignedShort())
        self.__firstDateTime = reader.readUnsignedInt()
        self.__period = reader.readUnsignedInt()
        self.__slidingFlag = reader.readUnsignedInt()
        # This is just here to help shorten lines.
        RPTSW = RecurPatternTypeSpecificWeekday
        # This field changes depending on the recurrence type.
        if self.__patternType == RecurPatternType.DAY:
            self.__patternTypeSpecific = None
        elif self.__patternType == RecurPatternType.WEEK:
            self.__patternTypeSpecific = RPTSW(reader.readUnsignedInt())
        elif self.__patternType in (RecurPatternType.MONTH_NTH, RecurPatternType.HJ_MONTH_NTH):
            self.__patternTypeSpecific = reader.readUnsignedInt()
        else:
            self.__patternTypeSpecific = (RPTSW(reader.readUnsignedInt()),
                                          RecurMonthNthWeek(reader.readUnsignedInt()))

        self.__endType = RecurEndType.fromInt(reader.readUnsignedInt())
        self.__occurrenceCount = reader.readUnsignedInt()
        self.__firstDOW = RecurDOW(reader.readUnsignedInt())
        deletedInstanceCount = reader.readUnsignedInt()
        self.__deletedInstanceDates = tuple(reader.readUnsignedInt() for _ in range(deletedInstanceCount))
        modifiedInstanceCount = reader.readUnsignedInt()
        self.__modifiedInstanceDates = tuple(reader.readUnsignedInt() for _ in range(modifiedInstanceCount))
        self.__startDate = reader.readUnsignedInt()
        self.__endDate = reader.readUnsignedInt()

    def __bytes__(self) -> bytes:
        return self.toBytes()

    def toBytes(self) -> bytes:
        return self.__rawData

    @property
    def calendarType(self) -> RecurCalendarType:
        """
        The type of calendar that is used.
        """
        return self.__calendarType

    @property
    def deletedInstanceDates(self) -> Tuple[int, ...]:
        """
        A tuple of the dates (stored as number of minutes between midnight,
        January 1, 1601, and midnight on the specified day in the timezone
        specified in the calendar object), ordered from earliest to latest, of
        either a deleted instance or a modified instance for this recurrence.
        """
        return self.__deletedInstanceDates

    @property
    def endDate(self) -> int:
        """
        An integer that specifies the ending date for the recurrence.

        The value is the number of minutes between midnight, January 1, 1601,
        and midnight of the date of the last occurrence. When the value of the
        endType field is ``END_AFTER_N_OCCURRENCES``, this value is calculated
        based on the number of occurrences. If the recurrence does not have an
        end date, the value of the endDate field MUST be set to ``0x5AE980DF``.
        """
        return self.__endDate

    @property
    def endType(self) -> RecurEndType:
        """
        The ending type for the recurrence.
        """
        return self.__endType

    @property
    def firstDateTime(self) -> int:
        """
        The first ever dat, week, or month of a recurring series, dating back to
        a reference date, which is January 1, 1601, for a Gregorian calendar.
        """
        return self.__firstDateTime

    @property
    def firstDayOfWeek(self) -> RecurDOW:
        """
        The day on which the calendar week begins.
        """
        return self.__firstDOW

    @property
    def modifiedInstanceDates(self) -> Tuple[int, ...]:
        """
        A tuple of the dates (stored as number of minutes between midnight,
        January 1, 1601, and midnight on the specified day in the timezone
        specified in the calendar object), ordered from earliest to latest, of
        a modified instance.
        """
        return self.__modifiedInstanceDates

    @property
    def occurrenceCount(self) -> int:
        """
        Number of occurrences for a recurrence that ends after N occurrences.
        """
        return self.__occurrenceCount

    @property
    def patternType(self) -> RecurPatternType:
        """
        The type of recurrence pattern.
        """
        return self.__patternType

    @property
    def patternTypeSpecific(self) -> Any:
        """
        The specifics for the pattern type.

        Return is different depending on what type of pattern is being used.

        * RecurPatternType.DAY: No value is returned.
        * RecurPatternType.WEEK: A set of RecurPatternTypeSpecificWeekday bits.
        * RecurPatternType.MONTH: The day of the month on which the recurrence
          falls.
        * RecurPatternType.MONTH_NTH: A tuple containing information from
          [MS-OXOCAL] PatternTypeSpecific MonthNth.
        * RecurPatternType.MONTH_END: The day of the month on which the
          recurrence falls.
        * RecurPatternType.HJ_MONTH: The day of the month on which the
          recurrence falls.
        * RecurPatternType.HJ_MONTH_NTH: A tuple containing information from
          [MS-OXOCAL] PatternTypeSpecific MonthNth.
        * RecurPatternType.HJ_MONTH_END: The day of the month on which the
          recurrence falls.
        """
        return self.__patternTypeSpecific

    @property
    def period(self) -> int:
        """
        An integer that specifies the interval at which the meeting pattern
        specified in PatternTypeSpecific field repeats.

        The Period value MUST be between 1 and the maximum recurrence interval,
        which is 999 days for daily recurrences, 99 weeks for weekly
        recurrences, and 99 months for monthly recurrences.
        """
        return self.__period

    @property
    def readerVersion(self) -> int:
        return self.__readerVersion

    @property
    def recurFrequency(self) -> RecurFrequency:
        """
        The frequency of the recurring series.
        """
        return self.__recurFrequency

    @property
    def slidingFlag(self) -> int:
        return self.__slidingFlag

    @property
    def startDate(self) -> int:
        """
        An integer that specifies the date of the first occurrence.

        The value is the number of minutes between midnight, January 1, 1601,
        and midnight of the date of the first occurrence.
        """
        return self.__startDate

    @property
    def writerVersion(self) -> int:
        return self.__writerVersion
