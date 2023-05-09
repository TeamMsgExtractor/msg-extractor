__all__ = [
    'Calendar',
]


import datetime

from typing import Optional, Set

from . import constants
from .calendar_base import CalendarBase
from .enums import ClientIntentFlag


class Calendar(CalendarBase):
    """
    A calendar object.
    """

    @property
    def clientIntent(self) -> Optional[Set[ClientIntentFlag]]:
        """
        A set of the actions a user has taken on a Meeting object.
        """
        return self._ensureSetNamed('_clientIntent', '0015', constants.PSETID_CALENDAR_ASSISTANT, overrideClass = ClientIntentFlag.fromBits)

    @property
    def fExceptionalAttendees(self) -> Optional[bool]:
        """
        Indicates that it is a Recurring Calendar object with one or more
        excpetions and that at least one of the Exception Embedded Message
        objects has at least one RecipientRow structure.

        SHOULD NOT be set for any Calendar object other than that of the
        organizer's.
        """
        return self._ensureSetNamed('_fExceptionalAttendees', '822B', constants.PSETID_APPOINTMENT)

    @property
    def reminderDelta(self) -> Optional[int]:
        """
        The interval, in minutes, between the time at which the reminder first
        becomes overdue and the start time of the Calendar object.
        """
        return self._ensureSetNamed('_reminderDelta', '8501', constants.PSETID_COMMON)

    @property
    def reminderFileParameter(self) -> Optional[str]:
        """
        The full path (MAY only specify the file name) of the sound that a
        client SHOULD play when the reminder for the Message Object becomes
        overdue.
        """
        return self._ensureSetNamed('_reminderFileParameter', '851F', constants.PSETID_COMMON)

    @property
    def reminderOverride(self) -> bool:
        """
        Specifies if clients SHOULD respect the value of the reminderPlaySound
        property and the reminderFileParameter property.
        """
        return self._ensureSetNamed('_reminderOverride', '851C', constants.PSETID_COMMON, overrideClass = bool, preserveNone = False)

    @property
    def reminderPlaySound(self) -> bool:
        """
        Specified that the cliebnt should play a sound when the reminder becomes
        overdue.
        """
        return self._ensureSetNamed('_reminderPlaySound', '851E', constants.PSETID_COMMON, overrideClass = bool, preserveNone = False)

    @property
    def reminderSet(self) -> bool:
        """
        Specifies whether a reminder is set on the object.
        """
        return self._ensureSetNamed('_reminderSet', '8503', constants.PSETID_COMMON, overrideClass = bool, preserveNone = False)

    @property
    def reminderSignalTime(self) -> Optional[datetime.datetime]:
        """
        The point in time when a reminder transitions from pending to overdue.
        """
        return self._ensureSetNamed('_reminderSignalTime', '8560', constants.PSETID_COMMON)

    @property
    def reminderTime(self) -> Optional[datetime.datetime]:
        """
        The time after which the user would be late.
        """
        return self._ensureSetNamed('_reminderTime', '8502', constants.PSETID_COMMON)
