__all__ = [
    'Calendar',
]


import datetime
import functools

from typing import Optional

from ..constants import ps
from .calendar_base import CalendarBase
from ..enums import ClientIntentFlag


class Calendar(CalendarBase):
    """
    A calendar object.
    """

    @functools.cached_property
    def clientIntent(self) -> Optional[ClientIntentFlag]:
        """
        A set of the actions a user has taken on a Meeting object.
        """
        return self.getNamedAs('0015', ps.PSETID_CALENDAR_ASSISTANT, ClientIntentFlag)

    @functools.cached_property
    def fExceptionalAttendees(self) -> Optional[bool]:
        """
        Indicates that it is a Recurring Calendar object with one or more
        excpetions and that at least one of the Exception Embedded Message
        objects has at least one RecipientRow structure.

        SHOULD NOT be set for any Calendar object other than that of the
        organizer's.
        """
        return self.getNamedProp('822B', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def reminderDelta(self) -> Optional[int]:
        """
        The interval, in minutes, between the time at which the reminder first
        becomes overdue and the start time of the Calendar object.
        """
        return self.getNamedProp('8501', ps.PSETID_COMMON)

    @functools.cached_property
    def reminderFileParameter(self) -> Optional[str]:
        """
        The full path (MAY only specify the file name) of the sound that a
        client SHOULD play when the reminder for the Message Object becomes
        overdue.
        """
        return self.getNamedProp('851F', ps.PSETID_COMMON)

    @functools.cached_property
    def reminderOverride(self) -> bool:
        """
        Specifies if clients SHOULD respect the value of the reminderPlaySound
        property and the reminderFileParameter property.
        """
        return bool(self.getNamedProp('851C', ps.PSETID_COMMON))

    @functools.cached_property
    def reminderPlaySound(self) -> bool:
        """
        Specified that the cliebnt should play a sound when the reminder becomes
        overdue.
        """
        return bool(self.getNamedProp('851E', ps.PSETID_COMMON))

    @functools.cached_property
    def reminderSet(self) -> bool:
        """
        Specifies whether a reminder is set on the object.
        """
        return bool(self.getNamedProp('8503', ps.PSETID_COMMON))

    @functools.cached_property
    def reminderSignalTime(self) -> Optional[datetime.datetime]:
        """
        The point in time when a reminder transitions from pending to overdue.
        """
        return self.getNamedProp('8560', ps.PSETID_COMMON)

    @functools.cached_property
    def reminderTime(self) -> Optional[datetime.datetime]:
        """
        The time after which the user would be late.
        """
        return self.getNamedProp('8502', ps.PSETID_COMMON)
