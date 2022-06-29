import datetime

from typing import Set

from .calendar_base import CalendarBase
from .enums import ClientIntentFlag


class Calendar(CalendarBase):
    """
    A calendar object.
    """

    @property
    def clientIntent(self) -> Set[ClientIntentFlag]:
        """
        A set of the actions a user has taken on a Meeting object.
        """
        return self._ensureSetNamed('_clientIntent', '0015', overrideClass = ClientIntentFlag.fromBits)

    @property
    def fExceptionalAttendees(self) -> bool:
        """
        Indicates that it is a Recurring Calendar object with one or more
        excpetions and that at least one of the Exception Embedded Message
        objects has at least one RecipientRow structure.

        SHOULD NOT be set for any Calendar object other than that of the
        organizer's.
        """
        return self._ensureSetNamed('fExceptionalAttendees', '822B')

    @property
    def reminderDelta(self) -> int:
        """
        The interval, in minutes, between the time at which the reminder first
        becomes overdue and the start time of the Calendar object.
        """
        return self._ensureSetNamed('_reminderDelta', '8501')

    @property
    def reminderFileParameter(self) -> str:
        """
        The full path (MAY only specify the file name) of the sound that a
        client SHOULD play when the reminder for the Message Object becomes
        overdue.
        """
        return self._ensureSetNamed('_reminderFileParameter', '851F')

    @property
    def reminderOverride(self) -> bool:
        """
        Specifies if clients SHOULD respect the value of the reminderPlaySound
        property and the reminderFileParameter property.
        """
        return self._ensureSetNamed('_reminderOverride', '851C')

    @property
    def reminderPlaySound(self) -> bool:
        """
        Specified that the cliebnt should play a sound when the reminder becomes
        overdue.
        """
        return self._ensureSetNamed('_reminderPlaySound', '851E')

    @property
    def reminderSet(self) -> bool:
        """
        Specifies whether a reminder is set on the object.
        """
        return self._ensureSetNamed('_reminderSet', '8503')

    @property
    def reminderSignalTime(self) -> datetime.datetime:
        """
        The point in time when a reminder transitions from pending to overdue.
        """
        return self._ensureSetNamed('_reminderSignalTime', '8560')

    @property
    def reminderTime(self) -> datetime.datetime:
        """
        The time after which the user would be late.
        """
        return self._ensureSetNamed('_reminderTime', '8502')
