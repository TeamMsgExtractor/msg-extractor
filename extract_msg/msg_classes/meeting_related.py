__all__ = [
    'MeetingRelated',
]


import datetime
import functools

from typing import Optional

from ..constants import ps
from .calendar_base import CalendarBase
from ..enums import ServerProcessingAction


class MeetingRelated(CalendarBase):
    """
    Base class for meeting-related objects.
    """

    @functools.cached_property
    def attendeeCriticalChange(self) -> Optional[datetime.datetime]:
        """
        The date and time at which the meeting-related object was sent.
        """
        return self._getNamedAs('0001', ps.PSETID_MEETING)

    @functools.cached_property
    def processed(self) -> bool:
        """
        Indicates whether a client has processed a meeting-related object.
        """
        return self._getPropertyAs('7D01000B', bool, False)

    @functools.cached_property
    def serverProcessed(self) -> bool:
        """
        Indicates that the Meeting Request object or Meeting Update object has
        been processed.
        """
        return self._getNamedAs('85CC', ps.PSETID_CALENDAR_ASSISTANT, bool, False)

    @functools.cached_property
    def serverProcessingActions(self) -> Optional[ServerProcessingAction]:
        """
        A union of which actions have been taken on the Meeting Request object
        or Meeting Update object.
        """
        return self._getNamedAs('85CD', ps.PSETID_CALENDAR_ASSISTANT, ServerProcessingAction)

    @functools.cached_property
    def timeZone(self) -> Optional[int]:
        """
        Specifies information about the time zone of a recurring meeting.

        See PidLidTimeZone in [MS-OXOCAL] for details.
        """
        return self._getNamedAs('000C', ps.PSETID_MEETING)

    @functools.cached_property
    def where(self) -> Optional[str]:
        """
        PidLidWhere. Should be the same as location.
        """
        return self._getNamedAs('0002', ps.PSETID_MEETING)
