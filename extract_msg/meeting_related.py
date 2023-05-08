__all__ = [
    'MeetingRelated',
]


import datetime

from typing import Optional, Set

from . import constants
from .calendar_base import CalendarBase
from .enums import ServerProcessingAction


class MeetingRelated(CalendarBase):
    """
    Base class for meeting-related objects.
    """

    @property
    def attendeeCriticalChange(self) -> Optional[datetime.datetime]:
        """
        The date and time at which the meeting-related object was sent.
        """
        return self._ensureSetNamed('_attendeeCriticalChange', '0001', constants.PSETID_MEETING)

    @property
    def processed(self) -> bool:
        """
        Indicates whether a client has processed a meeting-related object.
        """
        return self._ensureSetProperty('_processed', '7D01000B', overrideClass = bool, preserveNone = False)

    @property
    def serverProcessed(self) -> bool:
        """
        Indicates that the Meeting Request object or Meeting Update object has
        been processed.
        """
        return self._ensureSetNamed('_serverProcessed', '85CC', constants.PSETID_CALENDAR_ASSISTANT, overrideClass = bool, preserveNone = False)

    @property
    def serverProcessingActions(self) -> Optional[Set[ServerProcessingAction]]:
        """
        A set of which actions have been taken on the Meeting Request object or
        Meeting Update object.
        """
        return self._ensureSetNamed('_serverProcessingActions', '85CD', constants.PSETID_CALENDAR_ASSISTANT, overrideClass = ServerProcessingAction.fromBits)

    @property
    def timeZone(self) -> Optional[int]:
        """
        Specifies information about the time zone of a recurring meeting.

        See PidLidTimeZone in [MS-OXOCAL] for details.
        """
        return self._ensureSetNamed('_timeZone', '000C', constants.PSETID_MEETING)

    @property
    def where(self) -> Optional[str]:
        """
        PidLidWhere. Should be the same as location.
        """
        return self._ensureSetNamed('_where', '0002', constants.PSETID_MEETING)
