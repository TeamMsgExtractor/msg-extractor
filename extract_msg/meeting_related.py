import datetime

from typing import Set

from . import constants
from .calendar_base import CalendarBase
from .enums import ServerProcessingAction


class MeetingRelated(CalendarBase):
    """
    Base class for meeting-related objects.
    """

    @property
    def attendeeCriticalChange(self) -> datetime.datetime:
        """
        The date and time at which the meeting-related object was sent.
        """
        return self._ensureSetNamed('_attendeeCriticalChange', '0001')

    @property
    def processed(self) -> bool:
        """
        Indicates whether a client has processed a meeting-related object.
        """
        return self._ensureSetProperty('_processed', '7D01000B')

    @property
    def serverProcessed(self) -> bool:
        """
        Indicates that the Meeting Request object or Meeting Update object has
        been processed.
        """
        return self._ensureSetNamed('_serverProcessed', '85CC')

    @property
    def serverProcessingActions(self) -> Set[ServerProcessingAction]:
        """
        A set of which actions have been taken on the Meeting Request object or
        Meeting Update object.
        """
        return self._ensureSetNamed('_serverProcessingActions', '85CD', overrideClass = ServerProcessingAction.fromBits)

    @property
    def timeZone(self) -> int:
        """
        Specifies information about the time zone of a recurring meeting.

        See PidLidTimeZone in [MS-OXOCAL] for details.
        """
        return self._ensureSetNamed('_timeZone', '000C')

    @property
    def where(self) -> str:
        """
        PidLidWhere. Should be the same as location.
        """
        return self._ensureSetNamed('_where', '0002')
