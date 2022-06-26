import datetime

from typing import List, Set

from .calendar import Calendar
from .enums import BusyStatus, CalendarType, MeetingObjectChange, MeetingType


class MeetingRequest(Calendar):
    """

    """

    @property
    def appointmentClassType(self) -> str:
        """
        Indicates the value of the PidTagMessageClass property of the Meeting
        object that is to be generated from the Meeting Request object. MUST
        start with "IPM.Appointment".
        """
        return self._ensureSetNamed('_appointmentClassType', '0024')

    @property
    def calendarType(self) -> CalendarType:
        """
        The value of the CalendarType field from the PidLidAppointmentRecur
        property if the Meeting Request object represents a recurring series or
        an exception.
        """
        return self._ensureSetNamed('_calendarType', '001C', overrideClass = CalendarType)

    @property
    def changeHighlight(self) -> Set[MeetingObjectChange]:
        """
        Soecifies a bit field that indicates how the Meeting object has been
        changed.

        Returns a set of flags.
        """
        return self._ensureSetNamed('_changeHighlight', '8204', overrideClass = MeetingObjectChange.fromBits)

    @property
    def forwardInstance(self) -> bool:
        """
        Indicates that the Meeting Request object represents an exception to a
        recurring series, and it was forwarded (even when forwarded by the
        organizer) rather than being an invitation sent by the organizer.
        """
        return self._ensureSetNamed('_forwardInstance', '820A')

    @property
    def intendedBusyStatus(self) -> BusyStatus:
        """
        The value of the busyStatus on the Meeting object in the organizer's
        calendar at the time the Meeting Request object or Meeting Update object
        was sent.
        """
        return self._ensureSetNamed('_intendedBusyStatus', '8224', overrideClass = BusyStatus)

    @property
    def meetingType(self) -> MeetingType:
        """
        The type of Meeting Request object or Meeting UpdateObject.
        """
        return self._ensureSetNamed('meetingType', '0026', overrideClass = MeetingType)

    @property
    def oldLocation(self) -> str:
        """
        The original value of the location property before a meeting update.
        """
        return self._ensureSetNamed('_oldLocation', '0028')

    @property
    def oldWhenEndWhole(self) -> datetime.datetime:
        """
        The original value of the appointmentEndWhole property before a meeting
        update.
        """
        return self._ensureSetNamed('_oldWhenEndWhole', '002A')

    @property
    def oldWhenStartWhole(self) -> datetime.datetime:
        """
        The original value of the appointmentStartWhole property before a
        meeting update.
        """
        return self._ensureSetNamed('_oldWhenStartWhole', '0029')
