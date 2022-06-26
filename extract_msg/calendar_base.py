import datetime

from typing import List, Set

from .enums import AppointmentAuxilaryFlag, AppointmentStateFlag, BusyStatus, ResponseStatus
from .message_base import MessageBase


class CalendarBase(MessageBase):
    """
    Common base for all Appointment and Meeting objects.
    """

    @property
    def appointmentAuxilaryFlags(self) -> Set[AppointmentAuxilaryFlag]:
        """
        The auxiliary state of the object.
        """
        return self._ensureSetNamed('_appointmentAuxilaryFlags', '8207', overrideClass = AppointmentAuxilaryFlag.fromBits)

    @property
    def appointmentDuration(self) -> int:
        """
        The length of the event, in minutes.
        """
        return self._ensureSetNamed('_appointmentDuration', '8213')

    @property
    def appointmentEndWhole(self) -> datetime.datetime:
        """
        The end date and time of the event in UTC.
        """
        return self._ensureSetNamed('_appointmentEndWhole', '820E')

    @property
    def appointmentSequence(self) -> int:
        """
        Specified the sequence number of a Meeting object. A meeting object
        begins with the sequence number set to 0 and is incremented each time
        the organizer sends out a Meeting Update object.
        """
        return self._ensureSetNamed('_appointmentSequence', '8201')

    @property
    def appointmentStartWhole(self) -> datetime.datetime:
        """
        The start date and time of the event in UTC.
        """
        return self._ensureSetNamed('_appointmentStartWhole', '820D')

    @property
    def appointmentStateFlags(self) -> Set[AppointmentStateFlag]:
        """
        The appointment state of the object.
        """
        return self._ensureSetNamed('_appointmentStateFlags', '8217', overrideClass = AppointmentStateFlag.fromBits)

    @property
    def appointmentSubType(self) -> bool:
        """
        Whether the event is an all-day event or not.
        """
        return self._ensureSetNamed('_appointmentSubType', '8215', overrideClass = bool)

    @property
    def busyStatus(self) -> BusyStatus:
        """
        Specified the availability of a user for the event described by the
        object.
        """
        return self._ensureSetNamed('_busyStatus', '8205', overrideClass = BusyStatus)

    @property
    def endDate(self) -> datetime.datetime:
        """
        The end date of the appointment.
        """
        return self._ensureSetProperty('_endDate', '00610040')

    @property
    def isRecurring(self) -> bool:
        """
        Whether the object is associated with a recurring series.
        """
        return self._ensureSetNamed('appointmentSubType', '0005', overrideClass = bool)

    @property
    def keywords(self) -> List[str]:
        """
        The color to be used when displaying a Calendar object.
        """
        return self._ensureSet('_keywords', 'Keywords')

    @property
    def location(self) -> str:
        """
        Returns the location of the meeting.
        """
        return self._ensureSetNamed('_location', '8208')

    @property
    def optionalAttendees(self) -> str:
        """
        Returns the optional attendees of the meeting.
        """
        return self._ensureSetNamed('_optionalAttendees', '0007')

    @property
    def recurring(self) -> bool:
        """
        Specifies whether the object represents a recurring series.
        """
        return self._ensureSetNamed('_recurring', '8223', overrideClass = bool)

    @property
    def requiredAttendees(self) -> str:
        """
        Returns the required attendees of the meeting.
        """
        return self._ensureSetNamed('_requiredAttendees', '0006')

    @property
    def resourceAttendees(self) -> str:
        """
        Returns the resource attendees of the meeting.
        """
        return self._ensureSetNamed('_resourceAttendees', '0008')

    @property
    def responseStatus(self) -> ResponseStatus:
        """
        The response status of an attendee.
        """
        return self._ensureSetNamed('_responseStatus', '8218', overrideClass = ResponseStatus)

    @property
    def startDate(self) -> datetime.datetime:
        """
        The start date of the appointment.
        """
        return self._ensureSetProperty('_startDate', '00600040')

    @property
    def timeZone(self):
        """
        Returns the time zone of the meeting.
        """
        return self._ensureSetNamed('_timeZone', '000C')

    @property
    def where(self) -> str:
        """
        PidLidWhere. Should be the same as location.
        """
        return self._ensureSetNamed('_where', '0002')
