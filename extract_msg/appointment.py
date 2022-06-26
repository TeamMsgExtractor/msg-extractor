import datetime

from .enums import BusyStatus
from .calendar import Calendar


class AppointmentMeeting(Calendar):
    """
    Parser for Microsoft Outlook Appointment or Meeting files.

    Both Appointment and Meeting have the same class type but Meeting has
    additional properties. These properties are meaningless on an Appointment
    object.
    """

    @property
    def appointmentLastSequence(self) -> int:
        """
        The last sequence number that was sent to any attendee.
        """
        return self._ensureSetNamed('_appointmentLastSequence', '8203')

    @property
    def appointmentReplyTime(self) -> datetime.datetime:
        """
        The date and time at which the attendee responded to a received Meeting
        Request object of Meeting Update object in UTC.
        """
        return self._ensureSetNamed('_appointmentReplyTime', '8220')

    @property
    def appointmentSequenceTime(self) -> datetime.datetime:
        """
        The date and time at which the appointmentSequence property was last
        modified.
        """
        return self._ensureSetNamed('_appointmentSequenceTime', '8202')

    @property
    def fInvited(self) -> bool:
        """
        Whether a Meeting Request object has been sent out.
        """
        return self._ensureSetNamed('_fInvited', '8229', overrideClass = bool, preserveNone = False)

    @property
    def isMeeting(self) -> bool:
        """
        Attempts to determine if the object is a Meeting. True if meeting, False
        if appointment.
        """
        # TODO.
        pass
