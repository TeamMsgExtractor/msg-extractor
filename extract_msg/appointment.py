import datetime

from . import constants
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
    def headerFormatProperties(self) -> constants.HEADER_FORMAT_TYPE:
        """
        Returns a dictionary of properties, in order, to be formatted into the
        header. Keys are the names to use in the header while the values are one
        of the following:
        None: Signifies no data was found for the property and it should be
            omitted from the header.
        str: A string to be formatted into the header using the string encoding.
        Tuple[Union[str, None], bool]: A string should be formatted into the
            header. If the bool is True, then place an empty string if the value
            is None, otherwise follow the same behavior as regular None.

        Additional note: If the value is an empty string, it will be dropped as
        well by default.
        """
        return {
            {
                'Subject': self.subject,
                'Location': self.location,
            },
            {
                'Start': self.startDate,
                'End': self.endDate,
                'Show Time As': None, # TODO: Figure this one out.
            },
            {
                'Recurrance': None, # TODO: Requires recurrance property to be finished.
                'Recurrence Pattern': self.recurrencePattern,
            },
            {
                'Organizer': self.organizer,
            },
        }

    @property
    def isMeeting(self) -> bool:
        """
        Attempts to determine if the object is a Meeting. True if meeting, False
        if appointment.
        """
        # TODO.
        pass
