import datetime

from . import constants
from .enums import AppointmentStateFlag, BusyStatus, RecurPatternType, ResponseStatus
from .calendar import Calendar
from .structures.entry_id import EntryID


class AppointmentMeeting(Calendar):
    """
    Parser for Microsoft Outlook Appointment or Meeting files.

    Both Appointment and Meeting have the same class type but Meeting has
    additional properties. These properties are meaningless on an Appointment
    object.
    """

    @property
    def appointmentCounterProposal(self) -> bool:
        """
        Indicates to the organizer that there are counter proposals that have
        not been accepted or rejected by the organizer.
        """
        return self._ensureSetNamed('_appointmentCounterProposal', '8257')

    @property
    def appointmentLastSequence(self) -> int:
        """
        The last sequence number that was sent to any attendee.
        """
        return self._ensureSetNamed('_appointmentLastSequence', '8203')

    @property
    def appointmentProposalNumber(self) -> int:
        """
        The number of attendees who have sent counter propostals that have not
        been accepted or rejected by the organizer.
        """
        return self._ensureSetNamed('_appointmentProposalNumber', '8259')

    @property
    def appointmentReplyName(self) -> datetime.datetime:
        """
        The user who last replied to the meeting request or meeting update.
        """
        return self._ensureSetNamed('_appointmentReplyTime', '8230')

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
    def autoFillLocation(self) -> bool:
        """
        A value of True indicates that the value of the location property is set
        to the value of the displayName property from the recipientRow structure
        that represents a Resource object.

        A value of False indicates that the value of the location property is
        not automatically set.
        """
        return self._ensureSetNamed('_autoFillLocation', '823A', overrideClass = bool)

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
        meetingOrganizerString = {
            ResponseStatus.NONE: None,
            ResponseStatus.ORGANIZED: 'Meeting organizer',
            ResponseStatus.TENTATIVE: 'Tentatively accepted',
            ResponseStatus.ACCEPTED: 'Accepted',
            ResponseStatus.DECLINED: 'Declined',
            ResponseStatus.NOT_RESPONDED: 'Not yet responded',
        }

        # Get the recurrence string.
        recur = '(none)'
        if self.appointmentRecur:
            recur = {
                RecurPatternType.DAY: 'Daily',
                RecurPatternType.WEEK: 'Weekly',
                RecurPatternType.MONTH: 'Monthly',
                RecurPatternType.MONTH_NTH: 'Monthly',
                RecurPatternType.MONTH_END: 'Monthly',
                RecurPatternType.HJ_MONTH: 'Monthly',
                RecurPatternType.HJ_MONTH_NTH: 'Monthly',
                RecurPatternType.HJ_MONTH_END: 'Monthly',
            }[self.appointmentRecur.patternType]

        return {
            {
                'Subject': self.subject,
                'Location': self.location,
            },
            {
                'Start': self.startDate,
                'End': self.endDate,
            },
            {
                'Recurrance': recur,
                'Recurrence Pattern': self.recurrencePattern,
            },
            {
                'Meeting Status': meetingOrganizerString,
            },
            {
                'Organizer': self.organizer,
                'Required Attendees': self.requiredAttendees,
                'Optional Attendees': self.optionalAttendees,
            },
        }

    @property
    def isMeeting(self) -> bool:
        """
        Attempts to determine if the object is a Meeting. True if meeting, False
        if appointment.
        """
        return AppointmentStateFlag.MEETING in self.appointmentStateFlags

    @property
    def originalStoreEntryID(self) -> EntryID:
        """
        The EntryID of the delegator's message store.
        """
        return self._ensureSetNamed('_originalStoreEntryID', '8237', overrideClass = EntryID.autoCreate)
