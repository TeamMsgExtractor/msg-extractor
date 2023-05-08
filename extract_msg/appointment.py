__all__ = [
    'AppointmentMeeting',
]


import datetime

from typing import Optional

from . import constants
from .enums import AppointmentStateFlag, RecurPatternType, ResponseStatus
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
        return self._ensureSetNamed('_appointmentCounterProposal', '8257', constants.PSETID_APPOINTMENT, overrideClass = bool, preserveNone = False)

    @property
    def appointmentLastSequence(self) -> Optional[int]:
        """
        The last sequence number that was sent to any attendee.
        """
        return self._ensureSetNamed('_appointmentLastSequence', '8203', constants.PSETID_APPOINTMENT)

    @property
    def appointmentProposalNumber(self) -> Optional[int]:
        """
        The number of attendees who have sent counter propostals that have not
        been accepted or rejected by the organizer.
        """
        return self._ensureSetNamed('_appointmentProposalNumber', '8259', constants.PSETID_APPOINTMENT)

    @property
    def appointmentReplyName(self) -> Optional[datetime.datetime]:
        """
        The user who last replied to the meeting request or meeting update.
        """
        return self._ensureSetNamed('_appointmentReplyName', '8230', constants.PSETID_APPOINTMENT)

    @property
    def appointmentReplyTime(self) -> Optional[datetime.datetime]:
        """
        The date and time at which the attendee responded to a received Meeting
        Request object of Meeting Update object in UTC.
        """
        return self._ensureSetNamed('_appointmentReplyTime', '8220', constants.PSETID_APPOINTMENT)

    @property
    def appointmentSequenceTime(self) -> Optional[datetime.datetime]:
        """
        The date and time at which the appointmentSequence property was last
        modified.
        """
        return self._ensureSetNamed('_appointmentSequenceTime', '8202', constants.PSETID_APPOINTMENT)

    @property
    def autoFillLocation(self) -> bool:
        """
        A value of True indicates that the value of the location property is set
        to the value of the displayName property from the recipientRow structure
        that represents a Resource object.

        A value of False indicates that the value of the location property is
        not automatically set.
        """
        return self._ensureSetNamed('_autoFillLocation', '823A', constants.PSETID_APPOINTMENT, overrideClass = bool, preserveNone = False)

    @property
    def fInvited(self) -> bool:
        """
        Whether a Meeting Request object has been sent out.
        """
        return self._ensureSetNamed('_fInvited', '8229', constants.PSETID_APPOINTMENT, overrideClass = bool, preserveNone = False)

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

        Additionally you can group members of a header together by placing them
        in an embedded dictionary. Groups will be spaced out using a second
        instance of the join string. If any member of a group is being printed,
        it will be spaced apart from the next group/item.

        If you class should not do *any* header injection, return None from this
        property.
        """
        meetingStatusString = {
            ResponseStatus.NONE: None,
            ResponseStatus.ORGANIZED: 'Meeting organizer',
            ResponseStatus.TENTATIVE: 'Tentatively accepted',
            ResponseStatus.ACCEPTED: 'Accepted',
            ResponseStatus.DECLINED: 'Declined',
            ResponseStatus.NOT_RESPONDED: 'Not yet responded',
        }[self.responseStatus]

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
            '-main info-': {
                'Subject': self.subject,
                'Location': self.location,
            },
            '-date-': {
                'Start': self.startDate.__format__('%a, %d %b %Y %H:%M %z') if self.startDate else None,
                'End': self.endDate.__format__('%a, %d %b %Y %H:%M %z') if self.endDate else None,
            },
            '-recurrence-': {
                'Recurrance': recur,
                'Recurrence Pattern': self.recurrencePattern,
            },
            '-status-': {
                'Meeting Status': meetingStatusString,
            },
            '-attendees-': {
                'Organizer': self.organizer,
                'Required Attendees': self.to,
                'Optional Attendees': self.cc,
                'Resources': self.bcc,
            },
            '-importance-': {
                'Importance': self.importanceString,
            },
        }

    @property
    def isMeeting(self) -> bool:
        """
        Attempts to determine if the object is a Meeting. True if meeting, False
        if appointment.
        """
        return self.appointmentStateFlags and AppointmentStateFlag.MEETING in self.appointmentStateFlags

    @property
    def originalStoreEntryID(self) -> Optional[EntryID]:
        """
        The EntryID of the delegator's message store.
        """
        return self._ensureSetNamed('_originalStoreEntryID', '8237', constants.PSETID_APPOINTMENT, overrideClass = EntryID.autoCreate)
