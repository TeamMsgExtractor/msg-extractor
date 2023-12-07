__all__ = [
    'AppointmentMeeting',
]


import datetime
import functools
import json

from typing import Optional

from ..constants import HEADER_FORMAT_TYPE, ps
from .calendar import Calendar
from ..enums import AppointmentStateFlag, RecurPatternType, ResponseStatus
from ..structures.entry_id import EntryID


class AppointmentMeeting(Calendar):
    """
    Parser for Microsoft Outlook Appointment or Meeting files.

    Both Appointment and Meeting have the same class type but Meeting has
    additional properties. These properties are meaningless on an Appointment
    object.
    """

    def getJson(self) -> str:
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

        return json.dumps({
            'recurrence': recur,
            'recurrencePattern': self.recurrencePattern,
            'body': self.body,
            'meetingStatus': meetingStatusString,
            'organizer': self.organizer,
            'requiredAttendees': self.to,
            'optionalAttendees': self.cc,
            'resources': self.bcc,
            'start': self.startDate.__format__(self.datetimeFormat) if self.endDate else None,
            'end': self.endDate.__format__(self.datetimeFormat) if self.endDate else None,
        })

    @functools.cached_property
    def appointmentCounterProposal(self) -> bool:
        """
        Indicates to the organizer that there are counter proposals that have
        not been accepted or rejected by the organizer.
        """
        return bool(self.getNamedProp('8257', ps.PSETID_APPOINTMENT))

    @functools.cached_property
    def appointmentLastSequence(self) -> Optional[int]:
        """
        The last sequence number that was sent to any attendee.
        """
        return self.getNamedProp('8203', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def appointmentProposalNumber(self) -> Optional[int]:
        """
        The number of attendees who have sent counter propostals that have not
        been accepted or rejected by the organizer.
        """
        return self.getNamedProp('8259', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def appointmentReplyName(self) -> Optional[datetime.datetime]:
        """
        The user who last replied to the meeting request or meeting update.
        """
        return self.getNamedProp('8230', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def appointmentReplyTime(self) -> Optional[datetime.datetime]:
        """
        The date and time at which the attendee responded to a received Meeting
        Request object of Meeting Update object in UTC.
        """
        return self.getNamedProp('8220', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def appointmentSequenceTime(self) -> Optional[datetime.datetime]:
        """
        The date and time at which the appointmentSequence property was last
        modified.
        """
        return self.getNamedProp('8202', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def autoFillLocation(self) -> bool:
        """
        A value of True indicates that the value of the location property is set
        to the value of the displayName property from the recipientRow structure
        that represents a Resource object.

        A value of ``False`` indicates that the value of the location property
        is not automatically set.
        """
        return bool(self.getNamedProp('823A', ps.PSETID_APPOINTMENT))

    @functools.cached_property
    def fInvited(self) -> bool:
        """
        Whether a Meeting Request object has been sent out.
        """
        return bool(self.getNamedProp('8229', ps.PSETID_APPOINTMENT))

    @property
    def headerFormatProperties(self) -> HEADER_FORMAT_TYPE:
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
                'Start': self.startDate.__format__(self.datetimeFormat) if self.startDate else None,
                'End': self.endDate.__format__(self.datetimeFormat) if self.endDate else None,
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

    @functools.cached_property
    def isMeeting(self) -> bool:
        """
        Attempts to determine if the object is a Meeting.

        ``True`` if meeting, ``False`` if appointment.
        """
        return bool(self.appointmentStateFlags) and (AppointmentStateFlag.MEETING in self.appointmentStateFlags)

    @functools.cached_property
    def originalStoreEntryID(self) -> Optional[EntryID]:
        """
        The EntryID of the delegator's message store.
        """
        return self.getNamedAs('8237', ps.PSETID_APPOINTMENT, EntryID.autoCreate)
