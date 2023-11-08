__all__ = [
    'MeetingRequest',
]


import datetime
import functools
import json

from typing import Optional

from ..constants import HEADER_FORMAT_TYPE, ps
from .meeting_related import MeetingRelated
from ..enums import BusyStatus, MeetingObjectChange, MeetingType, RecurCalendarType, RecurPatternType, ResponseStatus


class MeetingRequest(MeetingRelated):
    """
    Class for handling Meeting Request and Meeting Update objects.
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
    def appointmentMessageClass(self) -> Optional[str]:
        """
        Indicates the value of the PidTagMessageClass property of the Meeting
        object that is to be generated from the Meeting Request object. MUST
        start with "IPM.Appointment".
        """
        return self.getNamedProp('0024', ps.PSETID_MEETING)

    @functools.cached_property
    def calendarType(self) -> Optional[RecurCalendarType]:
        """
        The value of the CalendarType field from the PidLidAppointmentRecur
        property if the Meeting Request object represents a recurring series or
        an exception.
        """
        return self.getNamedAs('001C', ps.PSETID_MEETING, RecurCalendarType)

    @functools.cached_property
    def changeHighlight(self) -> Optional[MeetingObjectChange]:
        """
        Soecifies a bit field that indicates how the Meeting object has been
        changed.

        Returns a union of the set flags.
        """
        return self.getNamedAs('8204', ps.PSETID_APPOINTMENT, MeetingObjectChange)

    @functools.cached_property
    def forwardInstance(self) -> bool:
        """
        Indicates that the Meeting Request object represents an exception to a
        recurring series, and it was forwarded (even when forwarded by the
        organizer) rather than being an invitation sent by the organizer.
        """
        return bool(self.getNamedProp('820A', ps.PSETID_APPOINTMENT))

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

        showTime = None if self.appointmentNotAllowPropose else 'Tentative'

        return {
            '-main info-': {
                'Subject': self.subject,
                'Location': self.location,
            },
            '-date-': {
                'Start': self.startDate.__format__(self.datetimeFormat) if self.startDate else None,
                'End': self.endDate.__format__(self.datetimeFormat) if self.endDate else None,
                'Show Time As': showTime,
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
    def intendedBusyStatus(self) -> Optional[BusyStatus]:
        """
        The value of the busyStatus on the Meeting object in the organizer's
        calendar at the time the Meeting Request object or Meeting Update object
        was sent.
        """
        return self.getNamedAs('8224', ps.PSETID_APPOINTMENT, BusyStatus)

    @functools.cached_property
    def meetingType(self) -> Optional[MeetingType]:
        """
        The type of Meeting Request object or Meeting Update object.
        """
        return self.getNamedAs('0026', ps.PSETID_MEETING, MeetingType)

    @functools.cached_property
    def oldLocation(self) -> Optional[str]:
        """
        The original value of the location property before a meeting update.
        """
        return self.getNamedProp('0028', ps.PSETID_MEETING)

    @functools.cached_property
    def oldWhenEndWhole(self) -> Optional[datetime.datetime]:
        """
        The original value of the appointmentEndWhole property before a meeting
        update.
        """
        return self.getNamedProp('002A', ps.PSETID_MEETING)

    @functools.cached_property
    def oldWhenStartWhole(self) -> Optional[datetime.datetime]:
        """
        The original value of the appointmentStartWhole property before a
        meeting update.
        """
        return self.getNamedProp('0029', ps.PSETID_MEETING)
