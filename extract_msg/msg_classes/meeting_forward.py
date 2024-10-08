__all__ = [
    'MeetingForwardNotification',
]


import functools
import json

from typing import Optional

from .. import constants
from .meeting_related import MeetingRelated
from ..enums import RecurPatternType, ResponseStatus


class MeetingForwardNotification(MeetingRelated):
    """
    Class for handling Meeting Forward Notification objects.
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
    def forwardNotificationRecipients(self) -> Optional[bytes]:
        """
        Bytes containing a list of RecipientRow structures that indicate the
        recipients of a meeting forward.

        Incomplete, looks to be the same structure as
        appointmentUnsendableRecipients, so we need more examples of this.
        """
        return self.getNamedProp('8261', constants.ps.PSETID_APPOINTMENT)

    @property
    def headerFormatProperties(self) -> constants.HEADER_FORMAT_TYPE:
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
    def promptSendUpdate(self) -> bool:
        """
        Indicates that the Meeting Forward Notification object was out-of-date
        when it was received.
        """
        return bool(self.getNamedProp('8045', constants.ps.PSETID_COMMON))
