__all__ = [
    'MeetingForwardNotification',
]


from typing import Optional

from . import constants
from .meeting_related import MeetingRelated
from .enums import RecurPatternType


class MeetingForwardNotification(MeetingRelated):
    """
    Class for handling Meeting Forward Notification objects.
    """

    @property
    def forwardNotificationRecipients(self) -> Optional[bytes]:
        """
        Bytes containing a list of RecipientRow structures that indicate the
        recipients of a meeting forward.

        Incomplete, looks to be the same structure as
        appointmentUnsendableRecipients, so we need more examples of this.
        """
        return self._ensureSetNamed('_forwardNotificationRecipients', '8261', constants.PSETID_APPOINTMENT)

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
    def promptSendUpdate(self) -> bool:
        """
        Indicates that the Meeting Forward Notification object was out-of-date
        when it was received.
        """
        return self._ensureSetNamed('_promptSendUpdate', '8045', constants.PSETID_COMMON, overrideClass = bool, preserveNone = False)
