__all__ = [
    'CalendarBase',
]


import datetime
import enum
import functools
import logging

from typing import List, Optional, Type, Union

from ..constants import ps
from ..enums import AppointmentAuxilaryFlag, AppointmentColor, AppointmentStateFlag, BusyStatus, IconIndex, MeetingRecipientType, ResponseStatus
from .message_base import MessageBase
from ..structures.entry_id import EntryID
from ..structures.misc_id import GlobalObjectID
from ..structures.recurrence_pattern import RecurrencePattern
from ..structures.time_zone_definition import TimeZoneDefinition
from ..structures.time_zone_struct import TimeZoneStruct


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class CalendarBase(MessageBase):
    """
    Common base for all Appointment and Meeting objects.
    """

    @functools.cached_property
    def allAttendeesString(self) -> Optional[str]:
        """
        A list of all attendees, excluding the organizer.
        """
        return self.getNamedProp('8238', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def appointmentAuxilaryFlags(self) -> Optional[AppointmentAuxilaryFlag]:
        """
        The auxiliary state of the object.
        """
        return self.getNamedAs('8207', ps.PSETID_APPOINTMENT, AppointmentAuxilaryFlag)

    @functools.cached_property
    def appointmentColor(self) -> Optional[AppointmentColor]:
        """
        The color to be used when displaying a Calendar object.
        """
        return self.getNamedAs('8214', ps.PSETID_APPOINTMENT, AppointmentColor)

    @functools.cached_property
    def appointmentDuration(self) -> Optional[int]:
        """
        The length of the event, in minutes.
        """
        return self.getNamedProp('8213', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def appointmentEndWhole(self) -> Optional[datetime.datetime]:
        """
        The end date and time of the event in UTC.
        """
        return self.getNamedProp('820E', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def appointmentNotAllowPropose(self) -> bool:
        """
        Indicates that attendees are not allowed to propose a new date and/or
        time for the meeting if True.
        """
        return bool(self.getNamedProp('8259', ps.PSETID_APPOINTMENT))

    @functools.cached_property
    def appointmentRecur(self) -> Optional[RecurrencePattern]:
        """
        Specifies the dates and times when a recurring series occurs by using
        one of the recurrence patterns and ranges specified in this section.
        """
        return self.getNamedAs('8216', ps.PSETID_APPOINTMENT, RecurrencePattern)

    @functools.cached_property
    def appointmentSequence(self) -> Optional[int]:
        """
        Specified the sequence number of a Meeting object.

        A meeting object begins with the sequence number set to 0 and is
        incremented each time the organizer sends out a Meeting Update object.
        """
        return self.getNamedProp('8201', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def appointmentStartWhole(self) -> Optional[datetime.datetime]:
        """
        The start date and time of the event in UTC.
        """
        return self.getNamedProp('820D', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def appointmentStateFlags(self) -> Optional[AppointmentStateFlag]:
        """
        The appointment state of the object.
        """
        return self.getNamedAs('8217', ps.PSETID_APPOINTMENT, AppointmentStateFlag)

    @functools.cached_property
    def appointmentSubType(self) -> bool:
        """
        Whether the event is an all-day event or not.
        """
        return bool(self.getNamedProp('8215', ps.PSETID_APPOINTMENT))

    @functools.cached_property
    def appointmentTimeZoneDefinitionEndDisplay(self) -> Optional[TimeZoneDefinition]:
        """
        Specifies the time zone information for the appointmentEndWhole property
        Used to convert the end date and time to and from UTC.
        """
        return self.getNamedAs('825F', ps.PSETID_APPOINTMENT, TimeZoneDefinition)

    @functools.cached_property
    def appointmentTimeZoneDefinitionRecur(self) -> Optional[TimeZoneDefinition]:
        """
        Specified the time zone information that specifies how to convert the
        meeting date and time on a recurring series to and from UTC.
        """
        return self.getNamedAs('8260', ps.PSETID_APPOINTMENT, TimeZoneDefinition)

    @functools.cached_property
    def appointmentTimeZoneDefinitionStartDisplay(self) -> Optional[TimeZoneDefinition]:
        """
        Specifies the time zone information for the appointmentStartWhole
        property.

        Used to convert the start date and time to and from UTC.
        """
        return self.getNamedAs('825E', ps.PSETID_APPOINTMENT, TimeZoneDefinition)

    @functools.cached_property
    def appointmentUnsendableRecipients(self) -> Optional[bytes]:
        """
        A list of unsendable attendees.

        Currently the raw bytes for the structure, but will return a usable
        class in the future.
        """
        return self.getNamedProp('825D', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def bcc(self) -> Optional[str]:
        """
        Returns the bcc field, if it exists.
        """
        return self._genRecipient('bcc', MeetingRecipientType.SENDABLE_RESOURCE_OBJECT)

    @functools.cached_property
    def birthdayContactAttributionDisplayName(self) -> Optional[str]:
        """
        Indicated the name of the contact associated with the birthday event.
        """
        return self.getNamedProp('BirthdayContactAttributionDisplayName', ps.PSETID_ADDRESS)

    @functools.cached_property
    def birthdayContactEntryID(self) -> Optional[EntryID]:
        """
        Indicates the EntryID of the contact associated with the birthday event.
        """
        return self.getNamedAs('BirthdayContactEntryId', ps.PSETID_ADDRESS, EntryID.autoCreate)

    @functools.cached_property
    def birthdayContactPersonGuid(self) -> Optional[bytes]:
        """
        Indicates the person ID's GUID of the contact associated with the
        birthday event.
        """
        return self.getNamedProp('BirthdayContactPersonGuid', ps.PSETID_ADDRESS)

    @functools.cached_property
    def busyStatus(self) -> Optional[BusyStatus]:
        """
        Specified the availability of a user for the event described by the
        object.
        """
        return self.getNamedAs('8205', ps.PSETID_APPOINTMENT, BusyStatus)

    @functools.cached_property
    def cc(self) -> Optional[str]:
        """
        Returns the cc field, if it exists.
        """
        return self._genRecipient('cc', MeetingRecipientType.SENDABLE_OPTIONAL_ATTENDEE)

    @functools.cached_property
    def ccAttendeesString(self) -> Optional[str]:
        """
        A list of all the sendable attendees, who are also optional attendees.
        """
        return self.getNamedProp('823C', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def cleanGlobalObjectID(self) -> Optional[GlobalObjectID]:
        """
        The value of the globalObjectID property for an object that represents
        an Exception object to a recurring series, where the year, month, and
        day fields are all 0.
        """
        return self.getNamedAs('0023', ps.PSETID_MEETING, GlobalObjectID)

    @functools.cached_property
    def clipEnd(self) -> Optional[datetime.datetime]:
        """
        For single-instance Calendar objects, the end date and time of the
        event in UTC.

        For a recurring series, midnight in the user's machine time zone, on
        the date of the last instance, then is persisted in UTC, unless the
        recurring series has no end, in which case the value MUST be
        "31 August 4500 11:49 PM".

        Honestly, not sure what this is. [MS-OXOCAL]: PidLidClipEnd.
        """
        return self.getNamedProp('8236', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def clipStart(self) -> Optional[datetime.datetime]:
        """
        For single-instance Calendar objects, the start date and time of the
        event in UTC.

        For a recurring series, midnight in the user's machine time zone, on
        the date of the first instance, then is persisted in UTC.

        Honestly, not sure what this is. [MS-OXOCAL]: PidLidClipStart.
        """
        return self.getNamedProp('8235', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def commonEnd(self) -> Optional[datetime.datetime]:
        """
        The end date and time of an event.

        MUST be equal to appointmentEndWhole.
        """
        return self.getNamedProp('8517', ps.PSETID_COMMON)

    @functools.cached_property
    def commonStart(self) -> Optional[datetime.datetime]:
        """
        The start date and time of an event.

        MUST be equal to appointmentStartWhole.
        """
        return self.getNamedProp('8516', ps.PSETID_COMMON)

    @functools.cached_property
    def endDate(self) -> Optional[datetime.datetime]:
        """
        The end date of the appointment.
        """
        return self.getPropertyVal('00610040')

    @functools.cached_property
    def globalObjectID(self) -> Optional[GlobalObjectID]:
        """
        The unique identifier or the Calendar object.
        """
        return self.getNamedAs('0003', ps.PSETID_MEETING, GlobalObjectID)

    @functools.cached_property
    def iconIndex(self) -> Optional[Union[IconIndex, int]]:
        """
        The icon to use for the object.
        """
        return self.getPropertyAs('10800003', IconIndex.tryMake)

    @functools.cached_property
    def isBirthdayContactWritable(self) -> bool:
        """
        Indicates whether the contact associated with the birthday event is
        writable.
        """
        return bool(self.getNamedProp('IsBirthdayContactWritable', ps.PSETID_ADDRESS))
    @functools.cached_property
    def isException(self) -> bool:
        """
        Whether the object represents an exception.

        ``False`` indicates that the object represents a recurring series or a
        single-instance object.
        """
        return bool(self.getNamedProp('000A', ps.PSETID_MEETING))

    @functools.cached_property
    def isRecurring(self) -> bool:
        """
        Whether the object is associated with a recurring series.
        """
        return bool(self.getNamedProp('0005', ps.PSETID_MEETING))

    @functools.cached_property
    def keywords(self) -> Optional[List[str]]:
        """
        The color to be used when displaying a Calendar object.
        """
        return self.getNamedProp('Keywords', ps.PS_PUBLIC_STRINGS)

    @functools.cached_property
    def linkedTaskItems(self) -> Optional[List[EntryID]]:
        """
        A list of PidTagEntryId properties of Task objects related to the
        Calendar object that are set by a client.
        """
        return self.getNamedAs('820C', ps.PSETID_APPOINTMENT, lambda x: list(EntryID.autoCreate(y) for y in x))

    @functools.cached_property
    def location(self) -> Optional[str]:
        """
        Returns the location of the meeting.
        """
        return self.getNamedProp('8208', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def meetingDoNotForward(self) -> bool:
        """
        Whether to allow the meeting to be forwarded.

        True disallows forwarding.
        """
        return bool(self.getNamedProp('DoNotForward', ps.PS_PUBLIC_STRINGS))

    @functools.cached_property
    def meetingWorkspaceUrl(self) -> Optional[str]:
        """
        The URL of the Meeting Workspace, as specified in [MS-MEETS], that is
        associated with a Calendar object.
        """
        return self.getNamedProp('8209', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def nonSendableBcc(self) -> Optional[str]:
        """
        A list of all unsendable attendees who are also resource objects.
        """
        return self.getNamedProp('8538', ps.PSETID_COMMON)

    @functools.cached_property
    def nonSendableCc(self) -> Optional[str]:
        """
        A list of all unsendable attendees who are also optional attendees.
        """
        return self.getNamedProp('8537', ps.PSETID_COMMON)

    @functools.cached_property
    def nonSendableTo(self) -> Optional[str]:
        """
        A list of all unsendable attendees who are also required attendees.
        """
        return self.getNamedProp('8536', ps.PSETID_COMMON)

    @functools.cached_property
    def nonSendBccTrackStatus(self) -> Optional[List[ResponseStatus]]:
        """
        A ResponseStatus for each of the attendees in nonSendableBcc.
        """
        return self.getNamedAs('8545', ps.PSETID_COMMON, ResponseStatus.fromIter)

    @functools.cached_property
    def nonSendCcTrackStatus(self) -> Optional[List[ResponseStatus]]:
        """
        A ResponseStatus for each of the attendees in nonSendableCc.
        """
        return self.getNamedAs('8544', ps.PSETID_COMMON, ResponseStatus.fromIter)

    @functools.cached_property
    def nonSendToTrackStatus(self) -> Optional[List[ResponseStatus]]:
        """
        A ResponseStatus for each of the attendees in nonSendableTo.
        """
        return self.getNamedAs('8543', ps.PSETID_COMMON, ResponseStatus.fromIter)

    @functools.cached_property
    def optionalAttendees(self) -> Optional[str]:
        """
        Returns the optional attendees of the meeting.
        """
        return self.getNamedProp('0007', ps.PSETID_MEETING)

    @property
    def organizer(self) -> Optional[str]:
        """
        The meeting organizer.
        """
        return self.getStringStream('__substg1.0_0042')

    @functools.cached_property
    def ownerAppointmentID(self) -> Optional[int]:
        """
        A quasi-unique value among all Calendar objects in a user's mailbox.
        Assists a client or server in finding a Calendar object but is not
        guaranteed to be unique among all objects.
        """
        return self.getPropertyVal('00620003')

    @functools.cached_property
    def ownerCriticalChange(self) -> Optional[datetime.datetime]:
        """
        The date and time at which a Meeting Request object was sent by the
        organizer, in UTC.
        """
        return self.getNamedProp('001A', ps.PSETID_MEETING)

    @property
    def recipientTypeClass(self) -> Type[enum.IntEnum]:
        return MeetingRecipientType

    @functools.cached_property
    def recurrencePattern(self) -> Optional[str]:
        """
        A description of the recurrence specified by the appointmentRecur
        property.
        """
        return self.getNamedProp('8232', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def recurring(self) -> bool:
        """
        Specifies whether the object represents a recurring series.
        """
        return bool(self.getNamedProp('8223', ps.PSETID_APPOINTMENT))

    @functools.cached_property
    def replyRequested(self) -> bool:
        """
        Whether the organizer requests a reply from attendees.
        """
        return bool(self.getPropertyVal('0C17000B'))

    @functools.cached_property
    def requiredAttendees(self) -> Optional[str]:
        """
        Returns the required attendees of the meeting.
        """
        return self.getNamedProp('0006', ps.PSETID_MEETING)

    @functools.cached_property
    def resourceAttendees(self) -> Optional[str]:
        """
        Returns the resource attendees of the meeting.
        """
        return self.getNamedProp('0008', ps.PSETID_MEETING)

    @functools.cached_property
    def responseStatus(self) -> ResponseStatus:
        """
        The response status of an attendee.
        """
        return ResponseStatus(self.getNamedProp('8218', ps.PSETID_APPOINTMENT, 0))

    @functools.cached_property
    def startDate(self) -> Optional[datetime.datetime]:
        """
        The start date of the appointment.
        """
        return self.getPropertyVal('00600040')

    @functools.cached_property
    def timeZoneDescription(self) -> Optional[str]:
        """
        A human-readable description of the time zone that is represented by the
        data in the timeZoneStruct property.
        """
        return self.getNamedProp('8234', ps.PSETID_APPOINTMENT)

    @functools.cached_property
    def timeZoneStruct(self) -> Optional[TimeZoneStruct]:
        """
        Set on a recurring series to specify time zone information.

        Specifies how to convert time fields between local time and UTC.
        """
        return self.getNamedAs('8233', ps.PSETID_APPOINTMENT, TimeZoneStruct)

    @functools.cached_property
    def to(self) -> Optional[str]:
        """
        Returns the to field, if it exists.
        """
        return self._genRecipient('to', MeetingRecipientType.SENDABLE_REQUIRED_ATTENDEE)

    @functools.cached_property
    def toAttendeesString(self) -> Optional[str]:
        """
        A list of all the sendable attendees, who are also required attendees.
        """
        return self.getNamedProp('823B', ps.PSETID_APPOINTMENT)
