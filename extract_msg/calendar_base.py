import datetime

from typing import List, Set, Tuple, Union

from .enums import AppointmentAuxilaryFlag, AppointmentColor, AppointmentStateFlag, BusyStatus, IconIndex, ResponseStatus
from .message_base import MessageBase
from .structures.entry_id import EntryID
from .structures.misc_id import GlobalObjectID
from .structures.recurrence_pattern import RecurrencePattern
from .structures.time_zone_definition import TimeZoneDefinition
from .structures.time_zone_struct import TimeZoneStruct


class CalendarBase(MessageBase):
    """
    Common base for all Appointment and Meeting objects.
    """

    @property
    def allAttendeesString(self) -> str:
        """
        A list of all attendees, excluding the organizer.
        """
        return self._ensureSetNamed('_allAttendeesString', '8238')

    @property
    def appointmentAuxilaryFlags(self) -> Set[AppointmentAuxilaryFlag]:
        """
        The auxiliary state of the object.
        """
        return self._ensureSetNamed('_appointmentAuxilaryFlags', '8207', overrideClass = AppointmentAuxilaryFlag.fromBits)

    @property
    def appointmentColor(self) -> AppointmentColor:
        """
        The color to be used when displaying a Calendar object.
        """
        return self._ensureSetNamed('_appointmentColor', '8214', overrideClass = AppointmentColor)

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
    def appointmentNotAllowPropose(self) -> bool:
        """
        Indicates that attendees are not allowed to propose a new date and/or
        time for the meeting if True.
        """
        return self._ensureSetNamed('_appointmentNotAllowPropose', '8259')

    @property
    def appointmentRecur(self) -> RecurrencePattern:
        """
        Specifies the dates and times when a recurring series occurs by using
        one of the recurrence patterns and ranges specified in this section.
        """
        return self._ensureSetNamed('_appointmentRecur', '8216', overrideClass = RecurrencePattern)

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
    def appointmentTimeZoneDefinitionEndDisplay(self) -> TimeZoneDefinition:
        """
        Specifies the time zone information for the appointmentEndWhole property
        Used to convert the end date and time to and from UTC.
        """
        return self._ensureSetNamed('_appointmentTimeZoneDefinitionEndDisplay', '825F', overrideClass = TimeZoneDefinition)

    @property
    def appointmentTimeZoneDefinitionRecur(self) -> TimeZoneDefinition:
        """
        Specified the time zone information that specifies how to convert the
        meeting date and time on a recurring series to and from UTC.
        """
        return self._ensureSetNamed('_appointmentTimeZoneDefinitionRecur', '8260', overrideClass = TimeZoneDefinition)

    @property
    def appointmentTimeZoneDefinitionStartDisplay(self) -> TimeZoneDefinition:
        """
        Specifies the time zone information for the appointmentStartWhole
        property. Used to convert the start date and time to and from UTC.
        """
        return self._ensureSetNamed('_appointmentTimeZoneDefinitionStartDisplay', '825E', overrideClass = TimeZoneDefinition)

    @property
    def appointmentUnsendableRecipients(self) -> bytes:
        """
        A list of unsendable attendees.

        I want to return the structure parsed, but my one example does not match
        the specifications. If you have examples, let me know and I can ask you
        to run a verification on it.
        """
        return self._ensureSetNamed('_appointmentUnsendableRecipients', '825D')

    @property
    def birthdayContactAttributionDisplayName(self) -> str:
        """
        Indicated the name of the contact associated with the birthday event.
        """
        return self._ensureSetNamed('_birthdayContactAttributionDisplayName', 'BirthdayContactAttributionDisplayName')

    @property
    def birthdayContactEntryID(self) -> EntryID:
        """
        Indicates the EntryID of the contact associated with the birthday event.
        """
        return self._ensureSetNamed('_birthdayContactEntryID', 'BirthdayContactEntryId', overrideClass = EntryID.autoCreate)

    @property
    def birthdayContactPersonGuid(self) -> bytes:
        """
        Indicates the person ID's GUID of the contact associated with the
        birthday event.
        """
        return self._ensureSetNamed('_birthdayContactPersonGuid', 'BirthdayContactPersonGuid')

    @property
    def busyStatus(self) -> BusyStatus:
        """
        Specified the availability of a user for the event described by the
        object.
        """
        return self._ensureSetNamed('_busyStatus', '8205', overrideClass = BusyStatus)

    @property
    def ccAttendeesString(self) -> str:
        """
        A list of all the sendable attendees, who are also optional attendees.
        """
        return self._ensureSetNamed('_ccAttendeesString', '823C')

    @property
    def cleanGlobalObjectID(self) -> GlobalObjectID:
        """
        The value of the globalObjectID property for an object that represents
        an Exception object to a recurring series, where the year, month, and
        day fields are all 0.
        """
        return self._ensureSetNamed('_globalObjectID', '0023', overrideClass = GlobalObjectID)

    @property
    def clipEnd(self) -> datetime.datetime:
        """
        For single-instance Calendar objects, the end date and time of the
        event in UTC. For a recurring series, midnight in the user's machine
        time zone, on the date of the last instance, then is persisted in UTC,
        unless the recurring series has no end, in which case the value MUST be
        "31 August 4500 11:49 PM".

        Honestly, not sure what this is. [MS-OXOCAL]: PidLidClipEnd.
        """
        return self._ensureSetNamed('_clipStart', '8236')

    @property
    def clipStart(self) -> datetime.datetime:
        """
        For single-instance Calendar objects, the start date and time of the
        event in UTC. For a recurring series, midnight in the user's machine
        time zone, on the date of the first instance, then is persisted in UTC.

        Honestly, not sure what this is. [MS-OXOCAL]: PidLidClipStart.
        """
        return self._ensureSetNamed('_clipStart', '8235')

    @property
    def commonEnd(self) -> datetime.datetime:
        """
        The end date and time of an event. MUST be equal to appointmentEndWhole.
        """
        return self._ensureSetNamed('_commonEnd', '8517')

    @property
    def commonStart(self) -> datetime.datetime:
        """
        The start date and time of an event. MUST be equal to
        appointmentStartWhole.
        """
        return self._ensureSetNamed('_commonStart', '8516')

    @property
    def endDate(self) -> datetime.datetime:
        """
        The end date of the appointment.
        """
        return self._ensureSetProperty('_endDate', '00610040')

    @property
    def globalObjectID(self) -> GlobalObjectID:
        """
        The unique identifier or the Calendar object.
        """
        return self._ensureSetNamed('_globalObjectID', '0023', overrideClass = GlobalObjectID)

    @property
    def iconIndex(self) -> Union[IconIndex, int]:
        """
        The icon to use for the object.
        """
        return self._ensureSetProperty('_iconIndex', '10800003', overrideClass = IconIndex.tryMake)

    @property
    def isBirthdayContactWritable(self) -> bool:
        """
        Indicates whether the contact associated with the birthday event is
        writable.
        """
        return self._ensureSetNamed('_isBirthdayContactWritable', 'IsBirthdayContactWritable')

    @property
    def isException(self) -> bool:
        """
        Whether the object represents an exception. False indicates that the
        object represents a recurring series or a single-instance object.
        """
        return self._ensureSetNamed('_isException', '000A')

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
    def linkedTaskItems(self) -> Tuple[EntryID]:
        """
        A list of PidTagEntryId properties of Task objects related to the
        Calendar object that are set by a client.
        """
        return self._ensureSetNamed('_linkedTaskItems', '820C', overrideClass = lambda x : tuple(EntryID.autoCreate(y) for y in x))

    @property
    def location(self) -> str:
        """
        Returns the location of the meeting.
        """
        return self._ensureSetNamed('_location', '8208')

    @property
    def meetingDoNotForward(self) -> bool:
        """
        Whether to allow the meeting to be forwarded. True disallows forwarding.
        """
        return self._ensureSetNamed('_meetingDoNotForward', 'DoNotForward')

    @property
    def meetingWorkspaceUrl(self) -> str:
        """
        The URL of the Meeting Workspace, as specified in [MS-MEETS], that is
        associated with a Calendar object.
        """
        return self._ensureSetNamed('_meetingWorkspaceUrl', '8209')

    @property
    def nonSendableBcc(self) -> str:
        """
        A list of all unsendable attendees who are also resource objects.
        """
        return self._ensureSetNamed('_nonSendableBcc', '8538')

    @property
    def nonSendableCc(self) -> str:
        """
        A list of all unsendable attendees who are also optional attendees.
        """
        return self._ensureSetNamed('_nonSendableCc', '8537')

    @property
    def nonSendableTo(self) -> str:
        """
        A list of all unsendable attendees who are also required attendees.
        """
        return self._ensureSetNamed('_nonSendableTo', '8536')

    @property
    def nonSendBccTrackStatus(self) -> List[ResponseStatus]:
        """
        A ResponseStatus for each of the attendees in nonSendableBcc.
        """
        return self._ensureSetNamed('_nonSendBccTrackStatus', '8545', overrideClass = (lambda x : (ResponseStatus(y) for y in x)))

    @property
    def nonSendCcTrackStatus(self) -> List[ResponseStatus]:
        """
        A ResponseStatus for each of the attendees in nonSendableCc.
        """
        return self._ensureSetNamed('_nonSendCcTrackStatus', '8544', overrideClass = (lambda x : (ResponseStatus(y) for y in x)))

    @property
    def nonSendToTrackStatus(self) -> List[ResponseStatus]:
        """
        A ResponseStatus for each of the attendees in nonSendableTo.
        """
        return self._ensureSetNamed('_nonSendToTrackStatus', '8543', overrideClass = (lambda x : (ResponseStatus(y) for y in x)))

    @property
    def optionalAttendees(self) -> str:
        """
        Returns the optional attendees of the meeting.
        """
        return self._ensureSetNamed('_optionalAttendees', '0007')

    @property
    def organizer(self) -> str:
        """
        The meeting organizer.
        """
        return self._ensureSet('_organizer', '__substg1.0_0042')

    @property
    def ownerAppointmentID(self) -> int:
        """
        A quasi-unique value amond all Calendar objects in a user's mailbox.
        Assists a client or server in finding a Calendar object but is not
        guarenteed to be unique amoung all objects.
        """
        return self._ensureSetProperty('_ownerAppointmentID', '00620003')

    @property
    def ownerCriticalChange(self) -> datetime.datetime:
        """
        The date and time at which a Meeting Request object was sent by the
        organizer, in UTC.
        """
        return self._ensureSetNamed('_ownerCriticalChange', '001A')

    @property
    def recurrencePattern(self) -> str:
        """
        A description of the recurrence specified by the appointmentRecur
        property.
        """
        return self._ensureSetNamed('_recurrencePattern', '8232')

    @property
    def recurring(self) -> bool:
        """
        Specifies whether the object represents a recurring series.
        """
        return self._ensureSetNamed('_recurring', '8223', overrideClass = bool)

    @property
    def replyRequested(self) -> bool:
        """
        Whether the organizer requests a reply from attendees.
        """
        return self._ensureSetProperty('_replyRequested', '0C17000B')

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
    def responseRequested(self) -> bool:
        """
        Whether to send Meeting Response objects to the organizer.
        """
        return self._ensureSetProperty('_responseRequested', '0063000B')

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
    def timeZoneDescription(self) -> str:
        """
        A human-readable description of the time zone that is represented by the
        data in the timeZoneStruct property.
        """
        return self._ensureSetNamed('_timeZoneDescription', '8234')

    @property
    def timeZoneStruct(self) -> TimeZoneStruct:
        """
        Set on a recurring series to specify time zone information. Specifies
        how to convert time fields between local time and UTC.
        """
        return self._ensureSetNamed('_timeZoneStruct', '8233', overrideClass = TimeZoneStruct)

    @property
    def toAttendeesString(self) -> str:
        """
        A list of all the sendable attendees, who are also required attendees.
        """
        return self._ensureSetNamed('_toAttendeesString', '823B')
