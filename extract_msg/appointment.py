from extract_msg import constants
from extract_msg.attachment import Attachment
from extract_msg.message_base import MessageBase

class Appointment(MessageBase):
    """
    Parser for Microsoft Outlook Appointment files.
    """

    def __init__(self, path, prefix = '', attachmentClass = Attachment, filename = None, delayAttachments = False, overrideEncoding = None, attachmentErrorBehavior = constants.ATTACHMENT_ERROR_THROW):
        MessageBase.__init__(self, path, prefix, attachmentClass, filename, delayAttachments, overrideEncoding, attachmentErrorBehavior)

    @property
    def appointmentClassType(self):
        """
        The class type of the appointment.
        """
        return self._ensureSetNamed('_appointmentClassType', '0024')

    @property
    def endDate(self):
        """
        The end date of the appointment.
        """
        return self._ensureSetProperty('_endDate', '00610040')

    @property
    def location(self):
        """
        Returns the location of the meeting.
        """
        try:
            return self.__location
        except AttributeError:
            self.__location = self.named.getNamedValue('8208')
            self.__location = self.named.getNamedValue('0002') if self.__location is None else self.__location
            return self.__location

    @property
    def optionalAttendees(self):
        """
        Returns the optional attendees of the meeting.
        """
        return self._ensureSetNamed('_optionalAttendees', '0007')

    @property
    def requiredAttendees(self):
        """
        Returns the required attendees of the meeting.
        """
        return self._ensureSetNamed('_requiredAttendees', '0006')

    @property
    def resourceAttendees(self):
        """
        Returns the resource attendees of the meeting.
        """
        return self._ensureSetNamed('_resourceAttendees', '0008')


    @property
    def startDate(self):
        """
        The start date of the appointment.
        """
        return self._ensureSetProperty('_startDate', '00600040')

    @property
    def timeZone(self):
        """
        Returns the time zone of the meeting.
        """
        return self._ensureSetNamed('_timeZone', '000C')
