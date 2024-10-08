"""
Classes for opening an MSG file with.
"""

__all__ = [
    # Classes:
    'AppointmentMeeting',
    'Calendar',
    'CalendarBase',
    'Contact',
    'Journal',
    'MeetingCancellation',
    'MeetingException',
    'MeetingForwardNotification',
    'MeetingRelated',
    'MeetingRequest',
    'MeetingResponse',
    'Message',
    'MessageBase',
    'MessageSigned',
    'MessageSignedBase',
    'MSGFile',
    'Post',
    'StickyNote',
    'Task',
    'TaskRequest',
]


from .appointment import AppointmentMeeting
from .calendar_base import CalendarBase
from .calendar import Calendar
from .contact import Contact
from .journal import Journal
from .meeting_cancellation import MeetingCancellation
from .meeting_exception import MeetingException
from .meeting_forward import MeetingForwardNotification
from .meeting_related import MeetingRelated
from .meeting_request import MeetingRequest
from .meeting_response import MeetingResponse
from .message import Message
from .message_base import MessageBase
from .message_signed import MessageSigned
from .message_signed_base import MessageSignedBase
from .msg import MSGFile
from .post import Post
from .sticky_note import StickyNote
from .task import Task
from .task_request import TaskRequest