#!/usr/bin/env python3
# -*- coding: latin-1 -*-
# Date Format: YYYY-MM-DD

"""
extract_msg:
    Extracts emails and attachments saved in Microsoft Outlook's .msg files.

https://github.com/TeamMsgExtractor/msg-extractor
"""

# --- LICENSE.txt --------------------------------------------------------------
#
#    Copyright 2013-2022 Matthew Walker and Destiny Peterson
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.

__author__ = 'Destiny Peterson & Matthew Walker'
__date__ = '2023-05-09'
__version__ = '0.41.0'

__all__ = [
    # Modules:
    'constants',
    'enums',
    'exceptions',

    # Classes:
    'AppointmentMeeting',
    'Attachment',
    'Contact',
    'MeetingForwardNotification',
    'MeetingRequest',
    'MeetingResponse',
    'Message',
    'MessageBase',
    'MessageSigned',
    'MessageSignedBase',
    'MSGFile',
    'Post',
    'Properties',
    'Recipient',
    'Task',

    #Functions:
    'createProp',
    'openMsg',
    'openMsgBulk',
]

from . import constants, enums, exceptions
from .appointment import AppointmentMeeting
from .attachment import Attachment
from .contact import Contact
from .meeting_forward import MeetingForwardNotification
from .meeting_request import MeetingRequest
from .meeting_response import MeetingResponse
from .message import Message
from .message_base import MessageBase
from .message_signed import MessageSigned
from .message_signed_base import MessageSignedBase
from .msg import MSGFile
from .post import Post
from .prop import createProp
from .properties import Properties
from .recipient import Recipient
from .task import Task
from .utils import openMsg, openMsgBulk
