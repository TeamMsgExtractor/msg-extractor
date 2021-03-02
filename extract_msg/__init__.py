#!/usr/bin/env python
# -*- coding: latin-1 -*-
# Date Format: YYYY-MM-DD

"""
extract_msg:
    Extracts emails and attachments saved in Microsoft Outlook's .msg files.

https://github.com/mattgwwalker/msg-extractor
"""

# --- LICENSE.txt -----------------------------------------------------------------
#
#    Copyright 2013 Matthew Walker
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
__date__ = '2021-03-02'
__version__ = '0.28.7'

import logging

from extract_msg import constants
from extract_msg.appointment import Appointment
from extract_msg.attachment import Attachment
from extract_msg.contact import Contact
from extract_msg.exceptions import UnrecognizedMSGTypeError
from extract_msg.message import Message
from extract_msg.message_base import MessageBase
from extract_msg.msg import MSGFile
from extract_msg.prop import create_prop
from extract_msg.properties import Properties
from extract_msg.recipient import Recipient
from extract_msg.utils import openMsg, properHex
