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
__date__ = '2021-07-13'
__version__ = '0.29.0'

import logging

from . import constants
from .appointment import Appointment
from .attachment import Attachment
from .contact import Contact
from .exceptions import UnrecognizedMSGTypeError
from .message import Message
from .message_base import MessageBase
from .msg import MSGFile
from .prop import createProp
from .properties import Properties
from .recipient import Recipient
from .utils import openMsg, properHex
