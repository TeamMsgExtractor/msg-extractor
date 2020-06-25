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

__author__ = 'Matthew Walker & The Elemental of Creation'
__date__ = '2020-06-18'
__version__ = '0.24.4'

import logging

from extract_msg import constants
from extract_msg.attachment import Attachment
from extract_msg.contact import Contact
from extract_msg.exceptions import UnrecognizedMSGTypeError
from extract_msg.message import Message
from extract_msg.msg import MSGFile
from extract_msg.prop import create_prop
from extract_msg.properties import Properties
from extract_msg.recipient import Recipient
from extract_msg.utils import parse_type, properHex



logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

def openMsg(path, prefix = '', attachmentClass = Attachment, filename = None, strict = True):
    """
    Function to automatically open an MSG file and detect what type it is.

    :param path: path to the msg file in the system or is the raw msg file.
    :param prefix: used for extracting embeded msg files
        inside the main one. Do not set manually unless
        you know what you are doing.
    :param attachmentClass: optional, the class the Message object
        will use for attachments. You probably should
        not change this value unless you know what you
        are doing.
    :param filename: optional, the filename to be used by default when saving.

    If :param strict: is set to `True`, this function will raise an exception
    when it cannot identify what MSGFile derivitive to use. Otherwise, it will
    log the error and return a basic MSGFile instance.
    """
    msg = MSGFile(path, prefix, attachmentClass, filename)
    if msg.classType.startswith('IPM.Contact'):
        return Contact(path, prefix, attachmentClass, filename)
    elif msg.classType.startswith('IPM.Note'):
        return Message(path, prefix, attachmentClass, filename)
    elif strict:
        raise UnrecognizedMSGTypeError('Could not recognize msg class type "{}". It is recommended you report this to the developers.'.format(msg.classType))
    else:
        logger.error('Could not recognize msg class type "{}". It is recommended you report this to the developers.'.format(msg.classType))
        return msg
