#!/usr/bin/env python
# -*- coding: latin-1 -*-
# Date Format: YYYY-MM-DD

# --- LICENSE -----------------------------------------------------------------
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
__date__ = '2018-05-22'
__version__ = '0.20.3'

from extract_msg import constants
from extract_msg.__doc__ import __doc__
from extract_msg.attachment import Attachment
from extract_msg.debug import debug
from extract_msg.message import Message
from extract_msg.prop import Prop
from extract_msg.properties import Properties
from extract_msg.recipient import Recipient
from extract_msg.utils import parse_type, properHex
