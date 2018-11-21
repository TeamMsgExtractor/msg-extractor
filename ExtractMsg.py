#!/usr/bin/env python
# -*- coding: latin-1 -*-
# Date Format: YYYY-MM-DD
"""
ExtractMsg:
    Extracts emails and attachments saved in Microsoft Outlook's .msg files

https://github.com/mattgwwalker/msg-extractor
"""

__author__ = 'Matthew Walker & The Elemental of Creation'
__date__ = '2018-05-22'
__version__ = '0.20'

debug = False


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

import copy
import datetime
import email.utils
import glob
import json
import olefile as OleFile
import os
import pprint
import random
import re
import string
import struct
import sys
import traceback
import tzlocal
from email.parser import Parser as EmailParser
from imapclient.imapclient import decode_utf7

# DEFINE CONSTANTS
# WARNING DO NOT CHANGE ANY OF THESE VALUES UNLESS YOU KNOW
# WHAT YOU ARE DOING! FAILURE TO FOLLOW THIS INSTRUCTION
# CAN AND WILL BREAK THIS SCRIPT!

INTELLIGENCE_DUMB  = 0
INTELLIGENCE_SMART = 1
INTELLIGENCE_DICT = {
    INTELLIGENCE_DUMB:  'INTELLIGENCE_DUMB',
    INTELLIGENCE_SMART: 'INTELLIGENCE_SMART',
}

TYPE_MESSAGE       = 0
TYPE_MESSAGE_EMBED = 1
TYPE_ATTACHMENT    = 2
TYPE_RECIPIENT     = 3
TYPE_DICT = {
    TYPE_MESSAGE:       'TYPE_MESSAGE',
    TYPE_MESSAGE_EMBED: 'TYPE_MESSAGE_EMBED',
    TYPE_ATTACHMENT:    'TYPE_ATTACHMENT',
    TYPE_RECIPIENT:     'TYPE_RECIPIENT',
}


RECIPIENT_SENDER   = 0
RECIPIENT_TO       = 1
RECIPIENT_CC       = 2
RECIPIENT_BCC      = 3
RECIPIENT_DICT = {
    RECIPIENT_SENDER:   'RECIPIENT_SENDER',
    RECIPIENT_TO:       'RECIPIENT_TO',
    RECIPIENT_CC:       'RECIPIENT_CC',
    RECIPIENT_BCC:      'RECIPIENT_BCC',
}

# Define pre-compiled structs to make unpacking slightly faster
ST1                = struct.Struct('<8x4I')
ST2                = struct.Struct('<H2xI8s')
ST3                = struct.Struct('<Q')
STI16              = struct.Struct('<h6x')
STI32              = struct.Struct('<i4x')
STI64              = struct.Struct('<q')
STF32              = struct.Struct('<f4x')
STF64              = struct.Struct('<d')

# END CONSTANTS

# This property information was sourced from
# http://www.fileformat.info/format/outlookmsg/index.htm
# on 2013-07-22.
# It was extened by The Elemental of Creation on 2018-10-12
# properties = {
#     '00010102': 'Template data',
#     '0002000B': 'Alternate recipient allowed',
#     '0004001F': 'Auto forward comment',
#     '00040102': 'Script data',
#     '0005000B': 'Auto forwarded',
#     '000F000F': 'Deferred delivery time',
#     '00100040': 'Deliver time',
#     '00150040': 'Expiry time',
#     '00170003': 'Importance',
#     '001A001F': 'Message class',
#     '0023001F': 'Originator delivery report requested',
#     '00250102': 'Parent key',
#     '00260003': 'Priority',
#     '0029000B': 'Read receipt requested',
#     '002A0040': 'Receipt time',
#     '002B000B': 'Recipient reassignment prohibited',
#     '002E0003': 'Original sensitivity',
#     '00300040': 'Reply time',
#     '00310102': 'Report tag',
#     '00320040': 'Report time',
#     '00360003': 'Sensitivity',
#     '0037001F': 'Subject',
#     '00390040': 'Client Submit Time',
#     '003A001F': '',
#     '003B0102': '',
#     '003D001F': 'Subject prefix',
#     '003F0102': '',
#     '0040001F': 'Received by name',
#     '00410102': '',
#     '0042001F': 'Sent repr name',
#     '00430102': '',
#     '0044001F': 'Rcvd repr name',
#     '00450102': '',
#     '0046001F': '',
#     '00470102': '',
#     '0049001F': '',
#     '004B001F': '',
#     '004C0102': '',
#     '004D001F': 'Org author name',
#     '004E0040': '',
#     '004F0102': '',
#     '0050001F': 'Reply rcipnt names',
#     '00510102': '',
#     '00520102': '',
#     '00530102': '',
#     '00540102': '',
#     '00550040': '',
#     '0057000B': '',
#     '0058000B': '',
#     '0059000B': '',
#     '005A001F': 'Org sender name',
#     '005B0102': '',
#     '005C0102': '',
#     '005D001F': '',
#     '005E0102': '',
#     '005F0102': '',
#     '00600040': '',
#     '00610040': '',
#     '00620003': '',
#     '0063000B': '',
#     '0064001F': 'Sent repr adrtype',
#     '0065001F': 'Sent repr email',
#     '0066001F': '',
#     '00670102': '',
#     '0068001F': '',
#     '0069001F': '',
#     '0070001F': 'Topic',
#     '00710102': '',
#     '0072001F': '',
#     '0073001F': '',
#     '0074001F': '',
#     '0075001F': 'Rcvd by adrtype',
#     '0076001F': 'Rcvd by email',
#     '0077001F': 'Repr adrtype',
#     '0078001F': 'Repr email',
#     '007D001F': 'Message header',
#     '007F0102': '',
#     '0080001F': '',
#     '0081001F': '',
#     '08070003': '',
#     '0809001F': '',
#     '0C040003': '',
#     '0C050003': '',
#     '0C06000B': '',
#     '0C08000B': '',
#     '0C150003': '',
#     '0C17000B': '',
#     '0C190102': '',
#     '0C1A001F': 'Sender name',
#     '0C1B001F': '',
#     '0C1D0102': '',
#     '0C1E001F': 'Sender adr type',
#     '0C1F001F': 'Sender email',
#     '0C200003': '',
#     '0C21001F': '',
#     '0E01000B': '',
#     '0E02001F': 'Display BCC',
#     '0E03001F': 'Display CC',
#     '0E04001F': 'Display To',
#     '0E060040': '',
#     '0E070003': '',
#     '0E080003': '',
#     '0E080014': '',
#     '0E090102': '',
#     '0E0F000B': '',
#     '0E12000D': '',
#     '0E13000D': '',
#     '0E170003': '',
#     '0E1B000B': '',
#     '0E1D001F': 'Subject (normalized)',
#     '0E1F000B': '',
#     '0E200003': '',
#     '0E210003': '',
#     '0E28001F': 'Recvd account1 (uncertain)',
#     '0E29001F': 'Recvd account2 (uncertain)',
#     '1000001F': 'Message body',
#     '1008': 'RTF sync body tag', # Where did this come from ??? It's not listed in the docs
#     '10090102': 'Compressed RTF body',
#     '1013001F': 'HTML body',
#     '1035001F': 'Message ID (uncertain)',
#     '1046001F': 'Sender email (uncertain)',
#     '3001001F': 'Display name',
#     '3002001F': 'Address type',
#     '3003001F': 'Email address',
#     '30070040': 'Creation date',
#     '39FE001F': '7-bit email (uncertain)',
#     '39FF001F': '7-bit display name',
#
#     # Attachments (37xx)
#     '37010102': 'Attachment data',
#     '37020102': '',
#     '3703001F': 'Attachment extension',
#     '3704001F': 'Attachment short filename',
#     '37050003': 'Attachment attach method',
#     '3707001F': 'Attachment long filename',
#     '370E001F': 'Attachment mime tag',
#     '3712001F': 'Attachment ID (uncertain)',
#
#     # Address book (3Axx):
#     '3A00001F': 'Account',
#     '3A02001F': 'Callback phone no',
#     '3A05001F': 'Generation',
#     '3A06001F': 'Given name',
#     '3A08001F': 'Business phone',
#     '3A09001F': 'Home phone',
#     '3A0A001F': 'Initials',
#     '3A0B001F': 'Keyword',
#     '3A0C001F': 'Language',
#     '3A0D001F': 'Location',
#     '3A11001F': 'Surname',
#     '3A15001F': 'Postal address',
#     '3A16001F': 'Company name',
#     '3A17001F': 'Title',
#     '3A18001F': 'Department',
#     '3A19001F': 'Office location',
#     '3A1A001F': 'Primary phone',
#     '3A1B101F': 'Business phone 2',
#     '3A1C001F': 'Mobile phone',
#     '3A1D001F': 'Radio phone no',
#     '3A1E001F': 'Car phone no',
#     '3A1F001F': 'Other phone',
#     '3A20001F': 'Transmit dispname',
#     '3A21001F': 'Pager',
#     '3A220102': 'User certificate',
#     '3A23001F': 'Primary Fax',
#     '3A24001F': 'Business Fax',
#     '3A25001F': 'Home Fax',
#     '3A26001F': 'Country',
#     '3A27001F': 'Locality',
#     '3A28001F': 'State/Province',
#     '3A29001F': 'Street address',
#     '3A2A001F': 'Postal Code',
#     '3A2B001F': 'Post Office Box',
#     '3A2C001F': 'Telex',
#     '3A2D001F': 'ISDN',
#     '3A2E001F': 'Assistant phone',
#     '3A2F001F': 'Home phone 2',
#     '3A30001F': 'Assistant',
#     '3A44001F': 'Middle name',
#     '3A45001F': 'Dispname prefix',
#     '3A46001F': 'Profession',
#     '3A47001F': '',
#     '3A48001F': 'Spouse name',
#     '3A4B001F': 'TTYTTD radio phone',
#     '3A4C001F': 'FTP site',
#     '3A4E001F': 'Manager name',
#     '3A4F001F': 'Nickname',
#     '3A51001F': 'Business homepage',
#     '3A57001F': 'Company main phone',
#     '3A58101F': 'Childrens names',
#     '3A59001F': 'Home City',
#     '3A5A001F': 'Home Country',
#     '3A5B001F': 'Home Postal Code',
#     '3A5C001F': 'Home State/Provnce',
#     '3A5D001F': 'Home Street',
#     '3A5F001F': 'Other adr City',
#     '3A60': 'Other adr Country',
#     '3A61': 'Other adr PostCode',
#     '3A62': 'Other adr Province',
#     '3A63': 'Other adr Street',
#     '3A64': 'Other adr PO box',
#
#     '3FF7': 'Server (uncertain)',
#     '3FF8': 'Creator1 (uncertain)',
#     '3FFA': 'Creator2 (uncertain)',
#     '3FFC': 'To email (uncertain)',
#     '403D': 'To adrtype (uncertain)',
#     '403E': 'To email (uncertain)',
#     '5FF6': 'To (uncertain)'
# }

types = {
    0x0000: 'PtypUnspecified',
    0x0001: 'PtypNull',
    0x0002: 'PtypInteger16', # Signed short
    0x0003: 'PtypInteger32', # Signed int
    0x0004: 'PtypFloating32', # Float
    0x0005: 'PtypFloating64', # Double
    0x0006: 'PtypCurrency',
    0x0007: 'PtypFloatingTime',
    0x000A: 'PtypErrorCode',
    0x000B: 'PtypBoolean',
    0x000D: 'PtypObject/PtypEmbeddedTable',
    0x0014: 'PtypInteger64', # Signed longlong
    0x001E: 'PtypString8',
    0x001F: 'PtypString',
    0x0040: 'PtypTime', # Use msgEpoch to convert to unix time stamp
    0x0048: 'PtypGuid',
    0x00FB: 'PtypServerId',
    0x00FD: 'PtypRestriction',
    0x00FE: 'PtypRuleAction',
    0x0102: 'PtypBinary',
    0x1002: 'PtypMultipleInteger16',
    0x1003: 'PtypMultipleInteger32',
    0x1004: 'PtypMultipleFloating32',
    0x1005: 'PtypMultipleFloating64',
    0x1006: 'PtypMultipleCurrency',
    0x1007: 'PtypMultipleFloatingTime',
    0x1014: 'PtypMultipleInteger64',
    0x101E: 'PtypMultipleString8',
    0x101F: 'PtypMultipleString',
    0x1040: 'PtypMultipleTime',
    0x1048: 'PtypMultipleGuid',
    0x1102: 'PtypMultipleBinary',
}

if sys.version_info[0] >= 3: # Python 3
    def xstr(s):
        return '' if s is None else str(s)

    def windowsUnicode(string):
        if string is None:
            return None
        return str(string, 'utf_16_le')

    stri = (str,)

    def properHex(inp):
        """
        Taken (with permission) from https://github.com/TheElementalOfCreation/creatorUtils
        """
        a = ''
        if isinstance(inp, stri):
            a = ''.join([hex(ord(inp[x]))[2:].rjust(2, '0') for x in range(len(inp))])
        if isinstance(inp, bytes):
            a = inp.hex()
        elif isinstance(inp, int):
            a = hex(inp)[2:]
        if len(a)%2 != 0:
            a = '0' + a
        return a

    def encode(inp):
        return inp
else:  # Python 2
    def xstr(s):
        if isinstance(s, unicode):
            return s.encode('utf-8')
        else:
            return '' if s is None else str(s)

    def windowsUnicode(string):
        if string is None:
            return None
        return unicode(string, 'utf_16_le')

    stri = (str, unicode)

    def properHex(inp):
        """
        Taken (with permission) from https://github.com/TheElementalOfCreation/creatorUtils
        """
        a = ''
        if isinstance(inp, stri):
            a = ''.join([hex(ord(inp[x]))[2:].rjust(2, '0') for x in range(len(inp))])
        elif isinstance(inp, int):
            a = hex(inp)[2:]
        elif isinstance(inp, long):
            a = hex(inp)[2:-1]
        if len(a)%2 != 0:
            a = '0' + a
        return a
    def encode(inp):
        return inp.encode('utf-8')

def int_to_recipient_type(integer):
    return RECIPIENT_DICT[integer]

def int_to_data_type(integer):
    return TYPE_DICT[integer]

def int_to_intelligence(integer):
    return INTELLIGENCE_DICT[integer]

def parse_type(type, stream):
    """
    Converts the data in :param stream: to a
    much more accurate type, specified by
    :param type:, if possible.

    Some types require that :param prop_value: be specified. This can be retrieved from the Properties instance.

    WARNING: Not done. Do not try to implement anywhere where it is not already implemented
    """
    #WARNING Not done. Do not try to implement anywhere where it is not already implemented
    value = stream
    if type == 0x0000: #PtypUnspecified
        pass;
    elif type == 0x0001: #PtypNull
        if value != b'\x00\x00\x00\x00\x00\x00\x00\x00':
            print('Warning: Property type is PtypNull, but is not equal to 0.')
        value = None
    elif type == 0x0002: #PtypInteger16
        value = STI16.unpack(value)[0]
    elif type == 0x0003: #PtypInteger32
        value = STI32.unpack(value)[0]
    elif type == 0x0004: #PtypFloating32
        value = STF32.unpack(value)[0]
    elif type == 0x0005: #PtypFloating64
        value = STF64.unpack(value)[0]
    elif type == 0x0006: #PtypCurrency
        value = (STI64.unpack(value)[0])/10000.0
    elif type == 0x0007: #PtypFloatingTime
        value = STF64.unpack(value)[0]
        #TODO parsing for this
        pass;
    elif type == 0x000A: #PtypErrorCode
        value = STI32.unpack(value)[0]
        #TODO parsing for this
        pass;
    elif type == 0x000B: #PtypBoolean
        value = bool(ST3.unpack(value)[0])
    elif type == 0x000D: #PtypObject/PtypEmbeddedTable
        #TODO parsing for this
        pass;
    elif type == 0x0014: #PtypInteger64
        value = STI64.unpack(value)[0]
    elif type == 0x001E: #PtypString8
        #TODO parsing for this
        pass;
    elif type == 0x001F: #PtypString
        value = value.decode('utf_16_le')
    elif type == 0x0040: #PtypTime
        value = ST3.unpack(value)[0]
    elif type == 0x0048: #PtypGuid
        #TODO parsing for this
        pass;
    elif type == 0x00FB: #PtypServerId
        #TODO parsing for this
        pass;
    elif type == 0x00FD: #PtypRestriction
        #TODO parsing for this
        pass;
    elif type == 0x00FE: #PtypRuleAction
        #TODO parsing for this
        pass;
    elif type == 0x0102: #PtypBinary
        #TODO parsing for this
        # Smh, how on earth am I going to code this???
        pass;
    elif type & 0x1000 == 0x1000: #PtypMultiple
        #TODO parsing for `multiple` types
        pass;
    return value;

def msgEpoch(inp):
    """
    Taken (with permission) from https://github.com/TheElementalOfCreation/creatorUtils
    """
    ep = 116444736000000000
    return (inp - ep)/10000000.0

def divide(string, length):
    """
    Taken (with permission) from https://github.com/TheElementalOfCreation/creatorUtils

    Divides a string into multiple substrings of equal length
    :param string: string to be divided.
    :param length: length of each division.
    :returns: list containing the divided strings.

    Example:
    >>>> a = divide('Hello World!', 2)
    >>>> print(a)
    ['He', 'll', 'o ', 'Wo', 'rl', 'd!']
    """
    return [string[length*x:length*(x+1)] for x in range(int(len(string)/length))]

def has_len(obj):
    """
    Checks if :param obj: has a __len__ attribute.
    """
    try:
        obj.__len__
        return True
    except:
        return False

def addNumToDir(dirName):
    """
    Attempt to create the directory with a '(n)' appended
    """
    for i in range(2, 100):
        try:
            newDirName = dirName + ' (' + str(i) + ')'
            os.makedirs(newDirName)
            return newDirName
        except Exception as e:
            pass
    return None

def fromTimeStamp(stamp):
    return datetime.datetime.fromtimestamp(stamp, tzlocal.get_localzone())

class Attachment:
    """
    Stores the attachment data of a Message instance.
    Should the attachment be an embeded message, the
    class used to create it will be the same as the
    Message class used to create the attachment.
    """
    def __init__(self, msg, dir_):
        """
        :param msg: the Message instance that the attachment belongs to.
        :param dir_: the directory inside the msg file where the attachment is located.
        """
        self.__msg = msg
        self.__dir = dir_
        # Get long filename
        self.__longFilename = msg._getStringStream([dir_, '__substg1.0_3707'])

        # Get short filename
        self.__shortFilename = msg._getStringStream([dir_, '__substg1.0_3704'])

        # Get Content-ID
        self.__cid = msg._getStringStream([dir_, '__substg1.0_3712'])

        # Get attachment data
        if msg.Exists([dir_, '__substg1.0_37010102']):
            self.__type = 'data'
            self.__data = msg._getStream([dir_, '__substg1.0_37010102'])
        elif msg.Exists([dir_, '__substg1.0_3701000D']):
            if (self.props['37050003'].value & 0x7) != 0x5:
                if not debug:
                    raise NotImplementedError('Current version of ExtractMsg.py does not support extraction of containers that are not embeded msg files.')
                    #TODO add implementation
                else:
                    print('DEBUG: Debugging is true, ignoring NotImplementedError and printing debug info...')
                    print('DEBUG: _dir = {}'.format(_dir))
                    print('DEBUG: Writing properties stream to output:')
                    print('DEBUG: --------Start-Properties-Stream--------')
                    print(properHex(self.props.stream))
                    print('DEBUG: ---------End-Properties-Stream---------')
                    print('DEBUG: Writing directory contents to output:')
                    print('DEBUG: --------Start-Directory-Content--------')
                    for x in msg.listDir(True, True): print(x)
                    print('DEBUG: ---------End-Directory-Content---------')
            else:
                self.__prefix = msg.prefixList + [dir_, '__substg1.0_3701000D']
                self.__type = 'msg'
                self.__data = msg.__class__(self.msg.path, self.__prefix, self.__class__)
        else:
            raise Exception('Unknown file type')

    def saveEmbededMessage(self, contentId = False, json = False, useFileName = False, raw = False, customPath = None, customFilename = None):
        """
        Seperate function from save to allow it to
        easily be overridden by a subclass.
        """
        self.data.save(json, useFileName, raw, contentId, customPath, customFilename)

    def save(self, contentId = False, json = False, useFileName = False, raw = False, customPath = None, customFilename = None):
        # Check if the user has specified a custom filename
        if customFilename != None and customFilename != '':
            filename = customFilename
        else:
            # If not...
            # Check if user wants to save the file under the Content-id
            if contentId:
                filename = self.__cid
            # If filename is None at this point, use long filename as first preference
            filename = self.__longFilename
            # Otherwise use the short filename
            if filename is None:
                filename = self.__shortFilename
            # Otherwise just make something up!
            if filename is None:
                filename = 'UnknownFilename ' + \
                    ''.join(random.choice(string.ascii_uppercase + string.digits)
                            for _ in range(5)) + '.bin'

        if custom_path:
            filename = customPath + filename

        if self.__type == "data":
            f = open(filename, 'wb')
            f.write(self.__data)
            f.close()
        else:
            self.saveEmbededMessage(contentId, json, useFileName, raw, customPath, customFilename)
        return filename

    @property
    def cid(self):
        """
        Returns the content ID of the attachment, if it exists.
        """
        return self.__cid

    contend_id = cid

    @property
    def data(self):
        """
        Returns the attachment data.
        """
        return self.__data

    @property
    def dir(self):
        """
        Returns the directory inside the msg file where the attachment is located.
        """
        return self.__dir

    @property
    def longFilename(self):
        """
        Returns the long file name of the attachment, if it exists.
        """
        return self.__longFilename

    @property
    def msg(self):
        """
        Returns the Message instance the attachment belongs to.
        """
        return self.__msg

    @property
    def props(self):
        """
        Returns the Properties instance of the attachment.
        """
        try:
            return self.__props
        except:
            self.__props = Properties(self.msg._getStream(self.msg.prefixList + [self.__dir, '__properties_version1.0']), TYPE_ATTACHMENT)
            return self.__props

    @property
    def shortFilename(self):
        """
        Returns the short file name of the attachment, if it exists.
        """
        return self.__shortFilename

    @property
    def type(self):
        """
        Returns the type of the data.
        """
        return self.__type

class Properties:
    """
    Parser for msg properties files.
    """
    def __init__(self, stream, type = None, skip = None):
        self.__stream = stream
        self.__pos = 0
        self.__len = len(stream)
        self.__props = {}
        self.__naid = None
        self.__nrid = None
        self.__ac = None
        self.__rc = None
        if type != None:
            self.__intel = INTELLIGENCE_SMART
            if type == TYPE_MESSAGE:
                skip = 32
                self.__naid, self.__nrid, self.__ac, self.__rc = ST1.unpack(self.__stream[:24])
            elif type == TYPE_MESSAGE_EMBED:
                skip = 24
                self.__naid, self.__nrid, self.__ac, self.__rc = ST1.unpack(self.__stream[:24])
            else:
                skip = 8
        else:
            self.__intel = INTELLIGENCE_DUMB
            if skip == None:
                # This section of the skip handling is not very good.
                # While it does work, it is likely to create extra
                # properties that are created from the properties file's
                # header data. While that won't actually mess anything
                # up, it is far from ideal. Basically, this is the dumb
                # skip length calculation. Preferably, we want the type
                # to have been specified so all of the additional fields
                # will have been filled out
                skip = self.__len % 16
                if skip == 0:
                    skip = 32
        streams = divide(self.__stream[skip:], 16)
        for st in streams:
            a = Prop(st)
            self.__props[a.name] = a
        self.__pl = len(self.__props)

    def get(self, name):
        """
        Retrieve the property of :param name:.
        """
        try:
            return self.__props[name]
        except:
            if debug:
                print('DEBUG:')
                print(properHex(self.__stream))
                print(self.__props)
            raise

    def has_key(self, key):
        """
        Checks if :param key: is a key in the properties dictionary.
        """
        return key in self.__props

    def items(self):
        return self.__props.items()

    def iteritems(self):
        return self.__props.iteritems()

    def iterkeys(self):
        return self.__props.iterkeys()

    def itervalues(self):
        return self.__props.itervalues()

    def keys(self):
        return self.__props.keys()

    def values(self):
        return self.__props.values()

    def viewitems(self):
        return self.__props.viewitems()

    def viewkeys(self):
        return self.__props.viewkeys()

    def viewvalues(self):
        return self.__props.viewvalues()

    def __contains__(self, key):
        self.__props.__contains__(key)

    def __getitem__(self, key):
        return self.__props.__getitem__(key)

    def __iter__(self):
        return self.__props.__iter__()

    def __len__(self):
        """
        Returns the number of properties.
        """
        return self.__pl

    @property
    def __repr__(self):
        return self.__props.__repr__

    items.__doc__         = dict.items.__doc__
    iteritems.__doc__     = dict.iteritems.__doc__
    iterkeys.__doc__      = dict.iterkeys.__doc__
    itervalues.__doc__    = dict.itervalues.__doc__
    keys.__doc__          = dict.keys.__doc__
    values.__doc__        = dict.values.__doc__
    viewitems.__doc__     = dict.viewitems.__doc__
    viewkeys.__doc__      = dict.viewkeys.__doc__
    viewvalues.__doc__    = dict.viewvalues.__doc__

    @property
    def attachment_count(self):
        if self.__ac == None:
            raise TypeError('Properties instance must be intelligent and of type TYPE_MESSAGE to get attachment count.')
        return self.__ac

    @property
    def date(self):
        """
        Returns the send date contained in the Properties file.
        """
        try:
            return self.__date
        except:
            if self.has_key('00390040'):
                self.__date = fromTimeStamp(msgEpoch(self.get('00390040').value)).__format__('%a, %d %b %Y %H:%M:%S GMT %z')
            elif self.has_key('30080040'):
                self.__date = fromTimeStamp(msgEpoch(self.get('30080040').value)).__format__('%a, %d %b %Y %H:%M:%S GMT %z')
            elif self.has_key('30070040'):
                self.__date = fromTimeStamp(msgEpoch(self.get('30070040').value)).__format__('%a, %d %b %Y %H:%M:%S GMT %z')
            else:
                print('Warning: Error retrieving date. Setting as "Unknown". Please send the following data to developer:\n--------------------')
                print(properHex(self.__stream))
                print(self.keys())
                print('--------------------')
                self.__date = 'Unknown'
            return self.__date

    @property
    def intelligence(self):
        """
        Returns the inteligence level of the Properties instance.
        """
        return self.__intel

    @property
    def next_attachment_id(self):
        if self.__naid == None:
            raise TypeError('Properties instance must be intelligent and of type TYPE_MESSAGE to get next attachment id.')
        return self.__naid

    @property
    def next_recipient_id(self):
        if self.__nrid == None:
            raise TypeError('Properties instance must be intelligent and of type TYPE_MESSAGE to get next recipient id.')
        return self.__nrid

    @property
    def props(self):
        """
        Returns a copy of the internal properties dict.
        """
        return copy.deepcopy(self.__props)

    @property
    def recipient_count(self):
        if self.__rc == None:
            raise TypeError('Properties instance must be intelligent and of type TYPE_MESSAGE to get recipient count.')
        return self.__rc

    @property
    def stream(self):
        """
        Returns the data stream used to generate this Properties instance.
        """
        return self.__stream

class Prop:
    """
    Class to contain the data for a single property.

    Currently a work in progress.
    """
    def __init__(self, string):
        n = string[:4][::-1]
        self.__raw = string
        self.__name = properHex(n).upper()
        self.__type, self.__flags, self.__value = ST2.unpack(string)
        self.__value = self.parse_type(self.__type, self.__value)
        self.__fm = self.__flags & 1 == 1
        self.__fr = self.__flags & 2 == 2
        self.__fw = self.__flags & 4 == 4

    def parse_type(self, type, stream):
        """
        Converts the data in :param stream: to a
        much more accurate type, specified by
        :param type:, if possible.

        WARNING: Not done.
        """
        #WARNING Not done.
        value = stream
        if type == 0x0000: #PtypUnspecified
            pass;
        elif type == 0x0001: #PtypNull
            if value != b'\x00\x00\x00\x00\x00\x00\x00\x00':
                print('Warning: Property type is PtypNull, but is not equal to 0.')
            value = None
        elif type == 0x0002: #PtypInteger16
            value = STI16.unpack(value)[0]
        elif type == 0x0003: #PtypInteger32
            value = STI32.unpack(value)[0]
        elif type == 0x0004: #PtypFloating32
            value = STF32.unpack(value)[0]
        elif type == 0x0005: #PtypFloating64
            value = STF64.unpack(value)[0]
        elif type == 0x0006: #PtypCurrency
            value = (STI64.unpack(value))[0]/10000.0
        elif type == 0x0007: #PtypFloatingTime
            value = STF64.unpack(value)[0]
            #TODO parsing for this
            pass;
        elif type == 0x000A: #PtypErrorCode
            value = STI32.unpack(value)[0]
            #TODO parsing for this
            pass;
        elif type == 0x000B: #PtypBoolean
            value = bool(ST3.unpack(value)[0])
        elif type == 0x000D: #PtypObject/PtypEmbeddedTable
            #TODO parsing for this
            pass;
        elif type == 0x0014: #PtypInteger64
            value = STI64.unpack(value)[0]
        elif type == 0x001E: #PtypString8
            #TODO parsing for this
            pass;
        elif type == 0x001F: #PtypString
            #TODO parsing for this
            pass;
        elif type == 0x0040: #PtypTime
            value = ST3.unpack(value)[0]
        elif type == 0x0048: #PtypGuid
            #TODO parsing for this
            pass;
        elif type == 0x00FB: #PtypServerId
            #TODO parsing for this
            pass;
        elif type == 0x00FD: #PtypRestriction
            #TODO parsing for this
            pass;
        elif type == 0x00FE: #PtypRuleAction
            #TODO parsing for this
            pass;
        elif type == 0x0102: #PtypBinary
            #TODO parsing for this
            # Smh, how on earth am I going to code this???
            pass;
        elif type & 0x1000 == 0x1000: #PtypMultiple
            #TODO parsing for `multiple` types
            pass;
        return value;

    @property
    def flag_mandatory(self):
        """
        Boolean, is the "mandatory" flag set?
        """
        return self.__fm

    @property
    def flag_readable(self):
        """
        Boolean, is the "readable" flag set?
        """
        return self.__fr

    @property
    def flag_writable(self):
        """
        Boolean, is the "writable" flag set?
        """
        return self.__fw

    @property
    def flags(self):
        """
        Integer that contains property flags.
        """
        return self.__flags

    @property
    def name(self):
        """
        Property "name".
        """
        return self.__name

    @property
    def raw(self):
        """
        Raw binary string that defined the property.
        """
        return self.__raw

    @property
    def value(self):
        """
        Property value.
        """
        return self.__value

class Recipient:
    """
    Contains the data of one or the recipients in an msg file.
    """
    def __init__(self, num, msg):
        self.__msg = msg #Allows calls to original msg file
        self.__dir = '__recip_version1.0_#' + num.rjust(8,'0')
        self.__props = Properties(msg._getStream(self.__dir + '/__properties_version1.0'), TYPE_RECIPIENT)
        self.__email = msg._getStringStream(self.__dir + '/__substg1.0_39FE')
        self.__name = msg._getStringStream(self.__dir + '/__substg1.0_3001')
        self.__type = self.__props.get('0C150003').value
        self.__formatted = '{0} <{1}>'.format(self.__name, self.__email)

    @property
    def email(self):
        """
        Returns the recipient's email.
        """
        return self.__email

    @property
    def formatted(self):
        """
        Returns the formatted recipient string.
        """
        return self.__formatted

    @property
    def name(self):
        """
        Returns the recipient's name.
        """
        return self.__name

    @property
    def props(self):
        """
        Returns the Properties instance of the recipient.
        """
        return self.__props

    @property
    def type(self):
        """
        Returns the recipient type.
        Sender if `type & 0xf == 0`
        To if `type & 0xf == 1`
        Cc if `type & 0xf == 2`
        Bcc if `type & 0xf == 3`
        """
        return self.__type



class Message(OleFile.OleFileIO):
    def __init__(self, path, prefix = '', attachmentClass = Attachment, filename = None):
        """
        :param path: path to the msg file in the system or is the raw msg file.
        :param prefix: used for extracting embeded msg files
            inside the main one. Do not set manually unless
            you know what you are doing.
        :param attachmentClass: optional, the class the Message object
            will use for attachments. You probably should
            not change this value unless you know what you
            are doing.
        :param filename: optional, the filename to be used by default when saving.
        """
        #WARNING DO NOT MANUALLY MODIFY PREFIX. Let the program set it.
        if debug:
            print('DEBUG: prefix: {}'.format(prefix))
        self.__path = path
        self.__attachmentClass = attachmentClass
        OleFile.OleFileIO.__init__(self, path)
        prefixl = []
        if prefix != '':
            if not isinstance(prefix, stri):
                try:
                    prefix = '/'.join(prefix)
                except:
                    raise TypeException('Invalid prefix type: ' + type(prefix) + '\n(This was probably caused by you setting it manually).')
            prefix = prefix.replace('\\', '/')
            g = prefix.split("/")
            if g[-1] == '':
                g.pop()
            prefixl = g
            if prefix[-1] != '/':
                prefix += '/'
            filename = self._getStringStream(prefixl[:-1] + ['__substg1.0_3001'], prefix = False)
        self.__prefix = prefix
        self.__prefixList = prefixl
        if filename != None:
            self.filename = filename
        elif has_len(path):
            if len(path) < 1536:
                self.filename = path
            else:
                self.filename = None
        else:
            self.filename = None

        # Initialize properties in the order that is least likely to cause bugs.
        # TODO have each function check for initialization of needed data so these
        # lines will be unnecessary.
        self.mainProperties
        self.recipients
        self.attachments
        self.to
        self.cc
        self.sender
        self.header
        self.date
        self.__crlf = '\n' #This variable keeps track of what the new line character should be
        self.body

    def listDir(self, streams = True, storages = False):
        """
        Replacement for OleFileIO.listdir that runs at the current prefix directory.
        """
        temp = self.listdir(streams, storages)
        if self.__prefix == '':
            return temp
        prefix = self.__prefix.split('/')
        if prefix[-1] == '':
            prefix.pop()
        out = []
        for x in temp:
            good = True
            if len(x) <= len(prefix):
                good = False
            if good:
                for y in range(len(prefix)):
                    if x[y] != prefix[y]:
                        good = False
            if good:
                out.append(x)
        return out

    def Exists(self, inp):
        """
        Checks if :param inp: exists in the msg file.
        """
        if isinstance(inp, list):
            inp = self.__prefixList + inp
        else:
            inp = self.__prefix + inp
        return self.exists(inp)

    def _getStream(self, filename, prefix = True):
        if isinstance(filename, list):
            filename = '/'.join(filename)
        if prefix:
            filename = self.__prefix + filename
        if self.exists(filename):
            stream = self.openstream(filename)
            return stream.read()
        else:
            return None

    def _getStringStream(self, filename, prefer = 'unicode', prefix = True):
        """
        Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then :param prefer: specifies which will be
        returned.
        """

        if isinstance(filename, list):
            # Join with slashes to make it easier to append the type
            filename = '/'.join(filename)

        asciiVersion = self._getStream(filename + '001E', prefix)
        unicodeVersion = windowsUnicode(self._getStream(filename + '001F', prefix))
        if debug:
            print('DEBUG: _getStringSteam called for {}. Ascii version found: {}. Unicode version found: {}.'.format(filename, asciiVersion != None, unicodeVersion != None))
        if asciiVersion is None:
            return unicodeVersion
        elif unicodeVersion is None:
            return asciiVersion
        else:
            if prefer == 'unicode':
                return unicodeVersion
            else:
                return asciiVersion

    @property
    def path(self):
        """
        Returns the message path if generated from a file,
        otherwise returns the data used to generate the
        Message instance.
        """
        return self.__path

    @property
    def prefix(self):
        """
        Returns the prefix of the Message instance.
        Intended for developer use.
        """
        return self.__prefix

    @property
    def prefixList(self):
        """
        Returns the prefix list of the Message instance.
        Intended for developer use.
        """
        return self.__prefixList

    @property
    def subject(self):
        """
        Returns the message subject, if it exists.
        """
        try:
            return self._subject
        except:
            self._subject = encode(self._getStringStream('__substg1.0_0037'))
            return self._subject

    @property
    def header(self):
        """
        Returns the message header, if it exists. Otherwise it will generate one.
        """
        try:
            return self._header
        except Exception:
            headerText = self._getStringStream('__substg1.0_007D')
            if headerText is not None:
                self._header = EmailParser().parsestr(headerText)
                self._header['date'] = self.date
            else:
                header = {
                    'date': self.date,
                    'from': self.sender,
                    'to': self.to,
                    'cc': self.cc
                }
                self._header = header
            return self._header

    def headerInit(self):
        """
        Checks whether the header has been initialized.
        """
        try:
            self._header
            return True
        except:
            return False

    @property
    def mainProperties(self):
        """
        Returns the Properties instance used by the Message instance.
        """
        try:
            return self._prop
        except:
            self._prop = Properties(self._getStream('__properties_version1.0'), TYPE_MESSAGE if self.__prefix == '' else TYPE_MESSAGE_EMBED)
            return self._prop

    @property
    def date(self):
        """
        Returns the send date, if it exists.
        """
        try:
            return self._date
        except:
            self._date = self._prop.date
            return self._date


    @property
    def parsedDate(self):
        return email.utils.parsedate(self.date)

    @property
    def sender(self):
        """
        Returns the message sender, if it exists.
        """
        try:
            return self._sender
        except Exception:
            # Check header first
            if self.headerInit():
                headerResult = self.header['from']
                if headerResult is not None:
                    self._sender = headerResult
                    return headerResult
            # Extract from other fields
            text = self._getStringStream('__substg1.0_0C1A')
            email = self._getStringStream('__substg1.0_5D01') #Will not give an email address sometimes. Seems to exclude the email address if YOU are the sender.
            result = None
            if text is None:
                result = email
            else:
                result = text
                if email is not None:
                    result += ' <' + email + '>'

            self._sender = result
            return result

    @property
    def to(self):
        """
        Returns the to field, if it exists.
        """
        try:
            return self._to
        except Exception:
            # Check header first
            if self.headerInit():
                headerResult = self.header['to']
                if headerResult is not None:
                    self._to = headerResult
            else:
                f = []
                for x in self.recipients:
                    if x.type & 0x0000000f == 1:
                        f.append(x.formatted)
                if len(f) > 0:
                    st = f[0]
                    if len(f) > 1:
                        for x in range(1, len(f)):
                            st += '; {0}'.format(f[x])
                    self._to = st
                else:
                    self._to = None
            return self._to

    @property
    def compressedRtf(self):
        """
        Returns the compressed RTF stream, if it exists.
        """
        try:
            return self._compressedRtf
        except Exception:
            self._compressedRtf = self._getStream('__substg1.0_10090102')
            return self._compressedRtf

    @property
    def htmlBody(self):
        """
        Returns the html body, if it exists.
        """
        try:
            return self._htmlBody
        except Exception:
            self._htmlBody = self._getStream('__substg1.0_10130102')
            return self._htmlBody

    @property
    def cc(self):
        """
        Returns the cc field, if it exists.
        """
        try:
            return self._cc
        except Exception:
            # Check header first
            if self.headerInit():
                headerResult = self.header['cc']
                if headerResult is not None:
                    self._cc = headerResult
            else:
                f = []
                for x in self.recipients:
                    if x.type & 0x0000000f == 2:
                        f.append(x.formatted)
                if len(f) > 0:
                    st = f[0]
                    if len(f) > 1:
                        for x in range(1, len(f)):
                            st += '; {0}'.format(f[x])
                    self._cc = st
                else:
                    self._cc = None
            return self._cc

    @property
    def body(self):
        """
        Returns the message body, if it exists.
        """
        try:
            return self._body
        except Exception:
            self._body = encode(self._getStringStream('__substg1.0_1000'))
            a = re.search('\n', self._body)
            if a != None:
                if re.search('\r\n', self._body) != None:
                    self.__crlf = '\r\n'
            return self._body

    @property
    def crlf(self):
        """
        Returns the value of self.__crlf, should you need it for whatever reason.
        """
        self.body
        return self.__crlf

    @property
    def attachmentClass(self):
        """
        Returns the Attachment class being used, should you need to use it externally for whatever reason.
        """
        return self.__attachmentClass

    @property
    def attachments(self):
        """
        Returns a list of all attachments.
        """
        try:
            return self._attachments
        except Exception:
            # Get the attachments
            attachmentDirs = []

            for dir_ in self.listDir():
                if dir_[len(self.__prefixList)].startswith('__attach') and dir_[len(self.__prefixList)] not in attachmentDirs:
                    attachmentDirs.append(dir_[len(self.__prefixList)])

            self._attachments = []

            for attachmentDir in attachmentDirs:
                self._attachments.append(self.__attachmentClass(self, attachmentDir))

            return self._attachments

    @property
    def recipients(self):
        """
        Returns a list of all recipients.
        """
        try:
            return self._recipients
        except Exception:
            # Get the recipients
            recipientDirs = []

            for dir_ in self.listDir():
                if dir_[len(self.__prefixList)].startswith('__recip') and dir_[len(self.__prefixList)] not in recipientDirs:
                    recipientDirs.append(dir_[len(self.__prefixList)])

            self._recipients = []

            for recipientDir in recipientDirs:
                self._recipients.append(Recipient(recipientDir.split('#')[-1], self))

            return self._recipients

    def save(self, toJson=False, useFileName=False, raw=False, ContentId=False, customPath = None, customFilename = None):
        """
        Saves the message body and attachments found in the message. Setting toJson
        to true will output the message body as JSON-formatted text. The body and
        attachments are stored in a folder. Setting useFileName to true will mean that
        the filename is used as the name of the folder; otherwise, the message's date
        and subject are used as the folder name.

        Currently, :param customPath: and :param customFilename: don't do anything.
        """

        if useFileName:
            # strip out the extension
            if self.filename != None:
                dirName = self.filename.split('/').pop().split('.')[0]
            else:
                ValueError('Filename must be specified, or path must have been an actual path, to save using filename')
        else:
            # Create a directory based on the date and subject of the message
            d = self.parsedDate
            if d is not None:
                dirName = '{0:02d}-{1:02d}-{2:02d}_{3:02d}{4:02d}'.format(*d)
            else:
                dirName = 'UnknownDate'

            if self.subject is None:
                subject = '[No subject]'
            else:
                subject = ''.join(i for i in self.subject if i not in r'\/:*?"<>|')

            dirName = dirName + ' ' + subject

        try:
            os.makedirs(dirName)
        except Exception:
            newDirName = addNumToDir(dirName)
            if newDirName is not None:
                dirName = newDirName
            else:
                raise Exception(
                    "Failed to create directory '%s'. Does it already exist?" %
                    dirName
                    )

        oldDir = os.getcwd()
        try:
            os.chdir(dirName)

            # Save the message body
            fext = 'json' if toJson else 'text'
            f = open('message.' + fext, 'w')
            # From, to , cc, subject, date



            attachmentNames = []
            # Save the attachments
            for attachment in self.attachments:
                attachmentNames.append(attachment.save(ContentId))

            if toJson:

                emailObj = {'from': xstr(self.sender),
                            'to': xstr(self.to),
                            'cc': xstr(self.cc),
                            'subject': xstr(self.subject),
                            'date': xstr(self.date),
                            'attachments': attachmentNames,
                            'body': decode_utf7(self.body)}

                f.write(json.dumps(emailObj, ensure_ascii=True))
            else:
                f.write('From: ' + xstr(self.sender) + self.__crlf)
                f.write('To: ' + xstr(self.to) + self.__crlf)
                f.write('CC: ' + xstr(self.cc) + self.__crlf)
                f.write('Subject: ' + xstr(self.subject) + self.__crlf)
                f.write('Date: ' + xstr(self.date) + self.__crlf)
                f.write('-----------------' + self.__crlf + self.__crlf)
                f.write(self.body)

            f.close()

        except Exception as e:
            self.saveRaw()
            raise

        finally:
            # Return to previous directory
            os.chdir(oldDir)

    def saveRaw(self):
        # Create a 'raw' folder
        oldDir = os.getcwd()
        try:
            rawDir = 'raw'
            os.makedirs(rawDir)
            os.chdir(rawDir)
            sysRawDir = os.getcwd()

            # Loop through all the directories
            for dir_ in self.listdir():
                sysdir = '/'.join(dir_)
                code = dir_[-1][-8:-4]
                global properties
                if code in properties:
                    sysdir = sysdir + ' - ' + properties[code]
                os.makedirs(sysdir)
                os.chdir(sysdir)

                # Generate appropriate filename
                if dir_[-1].endswith('001E'):
                    filename = 'contents.txt'
                else:
                    filename = 'contents'

                # Save contents of directory
                f = open(filename, 'wb')
                f.write(self._getStream(dir_))
                f.close()

                # Return to base directory
                os.chdir(sysRawDir)

        finally:
            os.chdir(oldDir)

    def dump(self):
        """
        Prints out a summary of the message
        """
        print('Message')
        print('Subject:', self.subject)
        print('Date:', self.date)
        print('Body:')
        print(self.body)

    def debug(self):
        for dir_ in self.listDir():
            if dir_[-1].endswith('001E') or dir_[-1].endswith('001F'):
                print('Directory: ' + str(dir_[:-1]))
                print('Contents: {}'.format(self._getStream(dir_)))

    def save_attachments(self, contentId = False, json = False, useFileName = False, raw = False, customPath = None):
        """
        Saves only attachments in the same folder.
        """
        for attachment in self.attachments:
            attachment.save(contentId, json, useFileName, raw, customPath)


if __name__ == '__main__':
    if len(sys.argv) <= 1:
        print(__doc__)
        print("""
Launched from command line, this script parses Microsoft Outlook Message files
and save their contents to the current directory. On error the script will
write out a 'raw' directory will all the details from the file, but in a
less-than-desirable format. To force this mode, the flag '--raw'
can be specified.

Usage:  <file> [file2 ...]
   or:  --raw <file>
   or:  --json

Additionally, use the flag '--use-content-id' to save files by their content ID (should they have one)

To name the directory as the .msg file, use the flag '--use-file-name'

To turn on the printing of debugging information, use the flag '--debug'
""")
        sys.exit()

    writeRaw = False
    toJson = False
    useFileName = False
    useContentId = False

    for rawFilename in sys.argv[1:]:
        if rawFilename == '--raw':
            writeRaw = True

        if rawFilename == '--json':
            toJson = True

        if rawFilename == '--use-file-name':
            useFileName = True

        if rawFilename == '--use-content-id':
            useContentId = True

        if rawFilename == '--debug':
            debug = True

        for filename in glob.glob(rawFilename):
            msg = Message(filename)
            try:
                if writeRaw:
                    msg.saveRaw()
                else:
                    msg.save(toJson, useFileName, False, useContentId)
            except Exception as e:
                # msg.debug()
                print("Error with file '" + filename + "': " +
                      traceback.format_exc())
