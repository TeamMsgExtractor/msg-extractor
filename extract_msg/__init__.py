"""
extract_msg:
    Extracts emails and attachments saved in Microsoft Outlook's .msg files

https://github.com/mattgwwalker/msg-extractor
"""

__author__ = 'Matthew Walker & The Elemental of Creation'
__date__ = '2018-05-22'
__version__ = '0.20.1'

from extract_msg.extract_msg import Attachment, Properties, Props, Recipient, Message, msg_epoch, fromdatetime, constants, parse_type, properHex
