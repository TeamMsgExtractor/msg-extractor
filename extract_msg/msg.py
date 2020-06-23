import copy
import email.utils
import json
import logging
import re

import compressed_rtf
from imapclient.imapclient import decode_utf7
import olefile

from email.parser import Parser as EmailParser
from extract_msg import constants
from extract_msg.attachment import Attachment
from extract_msg.compat import os_ as os
from extract_msg.properties import Properties
from extract_msg.recipient import Recipient
from extract_msg.utils import addNumToDir, has_len, inputToBytes, inputToString, windowsUnicode
from extract_msg.exceptions import InvalidFileFormat



class MSGFile(olefile.OleFileIO):
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
        # WARNING DO NOT MANUALLY MODIFY PREFIX. Let the program set it.
        self.__path = path
        self.__attachmentClass = attachmentClass

        try:
            olefile.OleFileIO.__init__(self, path)
        except IOError as e:    # py2 and py3 compatible
            logger.error(e)
            if str(e) == 'not an OLE2 structured storage file':
                raise InvalidFileFormat(e)
            else:
                raise

        prefixl = []
        tmp_condition = prefix != ''
        if tmp_condition:
            try:
                prefix = inputToString(prefix, 'utf-8')
            except:
                try:
                    prefix = '/'.join(prefix)
                except:
                    raise TypeError('Invalid prefix type: ' + str(type(prefix)) +
                                    '\n(This was probably caused by you setting it manually).')
            prefix = prefix.replace('\\', '/')
            g = prefix.split("/")
            if g[-1] == '':
                g.pop()
            prefixl = g
            if prefix[-1] != '/':
                prefix += '/'
        self.__prefix = prefix
        self.__prefixList = prefixl
        if tmp_condition:
            filename = self._getStringStream(prefixl[:-1] + ['__substg1.0_3001'], prefix=False)
        if filename is not None:
            self.filename = filename
        elif has_len(path):
            if len(path) < 1536:
                self.filename = path
            else:
                self.filename = None
        else:
            self.filename = None
