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
    """
    Parser for .msg files
    """
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
            
    def _ensureSet(self, variable, streamID, stringStream = True):
        """
        Ensures that the variable exists, otherwise will set it using the specified stream.
        After that, return said variable.
        If the specified stream is not a string stream, make sure to set :param string stream: to False.
        """
        try:
            return getattr(self, variable)
        except AttributeError:
            if stringStream:
                value = self._getStringStream(streamID)
            else:
                value = self._getStream(streamID)
            setattr(self, variable, value)
            return value

    def _getStream(self, filename, prefix = True):
        """
        Gets a binary representation of the requested filename.
        This should ALWAYS return a bytes object (string in python 2)
        """
        filename = self.fix_path(filename, prefix)
        if self.exists(filename):
            with self.openstream(filename) as stream:
                return stream.read()
        else:
            logger.info('Stream "{}" was requested but could not be found. Returning `None`.'.format(filename))
            return None

    def _getStringStream(self, filename, prefix = True):
        """
        Gets a string representation of the requested filename.
        This should ALWAYS return a string (Unicode in python 2)
        """

        filename = self.fix_path(filename, prefix)
        if self.areStringsUnicode:
            return windowsUnicode(self._getStream(filename + '001F', prefix = False))
        else:
            tmp = self._getStream(filename + '001E', prefix = False)
            return None if tmp is None else tmp.decode(self.stringEncoding)
        
    def debug(self):
        for dir_ in self.listDir():
            if dir_[-1].endswith('001E') or dir_[-1].endswith('001F'):
                print('Directory: ' + str(dir_[:-1]))
                print('Contents: {}'.format(self._getStream(dir_)))
                
    def Exists(self, inp):
        """
        Checks if :param inp: exists in the msg file. Does not always go to the top, starts at specified point
        """
        inp = self.fix_path(inp)
        return self.exists(inp)
    
    def sExists(self, inp):
        """
        Checks if string stream :param inp: exists in the msg file.
        """
        inp = self.fix_path(inp)
        return self.exists(inp + '001F') or self.exists(inp + '001E')
    
    def fix_path(self, inp, prefix = True):
        """
        Changes paths so that they have the proper
        prefix (should :param prefix: be True) and
        are strings rather than lists or tuples.
        """
        if isinstance(inp, (list, tuple)):
            inp = '/'.join(inp)
        if prefix:
            inp = self.__prefix + inp
        return inp
