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
from extract_msg.utils import has_len, inputToString, windowsUnicode
from extract_msg.exceptions import InvalidFileFormatError, MissingEncodingError



logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

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
        :param attachmentClass: optional, the class the MSGFile object
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

    @property
    def areStringsUnicode(self):
        """
        Returns a boolean telling if the strings are unicode encoded.
        """
        try:
            return self.__bStringsUnicode
        except AttributeError:
            if self.mainProperties.has_key('340D0003'):
                if (self.mainProperties['340D0003'].value & 0x40000) != 0:
                    self.__bStringsUnicode = True
                    return self.__bStringsUnicode
            self.__bStringsUnicode = False
            return self.__bStringsUnicode

    @property
    def attachmentClass(self):
        """
        Returns the Attachment class being used, should you need to use it externally for whatever reason.
        """
        return self.__attachmentClass

    @property
    def classType(self):
        """
        The class type of the MSG file.
        """
        return self._ensureSet('_classType', '__substg1.0_001A')

    @property
    def mainProperties(self):
        """
        Returns the Properties instance used by the MSGFile instance.
        """
        try:
            return self._prop
        except AttributeError:
            self._prop = Properties(self._getStream('__properties_version1.0'),
                                    constants.TYPE_MESSAGE if self.prefix == '' else constants.TYPE_MESSAGE_EMBED)
            return self._prop

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
        return copy.deepcopy(self.__prefixList)


    @property
    def stringEncoding(self):
        try:
            return self.__stringEncoding
        except AttributeError:
            # We need to calculate the encoding
            # Let's first check if the encoding will be unicode:
            if self.areStringsUnicode:
                self.__stringEncoding = "utf-16-le"
                return self.__stringEncoding
            else:
                # Well, it's not unicode. Now we have to figure out what it IS.
                if not self.mainProperties.has_key('3FFD0003'):
                    raise MissingEncodingError('Encoding property not found')
                enc = self.mainProperties['3FFD0003'].value
                # Now we just need to translate that value
                # Now, this next line SHOULD work, but it is possible that it might not...
                self.__stringEncoding = str(enc)
                return self.__stringEncoding
