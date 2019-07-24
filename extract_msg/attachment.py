import logging
import random
import string

from extract_msg import constants
from extract_msg.properties import Properties
from extract_msg.utils import properHex

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class Attachment(object):
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
        object.__init__(self)
        self.msg = msg
        self.dir = dir_
        self.props = Properties(self._getStream('__properties_version1.0'),
            constants.TYPE_ATTACHMENT)
        # Get long filename
        self.longFilename = self._getStringStream('__substg1.0_3707')

        # Get short filename
        self.shortFilename = self._getStringStream('__substg1.0_3704')

        # Get Content-ID
        self.cid = self._getStringStream('__substg1.0_3712')
        self.contend_id = self.cid

        # Get attachment data
        if self.Exists('__substg1.0_37010102'):
            self.type = 'data'
            self.data = self._getStream('__substg1.0_37010102')
        elif self.Exists('__substg1.0_3701000D'):
            if (self.props['37050003'].value & 0x7) != 0x5:
                raise NotImplementedError(
                    'Current version of extract_msg does not support extraction of containers that are not embedded msg files.')
                # TODO add implementation
            else:
                self.prefix = msg.prefixList + [dir_, '__substg1.0_3701000D']
                self.type = 'msg'
                self.data = msg.__class__(self.msg.path, self.prefix, self.__class__)
        else:
            # TODO Handling for special attacment types (like 0x00000007)
            raise TypeError('Unknown attachment type.')

    def _getStream(self, filename):
        return self.msg._getStream([self.dir, filename])

    def _getStringStream(self, filename):
        """
        Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then :param prefer: specifies which will be
        returned.
        """
        return self.msg._getStringStream([self.dir, filename])

    def Exists(self, filename):
        """
        Checks if stream exists inside the attachment folder.
        """
        return self.msg.Exists([self.dir, filename])

    def sExists(self, filename):
        """
        Checks if the string stream exists inside the attachment folder.
        """
        return self.msg.sExists([self.dir, filename])

    def save(self, contentId=False, json=False, useFileName=False, raw=False, customPath=None, customFilename=None):
        # Check if the user has specified a custom filename
        filename = None
        if customFilename is not None and customFilename != '':
            filename = customFilename
        else:
            # If not...
            # Check if user wants to save the file under the Content-id
            if contentId:
                filename = self.cid
            # If filename is None at this point, use long filename as first preference
            if filename is None:
                filename = self.longFilename
            # Otherwise use the short filename
            if filename is None:
                filename = self.shortFilename
            # Otherwise just make something up!
            if filename is None:
                filename = 'UnknownFilename ' + \
                           ''.join(random.choice(string.ascii_uppercase + string.digits)
                                   for _ in range(5)) + '.bin'

        if customPath is not None and customPath != '':
            if customPath[-1] != '/' or customPath[-1] != '\\':
                customPath += '/'
            filename = customPath + filename

        if self.type == "data":
            with open(filename, 'wb') as f:
                f.write(self.data)
        else:
            self.saveEmbededMessage(contentId, json, useFileName, raw, customPath, customFilename)
        return filename

    def saveEmbededMessage(self, contentId=False, json=False, useFileName=False, raw=False, customPath=None,
                           customFilename=None):
        """
        Seperate function from save to allow it to
        easily be overridden by a subclass.
        """
        self.data.save(json, useFileName, raw, contentId, customPath, customFilename)
