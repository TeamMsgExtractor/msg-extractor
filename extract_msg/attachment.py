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
        self.__msg = msg
        self.__dir = dir_
        self.__props = Properties(
            self.msg._getStream([self.__dir, '__properties_version1.0']),
            constants.TYPE_ATTACHMENT)
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
                    raise NotImplementedError(
                        'Current version of extract_msg does not support extraction of containers that are not embeded msg files.')
                    # TODO add implementation
                else:
                    # DEBUG
                    logger.debug('Debugging is true, ignoring NotImplementedError and printing debug info...')
                    logger.debug('dir_ = {}'.format(dir_))
                    logger.debug('Writing properties stream to output:')
                    logger.debug('--------Start-Properties-Stream--------\n' +
                                 properHex(self.props.stream) +
                                 '\n---------End-Properties-Stream---------')
                    logger.debug('Writing directory contents to output:')
                    logger.debug('--------Start-Directory-Content--------')
                    logger.debug('\n'.join([repr(x) for x in msg.listDir(True, True)]))
                    logger.debug('---------End-Directory-Content---------')
            else:
                self.__prefix = msg.prefixList + [dir_, '__substg1.0_3701000D']
                self.__type = 'msg'
                self.__data = msg.__class__(self.msg.path, self.__prefix, self.__class__)
        else:
            raise TypeError('Unknown attachment type.')

    def save(self, contentId=False, json=False, useFileName=False, raw=False, customPath=None, customFilename=None):
        # Check if the user has specified a custom filename
        filename = None
        if customFilename is not None and customFilename != '':
            filename = customFilename
        else:
            # If not...
            # Check if user wants to save the file under the Content-id
            if contentId:
                filename = self.__cid
            # If filename is None at this point, use long filename as first preference
            if filename is None:
                filename = self.__longFilename
            # Otherwise use the short filename
            if filename is None:
                filename = self.__shortFilename
            # Otherwise just make something up!
            if filename is None:
                filename = 'UnknownFilename ' + \
                           ''.join(random.choice(string.ascii_uppercase + string.digits)
                                   for _ in range(5)) + '.bin'

        if customPath is not None and customPath != '':
            if customPath[-1] != '/' or customPath[-1] != '\\':
                customPath += '/'
            filename = customPath + filename

        if self.__type == "data":
            with open(filename, 'wb') as f:
                f.write(self.__data)
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
