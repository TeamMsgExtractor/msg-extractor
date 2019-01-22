import logging

from extract_msg import constants
from extract_msg.properties import Properties
from extract_msg.utils import properHex

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class Attachment(object):
    """
    Developer version of the `extract_msg.attachment.Attachment` class.
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
            msg._getStream([self.__dir, '__properties_version1.0']),
            constants.TYPE_ATTACHMENT)

        # Get attachment data
        if msg.Exists([dir_, '__substg1.0_37010102']):
            self.__type = 'data'
            self.__data = msg._getStream([dir_, '__substg1.0_37010102'])
        elif msg.Exists([dir_, '__substg1.0_3701000D']):
            if (self.__props['37050003'].value & 0x7) != 0x5:
                logger.log(5, 'Printing details of NotImplementedError...')
                logger.log(5, 'dir_ = {}'.format(dir_))
                logger.log(5, 'Writing properties stream to output:')
                logger.log(5, '--------Start-Properties-Stream--------\n' +
                              properHex(self.__props.stream) +
                              '\n---------End-Properties-Stream---------')
                logger.log(5, 'Writing directory contents to output:')
                logger.log(5, '--------Start-Directory-Content--------\n' +
                              '\n'.join([repr(x) for x in msg.listDir(True, True)]))
                logger.log(5, '---------End-Directory-Content---------')
                logger.log(5, 'End of NotImplementedError details')
            else:
                self.__prefix = msg.prefixList + [dir_, '__substg1.0_3701000D']
                self.__type = 'msg'
                self.__data = msg.__class__(msg.path, self.__prefix)
        else:
            raise TypeError('Unknown attachment type.')

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
    def type(self):
        """
        Returns the type of the data.
        """
        return self.__type
