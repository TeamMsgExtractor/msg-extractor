import logging

from extract_msg import constants
from extract_msg.properties import Properties


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class Recipient(object):
    """
    Contains the data of one of the recipients in an msg file.
    """

    def __init__(self, _dir, msg):
        object.__init__(self)
        self.__msg = msg  # Allows calls to original msg file
        self.__dir = _dir
        self.__props = Properties(msg._getStream(self.__dir + '/__properties_version1.0'), constants.TYPE_RECIPIENT)
        self.__email = msg._getStringStream(self.__dir + '/__substg1.0_39FE')
        if not self.__email:
            self.__email = msg._getStringStream(self.__dir + '/__substg1.0_3003')
        self.__name = msg._getStringStream(self.__dir + '/__substg1.0_3001')
        self.__type = self.__props.get('0C150003').value
        self.__formatted = u'{0} <{1}>'.format(self.__name, self.__email)

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
