import logging
import random
import string

from extract_msg import constants
from extract_msg.attachment_base import AttachmentBase
from extract_msg.named import NamedAttachmentProperties
from extract_msg.prop import FixedLengthProp, VariableLengthProp
from extract_msg.properties import Properties
from extract_msg.utils import openMsg, inputToString, prepareFilename, verifyPropertyId, verifyType

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class Attachment(AttachmentBase):
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
        AttachmentBase.__init__(self, msg, dir_)

        # Get attachment data
        if self.exists('__substg1.0_37010102'):
            self.__type = 'data'
            self.__data = self._getStream('__substg1.0_37010102')
        elif self.exists('__substg1.0_3701000D'):
            if (self.props['37050003'].value & 0x7) != 0x5:
                raise NotImplementedError(
                    'Current version of extract_msg does not support extraction of containers that are not embedded msg files.')
                # TODO add implementation
            else:
                self.__prefix = msg.prefixList + [dir_, '__substg1.0_3701000D']
                self.__type = 'msg'
                self.__data = openMsg(self.msg.path, self.__prefix, self.__class__, overrideEncoding = msg.overrideEncoding, attachmentErrorBehavior = msg.attachmentErrorBehavior)
        elif (self.props['37050003'].value & 0x7) == 0x7:
            # TODO Handling for special attacment type 0x7
            self.__type = 'web'
            raise NotImplementedError('Attachments of type afByWebReference are not currently supported.')
        else:
            raise TypeError('Unknown attachment type.')

    def getFilename(self, **kwargs):
        """
        Returns the filename to use for the attachment.

        :param contentId:      Use the contentId, if available.
        :param customFilename: A custom name to use for the file.

        If the filename starts with "UnknownFilename" then there is no guarentee
        that the files will have exactly the same filename.
        """
        filename = None
        customFilename = kwargs.get('customFilename')
        if customFilename:
            # First we need to validate it. If there are invalid characters, this will detect it.
            if constants.RE_INVALID_PATH_CHARACTERS.search(customFilename):
                raise ValueError('Invalid character found in customFilename. Must not contain any of the following characters: \\/:*?"<>|')
            filename = customFilename
        else:
            # If not...
            # Check if user wants to save the file under the Content-id
            if kwargs.get('contentId', False):
                filename = self.cid
            # If filename is None at this point, use long filename as first preference
            if filename is None:
                filename = self.longFilename
            # Otherwise use the short filename
            if filename is None:
                filename = self.shortFilename
            # Otherwise just make something up!
            if filename is None:
                return self.randomFilename

        return filename

    def regenerateRandomName(self):
        """
        Used to regenerate the random filename used if the attachment cannot
        find a usable filename.
        """
        self.__randomName = filename = 'UnknownFilename ' + \
                   ''.join(random.choice(string.ascii_uppercase + string.digits)
                           for _ in range(5)) + '.bin'

    def save(self, contentId = False, json = False, useFileName = False, raw = False, customPath = None,
             customFilename = None):#, html = False, rtf = False, allowFallback = False):
            """
            Saves the attachment data.

            The name of the file is determined by several factors. The first
            thing that is checked is if you have provided :param customFileName:
            to this function. If you have, that is the name that will be used.
            If no custom name has been provided and :param contentId: is True,
            the file will be saved using the content ID of the attachment. If
            it is not found or :param contentId: is False, the long filename
            will be used. If the long filename is not found, the short one will
            be used. If after all of this a usable filename has not been found,
            """
        # Check if the user has specified a custom filename
        filename = self.getFilename(**kwargs)
        kwargs['customFilename'] = None

        # Someone managed to have a null character here, so let's get rid of that
        filename = prepareFilename(inputToString(filename, self.msg.stringEncoding))

        customPath = os.path.abspath(kwargs.get('customPath', os.getcwdu())).replace('\\', '/')
        customPath += '' if customPath.endswith('/') else '/'
        filename = customPath + filename

        if self.__type == 'data':
            if os.path.exists(filename):
            with open(filename, 'wb') as f:
                f.write(self.__data)
        else:
            self.saveEmbededMessage(**kwargs)
        return filename

    def saveEmbededMessage(self, **kwargs):
        """
        Seperate function from save to allow it to easily be overridden by a
        subclass.
        """
        self.data.save(**kwargs)

    @property
    def cid(self):
        """
        Returns the Content ID of the attachment, if it exists.
        """
        return self._ensureSet('_cid', '__substg1.0_3712')

    contendId = cid

    @property
    def data(self):
        """
        Returns the attachment data.
        """
        return self.__data

    @property
    def longFilename(self):
        """
        Returns the long file name of the attachment, if it exists.
        """
        return self._ensureSet('_longFilename', '__substg1.0_3707')

    @property
    def randomFilename(self):
        """
        Returns the random filename to be used by this attachment.
        """
        try:
            return self.__randomName
        except AttributeError:
            self.regenerateRandomName()
            return self.__randomName

    @property
    def shortFilename(self):
        """
        Returns the short file name of the attachment, if it exists.
        """
        return self._ensureSet('_shortFilename', '__substg1.0_3704')

    @property
    def type(self):
        """
        Returns the (internally used) type of the data.
        """
        return self.__type



class BrokenAttachment(AttachmentBase):
    """
    An attachment that has suffered a fatal error. Will not generate from a
    NotImplementedError exception.
    """
    pass

class UnsupportedAttachment(AttachmentBase):
    """
    An attachment whose type is not currently supported.
    """
    pass
