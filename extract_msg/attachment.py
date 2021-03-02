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
        if self.Exists('__substg1.0_37010102'):
            self.__type = 'data'
            self.__data = self._getStream('__substg1.0_37010102')
        elif self.Exists('__substg1.0_3701000D'):
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

    def save(self, contentId = False, json = False, useFileName = False, raw = False, customPath = None,
             customFilename = None):#, html = False, rtf = False, allowFallback = False):
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

        # Someone managed to have a null character here, so let's get rid of that
        filename = prepareFilename(inputToString(filename, self.msg.stringEncoding))

        if customPath is not None and customPath != '':
            if customPath[-1] != '/' or customPath[-1] != '\\':
                customPath += '/'
            filename = customPath + filename

        if self.__type == "data":
            with open(filename, 'wb') as f:
                f.write(self.__data)
        else:
            self.saveEmbededMessage(contentId, json, useFileName, raw, customPath, customFilename)#, html, rtf, allowFallback)
        return filename

    def saveEmbededMessage(self, contentId = False, json = False, useFileName = False, raw = False, customPath = None,
                           customFilename = None):#, html = False, rtf = False, allowFallback = False):
        """
        Seperate function from save to allow it to
        easily be overridden by a subclass.
        """
        self.data.save(json, useFileName, raw, contentId, customPath, customFilename)#, html, rtf, allowFallback)

    @property
    def cid(self):
        """
        Returns the Content ID of the attachment, if it exists.
        """
        return self._ensureSet('_cid', '__substg1.0_3712')

    contend_id = cid

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
