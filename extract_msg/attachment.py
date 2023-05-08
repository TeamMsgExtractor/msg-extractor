from __future__ import annotations


__all__ = [
    'Attachment',
    'BrokenAttachment',
    'UnsupportedAttachment',
]


import logging
import os
import pathlib
import random
import string
import zipfile

from typing import Optional, TYPE_CHECKING, Union

from . import constants
from .attachment_base import AttachmentBase
from .enums import AttachmentType
from .exceptions import StandardViolationError
from .utils import createZipOpen, inputToString, openMsg, prepareFilename


# Allow for nice type checking.
if TYPE_CHECKING:
    from .msg import MSGFile

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
        :param dir_: the directory inside the msg file where the attachment is
            located.
        """
        super().__init__(msg, dir_)

        if '37050003' not in self.props:
            from .prop import createProp

            logger.warning(f'Attachment method property not found on attachment {dir_}. Code will attempt to guess the type.')
            logger.log(5, self.props)

            # Because this condition is actually kind of a violation of the
            # standard, we are just going to do this in a dumb way. Basically we
            # are going to try to set the attach method *manually* just so I
            # don't have to go and modify the following code.
            if self.exists('__substg1.0_37010102'):
                # Set it as data and call it a day.
                propData = b'\x03\x00\x057\x07\x00\x00\x00\x01\x00\x00\x00\x00\x00\x00\x00'
            elif self.exists('__substg1.0_3701000D'):
                # If it is a folder and we have properties, call it an MSG file.
                if self.exists('__substg1.0_3701000D/__properties_version1.0'):
                    propData = b'\x03\x00\x057\x07\x00\x00\x00\x05\x00\x00\x00\x00\x00\x00\x00'
                else:
                    # Call if custom attachment data.
                    propData = b'\x03\x00\x057\x07\x00\x00\x00\x06\x00\x00\x00\x00\x00\x00\x00'
            else:
                # Can't autodetect it, so throw an error.
                raise StandardViolationError(f'Attachment method missing on attachment {dir_}, and it could not be determined automatically.')

            self.props._propDict['37050003'] = createProp(propData)

        # Get attachment data.
        if self.exists('__substg1.0_37010102'):
            self.__type = AttachmentType.DATA
            self.__data = self._getStream('__substg1.0_37010102')
        elif self.exists('__substg1.0_3701000D'):
            if (self.props['37050003'].value & 0x7) != 0x5:
                raise NotImplementedError(
                    'Current version of extract_msg does not support extraction of containers that are not embedded msg files.')
                # TODO add implementation.
            else:
                self.__prefix = msg.prefixList + [dir_, '__substg1.0_3701000D']
                self.__type = AttachmentType.MSG
                self.__data = openMsg(self.msg.path, prefix = self.__prefix, parentMsg = self.msg, treePath = self.treePath, **self.msg.kwargs)
        elif (self.props['37050003'].value & 0x7) == 0x7:
            # TODO Handling for special attacment type 0x7.
            self.__type = AttachmentType.WEB
            raise NotImplementedError('Attachments of type afByWebReference are not currently supported.')
        else:
            raise TypeError('Unknown attachment type.')

    def getFilename(self, **kwargs) -> str:
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
            customFilename = str(customFilename)
            # First we need to validate it. If there are invalid characters,
            # this will detect it.
            if constants.RE_INVALID_FILENAME_CHARACTERS.search(customFilename):
                raise ValueError('Invalid character found in customFilename. Must not contain any of the following characters: \\/:*?"<>|')
            filename = customFilename
        else:
            # If not...
            # Check if user wants to save the file under the Content-ID.
            if kwargs.get('contentId', False):
                filename = self.cid
            # If filename is None at this point, use long filename as first
            # preference.
            if not filename:
                filename = self.name
            # Otherwise just make something up!
            if not filename:
                return self.randomFilename

        return filename

    def regenerateRandomName(self) -> str:
        """
        Used to regenerate the random filename used if the attachment cannot
        find a usable filename.
        """
        self.__randomName = inputToString('UnknownFilename ' + \
                   ''.join(random.choice(string.ascii_uppercase + string.digits)
                           for _ in range(5)) + '.bin', 'ascii')

    def save(self, **kwargs) -> Optional[Union[str, MSGFile]]:
        """
        Saves the attachment data.

        The name of the file is determined by several factors. The first
        thing that is checked is if you have provided :param customFilename:
        to this function. If you have, that is the name that will be used.
        If no custom name has been provided and :param contentId: is True,
        the file will be saved using the content ID of the attachment. If
        it is not found or :param contentId: is False, the long filename
        will be used. If the long filename is not found, the short one will
        be used. If after all of this a usable filename has not been found, a
        random one will be used (accessible from `Attachment.randomFilename`).
        After the name to use has been determined, it will then be shortened to
        make sure that it is not more than the value of :param maxNameLength:.

        To change the directory that the attachment is saved to, set the value
        of :param customPath: when calling this function. The default save
        directory is the working directory.

        If you want to save the contents into a ZipFile or similar object,
        either pass a path to where you want to create one or pass an instance
        to :param zip:. If :param zip: is an instance, :param customPath: will
        refer to a location inside the zip file.

        :param extractEmbedded: If True, causes the attachment, should it be an
            embedded MSG file, to save as a .msg file instead of calling it's
            save function.
        :param skipEmbedded: If True, skips saving this attachment if it is an
            embedded MSG file.
        """
        # First check if we are skipping embedded messages and stop
        # *immediately* if we are.
        if self.type is AttachmentType.MSG and kwargs.get('skipEmbedded'):
            return None

        # Get the filename to use.
        filename = self.getFilename(**kwargs)

        # Someone managed to have a null character here, so let's get rid of that
        filename = prepareFilename(inputToString(filename, self.msg.stringEncoding))

        # Get the maximum name length.
        maxNameLength = kwargs.get('maxNameLength', 256)

        # Make sure the filename is not longer than it should be.
        if len(filename) > maxNameLength:
            name, ext = os.path.splitext(filename)
            filename = name[:maxNameLength - len(ext)] + ext

        # Check if we are doing a zip file.
        _zip = kwargs.get('zip')

        # ZipFile handling.
        if _zip:
            # If we are doing a zip file, first check that we have been given a path.
            if isinstance(_zip, (str, pathlib.Path)):
                # If we have a path then we use the zip file.
                _zip = zipfile.ZipFile(_zip, 'a', zipfile.ZIP_DEFLATED)
                kwargs['zip'] = _zip
                createdZip = True
            else:
                createdZip = False
            # Path needs to be done in a special way if we are in a zip file.
            customPath = pathlib.Path(kwargs.get('customPath', ''))
            # Set the open command to be that of the zip file.
            _open = createZipOpen(_zip.open)
            # Zip files use w for writing in binary.
            mode = 'w'
        else:
            customPath = pathlib.Path(kwargs.get('customPath', '.')).absolute()
            mode = 'wb'
            _open = open

        fullFilename = customPath / filename

        if self.type is AttachmentType.DATA:
            if _zip:
                name, ext = os.path.splitext(filename)
                nameList = _zip.namelist()
                if str(fullFilename).replace('\\', '/') in nameList:
                    for i in range(2, 100):
                        testName = customPath / f'{name} ({i}){ext}'
                        if str(testName).replace('\\', '/') not in nameList:
                            fullFilename = testName
                            break
                    else:
                        # If we couldn't find one that didn't exist.
                        raise FileExistsError(f'Could not create the specified file because it already exists ("{fullFilename}").')
            else:
                if fullFilename.exists():
                    # Try to split the filename into a name and extention.
                    name, ext = os.path.splitext(filename)
                    # Try to add a number to it so that we can save without overwriting.
                    for i in range(2, 100):
                        testName = customPath / f'{name} ({i}){ext}'
                        if not testName.exists():
                            fullFilename = testName
                            break
                    else:
                        # If we couldn't find one that didn't exist.
                        raise FileExistsError(f'Could not create the specified file because it already exists ("{fullFilename}").')

            with _open(str(fullFilename), mode) as f:
                f.write(self.__data)

            # Close the ZipFile if this function created it.
            if _zip and createdZip:
                _zip.close()

            return str(fullFilename)
        else:
            if kwargs.get('extractEmbedded', False):
                with _open(str(fullFilename), mode) as f:
                    self.data.export(f)
            else:
                self.saveEmbededMessage(**kwargs)

            # Close the ZipFile if this function created it.
            if _zip and createdZip:
                _zip.close()

            return self.msg

    def saveEmbededMessage(self, **kwargs) -> None:
        """
        Seperate function from save to allow it to easily be overridden by a
        subclass.
        """
        self.data.save(**kwargs)

    @property
    def data(self) -> Optional[Union[bytes, MSGFile]]:
        """
        Returns the attachment data.
        """
        return self.__data

    @property
    def randomFilename(self) -> str:
        """
        Returns the random filename to be used by this attachment.
        """
        try:
            return self.__randomName
        except AttributeError:
            self.regenerateRandomName()
            return self.__randomName

    @property
    def type(self) -> AttachmentType:
        """
        Returns the (internally used) type of the data.
        """
        return self.__type



class BrokenAttachment(AttachmentBase):
    """
    An attachment that has suffered a fatal error. Will not generate from a
    NotImplementedError exception.
    """

    @property
    def type(self) -> AttachmentType:
        """
        Returns the (internally used) type of the data.
        """
        return AttachmentType.BROKEN

class UnsupportedAttachment(AttachmentBase):
    """
    An attachment whose type is not currently supported.
    """

    def save(self, **kwargs) -> None:
        """
        Raises a NotImplementedError unless :param skipNotImplemented: is set to
        True. If it is, returns None to signify the attachment was skipped. This
        allows for the easy implementation of the option to skip this type of
        attachment.
        """
        if not kwargs.get('skipNotImplemented', False):
            raise NotImplementedError('Unsupported attachments cannot be saved.')

    @property
    def type(self) -> AttachmentType:
        """
        Returns the (internally used) type of the data.
        """
        return AttachmentType.UNSUPPORTED
