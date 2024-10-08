from __future__ import annotations


__all__ = [
    'Attachment',
]


import logging
import os
import pathlib
import random
import string
import zipfile

from typing import TYPE_CHECKING

from .. import constants
from .attachment_base import AttachmentBase
from ..enums import AttachmentType, SaveType
from ..utils import createZipOpen, inputToString, prepareFilename
from ..properties import PropertiesStore


# Allow for nice type checking.
if TYPE_CHECKING:
    from ..msg_classes.msg import MSGFile

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class Attachment(AttachmentBase):
    """
    A standard data attachment of an MSG file.
    """

    def __init__(self, msg: MSGFile, dir_: str, propStore: PropertiesStore):
        super().__init__(msg, dir_, propStore)
        self.__data = self.getStream('__substg1.0_37010102')

    def getFilename(self, **kwargs) -> str:
        """
        Returns the filename to use for the attachment.

        :param contentId: Use the contentId, if available.
        :param customFilename: A custom name to use for the file.

        If the filename starts with "UnknownFilename" then there is no guarantee
        that the files will have exactly the same filename.
        """
        filename = None
        customFilename = kwargs.get('customFilename')
        if customFilename:
            customFilename = str(customFilename)
            # First we need to validate it. If there are invalid characters,
            # this will detect it.
            if constants.re.INVALID_FILENAME_CHARS.search(customFilename):
                raise ValueError('Invalid character found in customFilename. Must not contain any of the following characters: \\/:*?"<>|')
            filename = customFilename
        else:
            # If not...
            # Check if user wants to save the file under the Content-ID.
            if kwargs.get('contentId', False):
                filename = self.cid
            # If we are here, try to get the filename however else we can.
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

    def save(self, **kwargs) -> constants.SAVE_TYPE:
        """
        Saves the attachment data.

        The name of the file is determined by several factors. The first
        thing that is checked is if you have provided :param customFilename:
        to this function. If you have, that is the name that will be used.
        If no custom name has been provided and :param contentId: is ``True``,
        the file will be saved using the content ID of the attachment. If
        it is not found or :param contentId: is ``False``, the long filename
        will be used. If the long filename is not found, the short one will
        be used. If after all of this a usable filename has not been found, a
        random one will be used (accessible from :meth:`randomFilename`).
        After the name to use has been determined, it will then be shortened to
        make sure that it is not more than the value of :param maxNameLength:.

        To change the directory that the attachment is saved to, set the value
        of :param customPath: when calling this function. The default save
        directory is the working directory.

        If you want to save the contents into a ``ZipFile`` or similar object,
        either pass a path to where you want to create one or pass an instance
        to :param zip:. If :param zip: is an instance, :param customPath: will
        refer to a location inside the zip file.

        :param extractEmbedded: If ``True``, causes the attachment, should it
            be an embedded MSG file, to save as a .msg file instead of calling
            it's save function.
        :param skipEmbedded: If ``True``, skips saving this attachment if it is
            an embedded MSG file.
        """
        # Get the filename to use.
        filename = self.getFilename(**kwargs)

        # Someone managed to have a null character here, so let's get rid of
        # that
        filename = prepareFilename(inputToString(filename, self.msg.stringEncoding))

        # Get the maximum name length.
        maxNameLength = kwargs.get('maxNameLength', 256)

        # Make sure the filename is not longer than it should be.
        if len(filename) > maxNameLength:
            name, ext = os.path.splitext(filename)
            filename = name[:maxNameLength - len(ext)] + ext

        # Check if we are doing a zip file.
        _zip = kwargs.get('zip')

        createdZip = False
        try:
            # ZipFile handling.
            if _zip:
                # If we are doing a zip file, first check that we have been
                # given a path.
                if isinstance(_zip, (str, pathlib.Path)):
                    # If we have a path then we use the zip file.
                    _zip = zipfile.ZipFile(_zip, 'a', zipfile.ZIP_DEFLATED)
                    kwargs['zip'] = _zip
                    createdZip = True
                # Path needs to be done in a special way if we are in a zip
                # file.
                customPath = pathlib.Path(kwargs.get('customPath', ''))
                # Set the open command to be that of the zip file.
                _open = createZipOpen(_zip.open)
                # Zip files use w for writing in binary.
                mode = 'w'
            else:
                customPath = pathlib.Path(kwargs.get('customPath', '.')).absolute()
                mode = 'wb'
                _open = open

            fullFilename = self._handleFnc(_zip, filename, customPath, kwargs)

            with _open(str(fullFilename), mode) as f:
                f.write(self.__data)

            return (SaveType.FILE, str(fullFilename))

        finally:
            # Close the ZipFile if this function created it.
            if _zip and createdZip:
                _zip.close()

    @property
    def data(self) -> bytes:
        """
        The bytes making up the attachment data.
        """
        return self.__data

    @property
    def randomFilename(self) -> str:
        """
        The random filename to be used by this attachment.
        """
        try:
            return self.__randomName
        except AttributeError:
            self.regenerateRandomName()
            return self.__randomName

    @property
    def type(self) -> AttachmentType:
        return AttachmentType.DATA
