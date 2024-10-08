from __future__ import annotations


__all__ = [
    'CustomAttachment',
]


import os
import pathlib
import random
import string
import zipfile

from typing import Optional, TYPE_CHECKING

from .. import constants
from .attachment_base import AttachmentBase
from .custom_att_handler import CustomAttachmentHandler, getHandler
from ..enums import AttachmentType, SaveType
from ..utils import createZipOpen, inputToString, prepareFilename


if TYPE_CHECKING:
    from ..msg_classes import MSGFile
    from ..properties import PropertiesStore


_saveDoc = AttachmentBase.save.__doc__


class CustomAttachment(AttachmentBase):
    """
    The attachment entry for custom attachments.
    """

    def __init__(self, msg: MSGFile, dir_: str, propStore: PropertiesStore):
        super().__init__(msg, dir_, propStore)

        self.__customHandler = getHandler(self)
        self.__data = self.__customHandler.data

    def getFilename(self, **kwargs) -> str:
        """
        Returns the filename to use for the attachment.

        If the filename starts with "UnknownFilename" then there is no guarantee
        that the files will have exactly the same filename.

        :param contentId: Use the contentId, if available.
        :param customFilename: A custom name to use for the file.
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
            # Try to get the name from the custom handler.
            filename = self.customHandler.name
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
        self.__randomName = 'UnknownFilename ' + \
                   ''.join(random.choice(string.ascii_uppercase + string.digits)
                           for _ in range(5)) + '.bin'

    def save(self, **kwargs) -> constants.SAVE_TYPE:
        # Immediate check to see if there is anything to save.
        if self.data is None:
            return (SaveType.NONE, None)

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
                # If we are doing a zip file, first check that we have been given a path.
                if isinstance(_zip, (str, pathlib.Path)):
                    # If we have a path then we use the zip file.
                    _zip = zipfile.ZipFile(_zip, 'a', zipfile.ZIP_DEFLATED)
                    kwargs['zip'] = _zip
                    createdZip = True
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

            fullFilename = self._handleFnc(_zip, filename, customPath, kwargs)

            with _open(str(fullFilename), mode) as f:
                f.write(self.__data)

            return (SaveType.FILE, str(fullFilename))
        finally:
            # Close the ZipFile if this function created it.
            if _zip and createdZip:
                _zip.close()


    @property
    def customHandler(self) -> Optional[CustomAttachmentHandler]:
        """
        The instance of the custom handler associated with this attachment, if
        it has one.
        """
        return self.__customHandler

    @property
    def data(self) -> Optional[bytes]:
        """
        The attachment data, if any.

        Returns ``None`` if there is no data to save.
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
        return AttachmentType.CUSTOM
