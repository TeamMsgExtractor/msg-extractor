from __future__ import annotations


__all__ = [
    'EmbeddedMsgAttachment',
]


import os
import pathlib
import zipfile

from typing import TYPE_CHECKING

from .. import constants
from .attachment_base import AttachmentBase
from ..enums import AttachmentType, SaveType
from ..open_msg import openMsg
from ..utils import createZipOpen, prepareFilename


if TYPE_CHECKING:
    from ..msg_classes import MSGFile
    from ..properties import PropertiesStore


_saveDoc = AttachmentBase.save.__doc__


class EmbeddedMsgAttachment(AttachmentBase):
    """
    The attachment entry for an Embedded MSG file.
    """

    def __init__(self, msg: MSGFile, dir_: str, propStore: PropertiesStore):
        super().__init__(msg, dir_, propStore)
        self.__prefix = msg.prefixList + [dir_, '__substg1.0_3701000D']
        self.__data = openMsg(self.msg.path, prefix = self.__prefix, parentMsg = self.msg, treePath = self.treePath, **self.msg.kwargs)

    def getFilename(self, **kwargs) -> str:
        """
        Returns the filename to use for the attachment.

        :param contentId: Use the contentId, if available.
        :param customFilename: A custom name to use for the file.

        If the filename starts with "UnknownFilename" then there is no guarantee
        that the files will have exactly the same filename.
        """
        customFilename = kwargs.get('customFilename')
        if customFilename:
            customFilename = str(customFilename)
            # First we need to validate it. If there are invalid characters,
            # this will detect it.
            if constants.re.INVALID_FILENAME_CHARS.search(customFilename):
                raise ValueError('Invalid character found in customFilename. Must not contain any of the following characters: \\/:*?"<>|')
            return customFilename
        else:
            return self.name

    def save(self, **kwargs) -> constants.SAVE_TYPE:
        # First check if we are skipping embedded messages and stop
        # *immediately* if we are.
        if kwargs.get('skipEmbedded'):
            return (SaveType.NONE, None)

        # We only need to handle things if we are saving as bytes.
        if kwargs.get('extractEmbedded', False):
            # Get the filename to use.
            filename = self.getFilename(**kwargs)

            # Someone managed to have a null character here, so let's get rid of
            # that
            filename = prepareFilename(filename)

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
                    self.data.export(f)

                return (SaveType.FILE, str(fullFilename))
            finally:
                # Close the ZipFile if this function created it.
                if _zip and createdZip:
                    _zip.close()
        else:
            # If we are letting the MSG file create stuff, just let it handle
            # everything.
            return self.data.save(**kwargs)

    save.__doc__ = _saveDoc

    @property
    def data(self) -> MSGFile:
        """
        Returns the attachment data.
        """
        return self.__data

    @property
    def type(self) -> AttachmentType:
        return AttachmentType.MSG