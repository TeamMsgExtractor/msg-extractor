from __future__ import annotations


__all__ = [
    'SignedAttachment',
]


import email.message
import logging
import os
import pathlib
import weakref
import zipfile

from typing import List, Optional, Type, TYPE_CHECKING, Union

from .. import constants
from ..enums import AttachmentType, SaveType
from ..open_msg import openMsg
from ..utils import createZipOpen, inputToString, makeWeakRef, prepareFilename


# Allow for nice type checking.
if TYPE_CHECKING:
    from ..msg_classes.msg import MSGFile

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class SignedAttachment:
    def __init__(self, msg, data: bytes, name: str, mimetype: str, node: email.message.Message):
        """
        :param msg: The MSGFile instance this attachment is associated with.
        :param data: The bytes that compose this attachment.
        :param name: The reported name of the attachment.
        :param mimetype: The reported mimetype of the attachment.
        :param node: The email Message instance for this node.
        """
        self.__asBytes = data
        self.__name = name
        self.__mimetype = mimetype
        self.__msg = makeWeakRef(msg)
        self.__node = node
        self.__treePath = msg.treePath + [makeWeakRef(self)]

        self.__data = None
        # To add support for embedded MSG files, we are going to completely
        # ignore the mimetype and just do a few simple checks to see if we can
        # use the bytes as am embedded file.
        if data[:8] == b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1':
            try:
                # While we have to pass a lot of data down to the file, we don't
                # pass the prefix and parent MSG data, as it is not an *actual*
                # embedded MSG file. We are just pretending that it is for the
                # external API.
                self.__data = openMsg(data, treePath = self.__treePath, **msg.kwargs)
            except Exception:
                logger.exception('Signed message was an OLE file, but could not be read as an MSG file due to an exception.')

        if self.__data is None:
            self.__data = data

    def _handleFnc(self, _zip, filename, customPath: pathlib.Path, kwargs) -> pathlib.Path:
        """
        "Handle Filename Conflict"

        Internal function for use in determining how to modify the saving path
        when a file with the same name already exists. This is mainly because
        any save function that uses files will need to do this functionality.

        :returns: A ``pathlib.Path`` object to where the file should be saved.
        """
        fullFilename = customPath / filename

        overwriteExisting = kwargs.get('overwriteExisting', False)

        if _zip:
            # If we are writing to a zip file and are not overwriting.
            if not overwriteExisting:
                name, ext = os.path.splitext(filename)
                nameList = _zip.namelist()
                if str(fullFilename).replace('\\', '/') in nameList:
                    for i in range(2, 100):
                        testName = customPath / f'{name} ({i}){ext}'
                        if str(testName).replace('\\', '/') not in nameList:
                            return testName
                    else:
                        # If we couldn't find one that didn't exist.
                        raise FileExistsError(f'Could not create the specified file because it already exists ("{fullFilename}").')
        else:
            if not overwriteExisting and fullFilename.exists():
                # Try to split the filename into a name and extension.
                name, ext = os.path.splitext(filename)
                # Try to add a number to it so that we can save without overwriting.
                for i in range(2, 100):
                    testName = customPath / f'{name} ({i}){ext}'
                    if not testName.exists():
                        return testName
                else:
                    # If we couldn't find one that didn't exist.
                    raise FileExistsError(f'Could not create the specified file because it already exists ("{fullFilename}").')

        return fullFilename

    def save(self, **kwargs) -> constants.SAVE_TYPE:
        """
        Saves the attachment data.

        The name of the file is determined by several factors. The first
        thing that is checked is if you have provided :param customFilename:
        to this function. If you have, that is the name that will be used.
        Otherwise, the name from :attr:`name` will be used. After the name to
        use has been determined, it will then be shortened to make sure that it
        is not more than the value of :param maxNameLength:.

        To change the directory that the attachment is saved to, set the value
        of :param customPath: when calling this function. The default save
        directory is the working directory.

        If you want to save the contents into a ``ZipFile`` or similar object,
        either pass a path to where you want to create one or pass an instance
        to :param zip:. If :param zip: is an instance, :param customPath: will
        refer to a location inside the zip file.
        """
        # First check if we are skipping embedded messages and stop
        # *immediately* if we are.
        if self.type is AttachmentType.SIGNED_EMBEDDED and kwargs.get('skipEmbedded'):
            return (SaveType.NONE, None)

        # If we are running the save function for the MSG file, just let it
        # handle everything.
        if (self.type is AttachmentType.SIGNED_EMBEDDED and
            not kwargs.get('extractEmbedded', False)):
            return self.saveEmbededMessage(**kwargs)

        # Check if the user has specified a custom filename
        filename = self.name

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

        createdZip = True
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
                else:
                    createdZip = False
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

            if self.type is AttachmentType.DATA:
                with _open(str(fullFilename), mode) as f:
                    f.write(self.__data)

                return (SaveType.FILE, str(fullFilename))
            else:
                with _open(str(fullFilename), mode) as f:
                    # We just use the data we were given for this one.
                    f.write(self.__asBytes)
                return (SaveType.FILE, str(fullFilename))
        finally:
            # Close the ZipFile if this function created it.
            if _zip and createdZip:
                _zip.close()

    def saveEmbededMessage(self, **kwargs) -> constants.SAVE_TYPE:
        """
        Seperate function from save to allow it to easily be overridden by a
        subclass.
        """
        return self.data.save(**kwargs)

    @property
    def asBytes(self) -> bytes:
        return self.__asBytes

    @property
    def data(self) -> Union[bytes, MSGFile]:
        """
        The bytes that compose this attachment.
        """
        return self.__data

    @property
    def dataType(self) -> Optional[Type[type]]:
        """
        The class that the data type will use, if it can be retrieved.

        This is a safe way to do type checking on data before knowing if it will
        raise an exception. Returns ``None`` if no data will be returns or if an
        exception will be raised.
        """
        try:
            return None if self.data is None else self.data.__class__
        except Exception:
            # All exceptions that accessing data would cause should be silenced.
            return None

    @property
    def emailMessage(self) -> email.message.Message:
        """
        The email Message instance that is the source for this attachment.
        """
        return self.__node

    @property
    def mimetype(self) -> str:
        """
        The reported mimetype of the attachment.
        """
        return self.__mimetype

    @property
    def msg(self) -> MSGFile:
        """
        The ``MSGFile`` instance this attachment belongs to.

        :raises ReferenceError: The associated ``MSGFile`` instance has been
            garbage collected.
        """
        if (msg := self.__msg()) is None:
            raise ReferenceError('The MSGFile for this Attachment instance has been garbage collected.')
        return msg

    @property
    def name(self) -> str:
        """
        The reported name of this attachment.
        """
        return self.__name

    longFilename = name
    shortFilename = name

    @property
    def treePath(self) -> List[weakref.ReferenceType]:
        """
        A path, as a tuple of instances, needed to get to this instance through
        the MSGFile-Attachment tree.
        """
        return self.__treePath

    @property
    def type(self) -> AttachmentType:
        return AttachmentType.SIGNED if isinstance(self.__data, bytes) else AttachmentType.SIGNED_EMBEDDED
