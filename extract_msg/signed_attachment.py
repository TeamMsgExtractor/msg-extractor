import email
import logging
import os
import pathlib
import zipfile

from .enums import AttachmentType
from .utils import createZipOpen, inputToString, prepareFilename


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class SignedAttachment:
    def __init__(self, msg, data : bytes, name : str, mimetype : str, node : email.message.Message):
        """
        :param msg: The msg file this attachment is associated with.
        :param data: The bytes that compose this attachment.
        :param name: The reported name of the attachment.
        :param mimetype: The reported mimetype of the attachment.
        :param node: The email Message instance for this node.
        """
        self.__data = data
        self.__name = name
        self.__mimetype = mimetype
        self.__msg = msg
        self.__node = node

    def save(self, **kwargs):
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
        """
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


        if _zip:
            name, ext = os.path.splitext(filename)
            nameList = _zip.namelist()
            if fullFilename in nameList:
                for i in range(2, 100):
                    testName = customPath / f'{name} ({i}){ext}'
                    if testName not in nameList:
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

        return fullFilename

    @property
    def data(self) -> bytes:
        """
        The bytes that compose this attachment.
        """
        return self.__data

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
    def msg(self) -> "MSGFile":
        """
        The MSGFile instance this attachment belongs to.
        """
        return self.__msg

    @property
    def name(self) -> str:
        """
        The reported name of this attachment.
        """
        return self.__name

    @property
    def type(self) -> AttachmentType:
        """
        The AttachmentType.
        """
        return AttachmentType.SIGNED
