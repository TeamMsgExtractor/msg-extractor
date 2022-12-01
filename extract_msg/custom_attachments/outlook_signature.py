import struct

from .custom_handler import CustomAttachmentHandler
from ..enums import DVAspect


_MAILSTREAM_STRUCT = struct.Struct('<III')


class OutlookSignature(CustomAttachmentHandler):
    def __init__(self, attachment : 'Attachment'):
        super().__init__(attachment)
        # First we need to get the mailstream.
        stream = attachment._getStream('__substg1.0_3701000D/\x03MailStream')
        if not stream:
            raise ValueError('MailStream could not be found.')
        # Next get the bitmap data.
        self.__data = attachment._getStream('__substg1.0_3701000D/CONTENTS')
        if not self.__data:
            raise ValueError('Bitmap data could not be read for Outlook signature.')
        # Get the OLE data.
        # TODO.

        # Unpack the mailstream and create the HTML tag.
        vals = _MAILSTREAM_STRUCT.unpack(stream)
        self.__dvaspect = DVAspect(vals[0])
        self.__x = vals[1]
        self.__y = vals[2]
        hwStyle = f'height: {self.__x / 100.0:.2f}; width: {self.__y / 100.0:.2f};'
        imgData = f'data:image;base64,{base64.b64encode(self.__data)}';
        self.__htmlTag = f'<img src="{imgData}", style="{hwStyle}">'.encode('ascii')
