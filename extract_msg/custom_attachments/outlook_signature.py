import base64
import struct

from typing import List, Optional, Tuple

from . import registerHandler
from .custom_handler import CustomAttachmentHandler
from ..enums import DVAspect


_ST_OLE = struct.Struct('<IIIII')
_ST_MAILSTREAM = struct.Struct('<III')


class OutlookSignature(CustomAttachmentHandler):
    def __init__(self, attachment : 'Attachment'):
        super().__init__(attachment)
        # First we need to get the mailstream.
        stream = attachment._getStream('__substg1.0_3701000D/\x03MailStream')
        if not stream:
            raise ValueError('MailStream could not be found.')
        if len(stream) != 12:
            raise ValueError('MailStream is the wrong length.')
        # Next get the bitmap data.
        self.__data = attachment._getStream('__substg1.0_3701000D/CONTENTS')
        if not self.__data:
            raise ValueError('Bitmap data could not be read for Outlook signature.')
        # Get the OLE data.
        oleStream = attachment._getStream('__substg1.0_3701000D/\x01Ole')
        if not oleStream:
            raise ValueError('OLE stream could not be found.')

        # While I have only seen this stream be one length, it could in theory
        # be more than one length. So long as it is *at least* 20 bytes, we
        # call it valid.
        if len(oleStream) < 20:
            raise ValueError('OLE stream is too short.')

        # Unpack and verify the OLE stream.
        vals = _ST_OLE.unpack(oleStream[:20])
        # Check the version magic.
        if vals[0] != 0x2000001:
            raise ValueError('OLE stream has wrong version magic.')
        # Check the reserved bytes.
        if vals[3] != 0:
            raise ValueError('OLE stream has non-zero reserved int.')

        # Unpack the mailstream and create the HTML tag.
        vals = _ST_MAILSTREAM.unpack(stream)
        self.__dvaspect = DVAspect(vals[0])
        self.__x = vals[1]
        self.__y = vals[2]
        hwStyle = f'height: {self.__x / 100.0:.2f}mm; width: {self.__y / 100.0:.2f}mm;'
        imgData = f'data:image;base64,{base64.b64encode(self.__data)}'
        self.__htmlTag = f'<img src="{imgData}", style="{hwStyle}">'.encode('ascii')

    @classmethod
    def isCorrectHandler(cls, attachment : 'Attachment') -> bool:
        if attachment.clsid != '00000316-0000-0000-C000-000000000046':
            return False

        # Check for the required streams.
        if not attachment.exists('__substg1.0_3701000D/CONTENTS'):
            return False
        if not attachment.exists('__substg1.0_3701000D/\x01Ole'):
            return False
        if not attachment.exists('__substg1.0_3701000D/\x03MailStream'):
            return False

        return True

    def injectHTML(self, html : bytes, renderedList : Optional[List[str]] = None) -> Tuple[bytes, Optional[List[str]]]:
        return html # TODO.
        # Here we want to do the following:
        #     1. Decode the data: We need to know the encoding to be able to
        #        figure out what a "rendered character" is.
        #     2. Break the data up into rendered characters. This will likely
        #        require going one character at a time or using beautiful soup
        #        to first get the body, head, etc. separated.
        #     3. If something isn't going to be rendered, shove that whole
        #        section as part of the next "rendered character" so we can
        #        fully recombine the data later.
        #     4. Find our position and shove our tag onto the front of the next
        #        rendered character so it's in the correct position.
        #     5. Recombine and re-encode the data.
        #     6. Return the new bytes and a list of the rendered characters for
        #        the next function to call to save time.
        #     7. The next time one of these functions is called, it will know to
        #        generate it's own list if the list evaluates as False.

    @property
    def data(self) -> bytes:
        return self.__data

    @property
    def name(self) -> str:
        return self.attachment.shortFilename + '.bmp'




registerHandler(OutlookSignature)
