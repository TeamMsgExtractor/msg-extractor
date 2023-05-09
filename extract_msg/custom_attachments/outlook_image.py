import base64
import struct

from typing import List, Optional, Tuple

from . import registerHandler
from .custom_handler import CustomAttachmentHandler
from .utils import htmlSplitRendered
from ..enums import DVAspect
from ..exceptions import CustomAttachmentError


_ST_OLE = struct.Struct('<IIIII')
_ST_MAILSTREAM = struct.Struct('<III')


class OutlookImage(CustomAttachmentHandler):
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
        self.__y = vals[1]
        self.__x = vals[2]
        hwStyle = f'height: {self.__x / 100.0:.2f}mm; width: {self.__y / 100.0:.2f}mm;'
        imgData = f'data:image;base64,{base64.b64encode(self.__data).decode("ascii")}'
        self.__htmlTag = f'<img src="{imgData}", style="{hwStyle}">'

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
        if not renderedList:
            renderedList = htmlSplitRendered(html)

        rp = self.attachment.renderingPosition

        if rp >= len(renderedList):
            raise CustomAttachmentError(f'Rendering position beyond calculated number of rendered characters (expected less than {len(renderedList)}, got {rp}).')

        renderedList[rp] = self.__htmlTag + renderedList[rp]

        return (''.join(renderedList).encode('utf-8'), renderedList)

    @property
    def data(self) -> bytes:
        return self.__data

    @property
    def name(self) -> str:
        return self.attachment.shortFilename + '.bmp'




registerHandler(OutlookImage)