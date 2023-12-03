from __future__ import annotations


__all__ = [
    'OutlookImage',
]


import struct

from typing import Optional, TYPE_CHECKING

from . import registerHandler
from .custom_handler import CustomAttachmentHandler
from ...enums import DVAspect, InsecureFeatures
from ...exceptions import DependencyError, SecurityError


if TYPE_CHECKING:
    from ..attachment_base import AttachmentBase

_ST_OLE = struct.Struct('<IIIII')
_ST_MAILSTREAM = struct.Struct('<III')


class OutlookImageDIB(CustomAttachmentHandler):
    """
    Custom handler for a special attachment type, a Device Independent Bitmap
    stored in a way special to Outlook.
    """

    def __init__(self, attachment: AttachmentBase):
        super().__init__(attachment)
        # First we need to get the mailstream.
        stream = self.getStream('\x03MailStream')
        if not stream:
            raise ValueError('MailStream could not be found.')
        if len(stream) != 12:
            raise ValueError('MailStream is the wrong length.')
        # Next get the bitmap data.
        self.__data = self.getStream('CONTENTS')
        if not self.__data:
            raise ValueError('Bitmap data could not be read for Outlook signature.')
        # Get the OLE data.
        oleStream = self.getStream('\x01Ole')
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
        # Convert to twips for RTF.
        self.__xtwips = int(round(self.__x / 1.7639))
        self.__ytwips = int(round(self.__y / 1.7639))

    @classmethod
    def isCorrectHandler(cls, attachment: AttachmentBase) -> bool:
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

    def generateRtf(self) -> Optional[bytes]:
        """
        Generates the RTF to inject in place of the \\objattph tag.

        If this function should do nothing, returns ``None``.

        :raises DependencyError: PIL or Pillow could not be found.
        """
        if InsecureFeatures.PIL_IMAGE_PARSING not in self.attachment.msg.insecureFeatures:
            raise SecurityError('Generating the RTF for a custom attachment requires the insecure feature PIL_IMAGE_PARSING.')

        try:
            import PIL.Image
        except ImportError:
            raise DependencyError('PIL or Pillow is required for inserting an Outlook Image into the body.')

        # First, convert the bitmap into a PNG so we can insert it into the
        # body.
        import io

        # Note, use self.data instead of self.__data to allow support for
        # extensions.
        with PIL.Image.open(io.BytesIO(self.data)) as img:
            out = io.BytesIO()
            img.save(out, 'PNG')

        hexData = out.getvalue().hex()

        inject = '{\\*\\shppict\n{\\pict\\picscalex100\\picscaley100'
        inject += f'\\picw{img.width}\\pich{img.height}'
        inject += f'\\picwgoal{self.__xtwips}\\pichgoal{self.__ytwips}\n'
        inject += '\\pngblip ' + hexData + '}}'

        return inject.encode()

    @property
    def data(self) -> bytes:
        return self.__data

    @property
    def name(self) -> str:
        # Try to get the name from the attachment. If that fails, name it based
        # on the number.
        if not (name := self.attachment.name):
            name = f'attachment {int(self.attachment.dir[-8:], 16)}'
        return name + '.bmp'

    @property
    def obj(self) -> bytes:
        return self.data



registerHandler(OutlookImageDIB)
