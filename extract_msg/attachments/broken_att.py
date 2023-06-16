from .attachment_base import AttachmentBase
from ..enums import AttachmentType


class BrokenAttachment(AttachmentBase):
    """
    An attachment that has suffered a fatal error. Will not generate from a
    NotImplementedError exception.
    """

    def getFilename(self, **kwargs) -> str:
        raise NotImplementedError('Broken attachments cannot be saved.')

    def save(self, **kwargs):
        raise NotImplementedError('Broken attachments cannot be saved.')

    @property
    def data(self) -> None:
        """
        Broken attachments have no data.
        """
        return None

    @property
    def type(self) -> AttachmentType:
        """
        Returns the (internally used) type of the data.
        """
        return AttachmentType.BROKEN