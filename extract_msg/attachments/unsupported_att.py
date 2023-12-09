__all__ = [
    'UnsupportedAttachment',
]


from .. import constants
from .attachment_base import AttachmentBase
from ..enums import AttachmentType, SaveType


class UnsupportedAttachment(AttachmentBase):
    """
    An attachment whose type is not currently supported.
    """

    def getFilename(self, **_) -> str:
        raise NotImplementedError('Unsupported attachments cannot be saved.')

    def save(self, **kwargs) -> constants.SAVE_TYPE:
        """
        Raises a NotImplementedError unless :param skipNotImplemented: is set to
        ``True``.

        If it is, returns a value that indicates no data was saved.
        """
        if kwargs.get('skipNotImplemented', False):
            return (SaveType.NONE, None)

        raise NotImplementedError('Unsupported attachments cannot be saved.')

    @property
    def data(self) -> None:
        """
        Broken attachments have no data.
        """
        return None

    @property
    def type(self) -> AttachmentType:
        return AttachmentType.UNSUPPORTED