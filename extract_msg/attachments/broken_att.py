__all__ = [
    'BrokenAttachment',
]


from .. import constants
from .attachment_base import AttachmentBase
from ..enums import AttachmentType, SaveType


class BrokenAttachment(AttachmentBase):
    """
    An attachment that has suffered a fatal error.

    Will not generate from a NotImplementedError exception.
    """

    def getFilename(self, **_) -> str:
        raise NotImplementedError('Broken attachments cannot be saved.')

    def save(self, **kwargs) -> constants.SAVE_TYPE:
        """
        Raises a NotImplementedError unless :param skipNotImplemented: is set to
        True.

        If it is, returns a value that indicates no data was saved.
        """
        if kwargs.get('skipNotImplemented', False):
            return (SaveType.NONE, None)

        raise NotImplementedError('Broken attachments cannot be saved.')

    @property
    def data(self) -> None:
        """
        Broken attachments have no data.
        """
        return None

    @property
    def type(self) -> AttachmentType:
        return AttachmentType.BROKEN