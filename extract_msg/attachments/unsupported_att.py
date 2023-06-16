__all__ = [
    'UnsupportedAttachment',
]


from .attachment_base import AttachmentBase
from ..enums import AttachmentType


class UnsupportedAttachment(AttachmentBase):
    """
    An attachment whose type is not currently supported.
    """

    def getFilename(self, **kwargs) -> str:
        raise NotImplementedError('Unsupported attachments cannot be saved.')

    def save(self, **kwargs) -> None:
        """
        Raises a NotImplementedError unless :param skipNotImplemented: is set to
        True. If it is, returns None to signify the attachment was skipped. This
        allows for the easy implementation of the option to skip this type of
        attachment.
        """
        if not kwargs.get('skipNotImplemented', False):
            raise NotImplementedError('Unsupported attachments cannot be saved.')

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
        return AttachmentType.UNSUPPORTED